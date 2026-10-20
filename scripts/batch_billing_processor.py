import os
import re
import glob
import json
import openpyxl
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# --------------------------------------------------------------------------
# 1. 텍스트 정제 및 5대 표준 직종 매핑 헬퍼 함수
# --------------------------------------------------------------------------
def clean_vendor_name(name):
    """(주), ㈜, 주식회사 등의 수식어를 제거하고 순수 상호명만 반환"""
    if not name:
        return "기타"
    cleaned = str(name).strip()
    cleaned = re.sub(r'[\(（]?[주株式]+[\)）]?', '', cleaned)
    cleaned = cleaned.replace('㈜', '').replace('주식회사', '').strip()
    return cleaned

def clean_number(val):
    """문자열, 콤마, None 등을 안전한 float 숫자로 변환"""
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip().replace(',', '').replace('₩', '').replace('원', '')
    if val_str.startswith('(') and val_str.endswith(')'):
        val_str = '-' + val_str[1:-1]
    match = re.search(r'[-+]?\d*\.?\d+', val_str)
    return float(match.group()) if match else 0.0

def map_standard_craft(job_title):
    """
    현장 직종명을 회사의 5대 표준 직종으로 통합 매핑
    - 공사보조, 안전보조, 관리보조 등은 '관리자'로 통합
    """
    title = str(job_title).strip()
    if '배관' in title:
        return '배관공'
    elif '용접' in title:
        return '용접사'
    elif '도비' in title:
        return '도비공'
    elif '보통' in title:
        return '보통인부'
    elif any(k in title for k in ['관리', '보조', '공사보조', '안전보조', '관리보조']):
        return '관리자'
    return '기타'

# --------------------------------------------------------------------------
# 2. 개별 기성 엑셀 파일 파싱 엔진 (노무비명세서/근로내역 우선 탐색)
# --------------------------------------------------------------------------
def parse_single_excel(file_path):
    wb = openpyxl.load_workbook(file_path, data_only=True)
    file_name = os.path.basename(file_path)

    # (1) 기본 정보: 표지 및 갑지 탐색
    ws_cover = wb['용역청구서 표지'] if '용역청구서 표지' in wb.sheetnames else None
    ws_gap = wb['용역 청구서(갑지)'] if '용역 청구서(갑지)' in wb.sheetnames else wb.active

    # 프로젝트명
    pjt_name = ws_cover['B7'].value if (ws_cover and ws_cover['B7'].value) else "P4 Ph4 GAS 1공구 마감공사"

    # 협력사명 ((주), ㈜ 제외)
    raw_vendor = ws_cover['C15'].value if (ws_cover and ws_cover['C15'].value) else ""
    if not raw_vendor:
        for v in ['대성이엔지', '태성하이텍', '리도시스템']:
            if v in file_name:
                raw_vendor = v
                break
    vendor_name = clean_vendor_name(raw_vendor)

    # 청구월 (YYYY-MM 형식 추출)
    raw_month = ws_cover['C14'].value if (ws_cover and ws_cover['C14'].value) else file_name
    month_match = re.search(r'(\d{2,4})년?\s*(\d{1,2})월?', str(raw_month))
    if month_match:
        y = month_match.group(1)
        y = f"20{y}" if len(y) == 2 else y
        m = str(month_match.group(2)).zfill(2)
        claim_month = f"{y}-{m}"
    else:
        claim_month = "2026-09"

    # (2) 금회 위탁비용 (갑지 D13: 금회 청구금액)
    claim_cost = clean_number(ws_gap['D13'].value)

    # (3) 경비 내역 (숙소임차료, 숙소관리비, 식대비용)
    exp_rent = clean_number(wb['숙소임차료']['I2'].value) if '숙소임차료' in wb.sheetnames else 0.0
    exp_mgt = clean_number(wb['숙소관리비']['I2'].value) if '숙소관리비' in wb.sheetnames else 0.0
    exp_dorm = exp_rent + exp_mgt
    exp_meals = clean_number(wb['식대비용']['I2'].value) if '식대비용' in wb.sheetnames else 0.0
    exp_total = exp_dorm + exp_meals

    # (4) 시공 물량 (물량 내역서 시트의 금회 / 누계 DIA 합계)
    dia_now = 0.0
    dia_cum = 0.0
    if '물량 내역서' in wb.sheetnames:
        ws_qty = wb['물량 내역서']
        dia_now = clean_number(ws_qty['J50'].value)
        dia_cum = clean_number(ws_qty['L50'].value)

    # (5) 기술인 인건비: '내역서(을지)'는 완전히 제외하고 '근로내역'/'노무비명세서' 시트만 탐색
    craft_data = {
        '배관공': {'공수': 0.0, '금액': 0.0},
        '용접사': {'공수': 0.0, '금액': 0.0},
        '관리자': {'공수': 0.0, '금액': 0.0},
        '보통인부': {'공수': 0.0, '금액': 0.0},
        '도비공': {'공수': 0.0, '금액': 0.0}
    }

    # 근로내역/노무비명세서 시트 찾기
    work_sheet = None
    for s in wb.sheetnames:
        if any(k in s for k in ['근로내역', '노무비명세서', '노무비', '출력일보']):
            work_sheet = wb[s]
            break

    if work_sheet:
        header_row = 5
        job_col = 4
        md_col = 39
        amt_col = 41

        # 동적 헤더 탐색
        for r in range(1, 10):
            row_vals = [str(work_sheet.cell(r, c).value or '') for c in range(1, 45)]
            if any('직종' in v for v in row_vals) and any('공수' in v for v in row_vals):
                header_row = r
                for c, val in enumerate(row_vals, 1):
                    if '직종' in val: job_col = c
                    elif '공수' in val: md_col = c
                    elif any(k in val for k in ['총액', '지급액', '합계', '금액']): amt_col = c
                break

        # 개별 작업자 출력 데이터 집계
        for r in range(header_row + 1, work_sheet.max_row + 1):
            name_val = work_sheet.cell(r, 2).value
            job_val = work_sheet.cell(r, job_col).value
            md_val = clean_number(work_sheet.cell(r, md_col).value)
            amt_val = clean_number(work_sheet.cell(r, amt_col).value)

            if not name_val or any(k in str(name_val) for k in ['합계', '소계']):
                continue

            std_job = map_standard_craft(job_val)
            if std_job in craft_data and md_val > 0:
                craft_data[std_job]['공수'] += md_val
                craft_data[std_job]['금액'] += amt_val

    total_md = sum(v['공수'] for v in craft_data.values())
    total_labor = sum(v['금액'] for v in craft_data.values())

    return {
        "파일명": file_name,
        "PJT명": pjt_name,
        "협력업체명": vendor_name,
        "청구월": claim_month,
        "금회_위탁비용": claim_cost,
        "경비_숙소임차료": exp_rent,
        "경비_숙소관리비": exp_mgt,
        "경비_숙소소계": exp_dorm,
        "경비_식대": exp_meals,
        "경비_총합계": exp_total,
        "금회_DIA_수량": dia_now,
        "누계_DIA_수량": dia_cum,
        "배관공_공수": craft_data['배관공']['공수'],
        "배관공_금액": craft_data['배관공']['금액'],
        "용접사_공수": craft_data['용접사']['공수'],
        "용접사_금액": craft_data['용접사']['금액'],
        "관리자_공수": craft_data['관리자']['공수'],
        "관리자_금액": craft_data['관리자']['금액'],
        "보통인부_공수": craft_data['보통인부']['공수'],
        "보통인부_금액": craft_data['보통인부']['금액'],
        "도비공_공수": craft_data['도비공']['공수'],
        "도비공_금액": craft_data['도비공']['금액'],
        "실투입_총공수": total_md,
        "실투입_총노무비": total_labor
    }

# --------------------------------------------------------------------------
# 3. data/ 폴더 일괄 파싱 및 구글 스프레드시트 적재 메인 실행부
# --------------------------------------------------------------------------
def main():
    excel_files = glob.glob('data/*.xlsx') + glob.glob('data/*.xls')
    if not excel_files:
        print("[INFO] data/ 폴더에 처리할 기성 엑셀 파일이 없습니다.")
        return

    print(f"[INFO] 총 {len(excel_files)}개의 기성 파일을 처리합니다.")
    records = []
    for f in excel_files:
        try:
            rec = parse_single_excel(f)
            records.append(rec)
            print(f" [완료] {rec['협력업체명']} ({rec['청구월']}): 위탁 {rec['금회_위탁비용']:,}원 | 공수 {rec['실투입_총공수']} M/D | 노무비 {rec['실투입_총노무비']:,}원")
        except Exception as e:
            print(f" [오류] {f} 파싱 중 예외 발생: {e}")

    if not records:
        print("[WARN] 파싱된 유효 데이터가 없습니다.")
        return

    df = pd.DataFrame(records)
    # 청구월 오름차순, 협력업체명 정렬
    df = df.sort_values(by=['청구월', '협력업체명'], ascending=[True, True])

    # 구글 서비스 계정 인증
    service_account_json = os.environ.get('GCP_SERVICE_ACCOUNT_KEY')
    spreadsheet_id = os.environ.get('SPREADSHEET_ID')

    if not service_account_json or not spreadsheet_id:
        print("[ERROR] GCP_SERVICE_ACCOUNT_KEY 또는 SPREADSHEET_ID 환경변수가 설정되지 않았습니다.")
        # 로컬 테스트용 CSV 출력
        df.to_csv("parsed_billing_summary.csv", index=False, encoding='utf-8-sig')
        print("[INFO] parsed_billing_summary.csv 파일로 로컬 저장 완료.")
        return

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds_dict = json.loads(service_account_json)
    credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    gc = gspread.authorize(credentials)

    # 스프레드시트 열기 및 덮어쓰기 (기존 중복 방지 및 최신 동기화)
    sh = gc.open_by_key(spreadsheet_id)
    ws = sh.get_worksheet(0)

    # DataFrame 데이터를 2차원 리스트로 변환 (헤더 포함)
    data_to_write = [df.columns.values.tolist()] + df.fillna(0).values.tolist()
    
    ws.clear()
    ws.update('A1', data_to_write)
    print(f"[SUCCESS] 구글 시트에 총 {len(df)}개 레코드 덮어쓰기 적재 성공!")

if __name__ == '__main__':
    main()
