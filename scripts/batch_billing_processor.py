import os
import re
import glob
import json
import sys
import openpyxl
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# --------------------------------------------------------------------------
# 0. 파일 포맷(.xlsx / .xls) 통합 로더
# --------------------------------------------------------------------------
def load_any_workbook(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    
    if ext == '.xlsx':
        return openpyxl.load_workbook(file_path, data_only=True)
    
    elif ext == '.xls':
        try:
            import xlrd
        except ImportError:
            raise ImportError("구버전 .xls 파일을 읽으려면 'xlrd' 라이브러리가 필요합니다. pip install xlrd를 실행해 주세요.")
        
        rb = xlrd.open_workbook(file_path, formatting_info=False)
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        for sheet_idx in range(rb.nsheets):
            r_sheet = rb.sheet_by_index(sheet_idx)
            w_sheet = wb.create_sheet(title=r_sheet.name)

            for row_idx in range(r_sheet.nrows):
                row_vals = []
                for col_idx in range(r_sheet.ncols):
                    cell_val = r_sheet.cell_value(row_idx, col_idx)
                    row_vals.append(cell_val)
                w_sheet.append(row_vals)

        return wb
    else:
        raise ValueError(f"지원하지 않는 엑셀 파일 형식입니다: {ext}")

# --------------------------------------------------------------------------
# 1. 문자열 정제 및 직종 매핑 헬퍼 함수
# --------------------------------------------------------------------------
def clean_vendor_name(name):
    """(주), 주식회사, 공종 표기(FAB, SUP 등) 제거"""
    if not name:
        return ""
    cleaned = str(name).strip()
    cleaned = re.sub(r'[\(（]?[주株式]+[\)）]?', '', cleaned)
    cleaned = cleaned.replace('㈜', '').replace('주식회사', '').strip()
    cleaned = re.sub(r'[\(（]?(FAB|SUP|gas|pcw|pwc)[\)）]?', '', cleaned, flags=re.I).strip()
    return cleaned

def clean_number(val):
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

def find_sheet_by_keywords(wb, keywords):
    for name in wb.sheetnames:
        norm_name = re.sub(r'\s+', '', name)
        if any(re.sub(r'\s+', '', kw) in norm_name for kw in keywords):
            return wb[name]
    return None

# --------------------------------------------------------------------------
# 2. '집계장' 파싱 엔진 (물량 및 경비 추출)
# --------------------------------------------------------------------------
def parse_jibgye_sheet(sheet):
    result = {
        "dia_now": 0.0,
        "dia_cum": 0.0,
        "exp_dorm": 0.0,
        "exp_meals": 0.0
    }
    if not sheet:
        return result

    item_col = 2
    unit_col = 3
    qty_col = 4
    now_amt_col = 7
    cum_amt_col = 8

    for r in range(1, 7):
        row_vals = [str(sheet.cell(r, c).value or '').strip() for c in range(1, 15)]
        for c, val in enumerate(row_vals, 1):
            clean_v = re.sub(r'\s+', '', val)
            if '내역' in clean_v: item_col = c
            elif '단위' in clean_v: unit_col = c
            elif '수량' in clean_v: qty_col = c
            elif '금회' in clean_v and any(k in clean_v for k in ['기성', '금액']): now_amt_col = c
            elif '누계' in clean_v and any(k in clean_v for k in ['기성', '금액']): cum_amt_col = c

    for r in range(1, sheet.max_row + 1):
        item_val = str(sheet.cell(r, item_col).value or '').strip()
        unit_val = str(sheet.cell(r, unit_col).value or '').strip()
        row_str = " ".join([str(sheet.cell(r, c).value or '') for c in range(1, 10)])
        clean_row = re.sub(r'\s+', '', row_str)

        if 'WELDING' in item_val.upper() or 'D/I' in unit_val.upper() or 'WELDINGPOINT' in clean_row:
            dia_val = clean_number(sheet.cell(r, qty_col).value)
            if dia_val > 0:
                result["dia_now"] = dia_val
                result["dia_cum"] = dia_val

        if any(k in clean_row for k in ['임대료+관리비', '숙소', '임차료']):
            amt = clean_number(sheet.cell(r, now_amt_col).value)
            if amt > 0:
                result["exp_dorm"] = amt

        if '식대' in clean_row and '인원' not in clean_row:
            amt = clean_number(sheet.cell(r, now_amt_col).value)
            if amt > 0:
                result["exp_meals"] = amt

    return result

# --------------------------------------------------------------------------
# 3. 개별 엑셀 파일 동적 파싱 엔진
# --------------------------------------------------------------------------
def parse_single_excel(file_path):
    wb = load_any_workbook(file_path)
    file_name = os.path.basename(file_path)

    ws_cover = find_sheet_by_keywords(wb, ['표지', '용역청구서표지'])
    ws_gap = find_sheet_by_keywords(wb, ['갑지', '용역청구서(갑지)']) or wb.active

    # (1) PJT_NO 엄격 탐색 (S25-004, S26-001 등)
    pjt_no = ""
    if ws_cover:
        for r in range(1, 15):
            for c in range(1, 8):
                val = str(ws_cover.cell(r, c).value or '').strip()
                s_match = re.search(r'\b(S\d{2}-\d{3,4})\b', val, re.I)
                if s_match:
                    pjt_no = s_match.group(1).upper()
                    break
            if pjt_no: break

    if not pjt_no:
        s_match = re.search(r'[\(（]?(S\d{2}-\d{3,4})[\)）]?', file_name, re.I)
        if s_match:
            pjt_no = s_match.group(1).upper()

    if not pjt_no and ws_cover and ws_cover['B8'].value:
        cand_b8 = str(ws_cover['B8'].value).strip()
        if cand_b8 and len(cand_b8) >= 4 and not any(k in cand_b8 for k in ['공사', '청구서', '갑지', '공구']):
            pjt_no = cand_b8

    if not pjt_no:
        ph_match = re.search(r'(P\d+\s*Ph\d+)', file_name, re.I)
        pjt_no = ph_match.group(1).strip() if ph_match else "기타 PJT"

    # (2) PJT명 동적 추출
    pjt_name = ""
    if ws_cover and ws_cover['B7'].value:
        pjt_name = str(ws_cover['B7'].value).strip()

    if not pjt_name and ws_cover:
        for r in range(1, 15):
            for c in range(1, 6):
                val = str(ws_cover.cell(r, c).value or '').strip()
                if any(k in val for k in ['공구', '공사', 'PJT', 'PROJECT']) and len(val) >= 4 and not re.search(r'^S\d{2}-', val):
                    pjt_name = val
                    break
            if pjt_name: break

    if not pjt_name:
        pjt_match = re.search(r'\]_?([^_]+)_[^_]+', file_name)
        if pjt_match:
            candidate = pjt_match.group(1).strip()
            pjt_name = re.sub(r'용역청구서|청구서|기성', '', candidate).strip()
        if not pjt_name:
            pjt_name = "일반공사"

    # (3) 협력업체명 100% 동적 추출 (사전 하드코딩 완전 배제)
    raw_vendor = ""
    
    # 1. 엑셀 표지/갑지 내 라벨 기반 검색 (우측 4개 열 및 병합 셀 완벽 스캔)
    search_sheets = [s for s in [ws_cover, ws_gap] if s]
    for ws_target in search_sheets:
        for r in range(5, 25):
            for c in range(1, 6):
                lbl = str(ws_target.cell(r, c).value or '').strip()
                if any(k in lbl for k in ['상호', '업체명', '상 호', '협력업체', '수급인']):
                    for offset in range(1, 5):
                        val = clean_vendor_name(str(ws_target.cell(r, c + offset).value or ''))
                        if val and len(val) >= 2 and not any(k in val for k in ['상호', '성명', '대표', '주식회사']):
                            raw_vendor = val
                            break
                    if raw_vendor: break
            if raw_vendor: break
        if raw_vendor: break

    # 2. 표지에 상호가 없을 경우 파일명 괄호 파싱
    if not raw_vendor:
        clean_fn = re.sub(r'\.(xlsx|xls)$', '', file_name, flags=re.I)
        paren_matches = re.findall(r'[\(（]([^\)）]+)[\)）]', clean_fn)
        for cand in reversed(paren_matches):
            cand_clean = clean_vendor_name(cand)
            if cand_clean and len(cand_clean) >= 2 and not any(k in cand_clean.upper() for k in ['FAB', 'SUP', '공구', '회차', '월', '검토', 'R1', 'R2', 'S25', 'S26']):
                raw_vendor = cand_clean
                break

    # 3. 파일명 언더바 분할 파싱
    if not raw_vendor:
        clean_fn = re.sub(r'\.(xlsx|xls)$', '', file_name, flags=re.I)
        parts = clean_fn.split('_')
        invalid_keywords = [
            'R1', 'R2', 'REV', 'FINAL', '청구서', '갑지', '을지', '최종', '기성', '공구',
            '검토', '내역', '투입비', '회차', '정산', '마감', 'PJT', '공사', 'FAB', 'SUP', '일반공사'
        ]
        for p in reversed(parts):
            p_cand = clean_vendor_name(p)
            if len(re.findall(r'\d', p_cand)) >= 2 or any(k in p_cand.upper() for k in invalid_keywords):
                continue
            if len(p_cand) >= 2:
                raw_vendor = p_cand
                break

    vendor_name = clean_vendor_name(raw_vendor) if raw_vendor else "기타"

    # (4) 청구월 (YYYY-MM)
    raw_month = ws_cover['C14'].value if (ws_cover and ws_cover['C14'].value) else file_name
    month_match = re.search(r'(\d{2,4})년?\s*(\d{1,2})월?', str(raw_month))
    if month_match:
        y = month_match.group(1)
        y = f"20{y}" if len(y) == 2 else y
        m = str(month_match.group(2)).zfill(2)
        claim_month = f"{y}-{m}"
    else:
        claim_month = "2026-09"

    # (5) 금회 위탁비용
    claim_cost = clean_number(ws_gap['D13'].value)
    if claim_cost == 0.0:
        for r in range(10, 20):
            val = clean_number(ws_gap.cell(r, 4).value)
            if val > 0:
                claim_cost = val
                break

    # (6) 집계장 분석
    ws_jibgye = find_sheet_by_keywords(wb, ['집계장', '집계'])
    j_res = parse_jibgye_sheet(ws_jibgye)

    dia_now = j_res["dia_now"]
    dia_cum = j_res["dia_cum"]
    exp_dorm = j_res["exp_dorm"]
    exp_meals = j_res["exp_meals"]

    if dia_now == 0.0:
        ws_q = find_sheet_by_keywords(wb, ['물량내역서', 'DIA-INCH', '물량'])
        if ws_q:
            dia_now = clean_number(ws_q['J50'].value)
            dia_cum = clean_number(ws_q['L50'].value)

    if exp_dorm == 0.0:
        ws_rent = find_sheet_by_keywords(wb, ['숙소임차료', '숙소임차'])
        ws_mgt = find_sheet_by_keywords(wb, ['숙소관리비', '숙소관리'])
        r_amt = clean_number(ws_rent['I2'].value) if ws_rent else 0.0
        m_amt = clean_number(ws_mgt['I2'].value) if ws_mgt else 0.0
        exp_dorm = r_amt + m_amt

    if exp_meals == 0.0:
        ws_meal = find_sheet_by_keywords(wb, ['식대비용', '식대'])
        if ws_meal:
            exp_meals = clean_number(ws_meal['I2'].value)

    exp_total = exp_dorm + exp_meals

    # (7) 인건비
    craft_data = {
        '배관공': {'공수': 0.0, '금액': 0.0},
        '용접사': {'공수': 0.0, '금액': 0.0},
        '관리자': {'공수': 0.0, '금액': 0.0},
        '보통인부': {'공수': 0.0, '금액': 0.0},
        '도비공': {'공수': 0.0, '금액': 0.0}
    }

    work_sheet = find_sheet_by_keywords(wb, ['근로내역', '노무비명세서', '노무비', '출력일보'])
    if work_sheet:
        header_row = 5
        job_col, md_col, amt_col = 4, 39, 41
        for r in range(1, 10):
            row_vals = [str(work_sheet.cell(r, c).value or '') for c in range(1, 45)]
            if any('직종' in v for v in row_vals) and any('공수' in v for v in row_vals):
                header_row = r
                for c, val in enumerate(row_vals, 1):
                    if '직종' in val: job_col = c
                    elif '공수' in val: md_col = c
                    elif any(k in val for k in ['총액', '지급액', '합계', '금액']): amt_col = c
                break

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
        "PJT_NO": pjt_no,
        "PJT명": pjt_name,
        "협력업체명": vendor_name,
        "청구월": claim_month,
        "금회_위탁비용": claim_cost,
        "경비_숙소임차료": 0.0,
        "경비_숙소관리비": 0.0,
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
# 4. 전체 데이터 후처리: 동적 상호 통합 및 구글 시트 적재
# --------------------------------------------------------------------------
def main():
    excel_files = glob.glob('data/*.xlsx') + glob.glob('data/*.xls')
    if not excel_files:
        print("[INFO] data/ 폴더에 처리할 기성 파일이 없습니다.")
        return

    print(f"[INFO] 총 {len(excel_files)}개의 기성 파일을 일괄 처리합니다.")
    records = []
    for f in excel_files:
        try:
            rec = parse_single_excel(f)
            records.append(rec)
        except Exception as e:
            print(f" [오류] {f} 파싱 중 예외 발생: {e}")

    if not records:
        print("[WARN] 유효한 기성 데이터가 없습니다.")
        return

    # [핵심] 2-Pass 동적 상호 자동 통합 (코드 수정 없이 신규 업체 자동 흡수)
    # 1. 4글자 이상의 정식 풀네임 목록을 동적으로 자동 수집
    distinct_vendors = {r['협력업체명'] for r in records if r['협력업체명'] and r['협력업체명'] != '기타'}
    long_names = sorted([v for v in distinct_vendors if len(v) >= 4], key=len, reverse=True)

    # 2. 2~3글자 약칭으로 들어간 업체명을 해당 풀네임으로 자동 치환
    for r in records:
        v = r['협력업체명']
        if len(v) <= 3 and v != '기타':
            for full in long_names:
                if full.startswith(v):
                    r['협력업체명'] = full
                    break
        print(f" [완료] [{r['PJT_NO']}] {r['협력업체명']} ({r['청구월']}): 위탁 {r['금회_위탁비용']:,}원 | DIA {r['금회_DIA_수량']}")

    df = pd.DataFrame(records)
    df = df.sort_values(by=['PJT_NO', '청구월', '협력업체명'], ascending=[True, True, True])

    # 구글 시트 적재
    service_account_json = os.environ.get('GCP_SERVICE_ACCOUNT_KEY')
    spreadsheet_id = os.environ.get('SPREADSHEET_ID')

    if not service_account_json or not spreadsheet_id:
        print("[ERROR] GCP_SERVICE_ACCOUNT_KEY 또는 SPREADSHEET_ID 환경변수가 설정되지 않았습니다.")
        df.to_csv("parsed_billing_summary.csv", index=False, encoding='utf-8-sig')
        sys.exit(1)

    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

    try:
        creds_dict = json.loads(service_account_json.strip())
        credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        gc = gspread.authorize(credentials)
        sh = gc.open_by_key(spreadsheet_id.strip())
        ws = sh.get_worksheet(0)

        data_to_write = [df.columns.values.tolist()] + df.fillna(0).values.tolist()
        ws.clear()
        ws.update(range_name='A1', values=data_to_write)
        print(f"[SUCCESS] 구글 스프레드시트에 총 {len(df)}개 기성 데이터 덮어쓰기 적재 완료!")
    except Exception as e:
        print(f"[ERROR] 구글 시트 적재 실패: {e}")
        sys.exit(1)

if __name__ == '__main__':
    main()
