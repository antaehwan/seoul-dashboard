"""
서울역 매출 데이터 업데이트 스크립트
실행: python update_data.py
"""
import json
import os
from datetime import datetime
from openpyxl import load_workbook

if os.name == "nt":  # Windows
    _ONEDRIVE = r"C:\Users\안태환\Desktop\OneDrive"
else:  # macOS
    _ONEDRIVE = os.path.expanduser("~/Library/CloudStorage/OneDrive-개인")

EXCEL_PATH = os.path.join(_ONEDRIVE, "01. 외식업 신규 ( 25.06 ~ )", "00. 매출 & 손익", "00. 매출", "H.서울역 월간 매출 _ 26Y.xlsx")
PNL_PATH   = os.path.join(_ONEDRIVE, "01. 외식업 신규 ( 25.06 ~ )", "00. 매출 & 손익", "01. 실적", "H.서울역 누적 실적 _ 23.11 ~.xlsx")
PMIX_PATH  = os.path.join(_ONEDRIVE, "01. 외식업 신규 ( 25.06 ~ )", "00. 매출 & 손익", "01. 실적", "H.서울역 P-MIX.xlsx")
OUTPUT_PATH = os.path.join(os.path.dirname(__file__), "data", "data.json")

WEEKDAY_KO = ["월", "화", "수", "목", "금", "토", "일"]

def parse_sheet(ws, month):
    days = []
    for row in ws.iter_rows(min_row=3, values_only=True):
        date_25y = row[1]  # col B: 25Y 날짜
        if not isinstance(date_25y, datetime):
            continue

        # 26Y 날짜 = 25Y 날짜 + 1년
        try:
            date_26y = date_25y.replace(year=date_25y.year + 1)
        except ValueError:
            # 2월 29일 윤년 처리
            date_26y = date_25y.replace(year=date_25y.year + 1, day=28)

        target = row[16] or 0       # col Q: 26Y 목표
        sales_l = row[18] or 0      # col S: 26Y 홀 매출
        sales_d = row[19] or 0      # col T: 26Y 배달 매출
        count_l = row[21] or 0      # col V: 26Y 홀 객수
        count_d = row[22] or 0      # col W: 26Y 배달 객수

        actual = sales_l + sales_d
        count = count_l + count_d
        avg_spend = round(actual / count, 1) if count > 0 else 0

        # 전년 데이터
        prev_target = row[4] or 0   # col E: 25Y 목표
        prev_actual = row[5] or 0   # col F: 25Y 실적
        prev_count = (row[8] or 0)  # col I: 25Y 객수

        # 실적이 없고 목표도 없는 행은 빈 행으로 간주
        if target == 0 and sales_l == 0 and sales_d == 0:
            continue

        days.append({
            "date": date_26y.strftime("%Y-%m-%d"),
            "weekday": WEEKDAY_KO[date_26y.weekday()],
            "target": int(target),
            "actual": int(actual),
            "sales_l": int(sales_l),
            "sales_d": int(sales_d),
            "count": int(count),
            "count_l": int(count_l),
            "count_d": int(count_d),
            "avg_spend": avg_spend,
            "prev_target": int(prev_target),
            "prev_actual": int(prev_actual),
            "prev_count": int(prev_count),
        })
    return days

def parse_pnl():
    wb = load_workbook(PNL_PATH, data_only=True)
    ws = wb['26Y 실적']
    rows = list(ws.iter_rows(min_row=1, max_row=200, max_col=20, values_only=True))

    # 행 찾기
    idx = {}
    for i, row in enumerate(rows):
        label = row[1]
        if label == '매출': idx['매출'] = i
        elif label == '재료비': idx['재료비'] = i
        elif label == '노무비': idx['노무비'] = i
        elif label == '일반관리비': idx['일반관리비'] = i
        elif label == '영업이익(매장)': idx['영업이익'] = i

    pnl = {}
    for m in range(1, 13):
        vc = 4 + (m - 1) * 2  # value col index
        rc = vc + 1             # ratio col index

        def v(key):
            if key not in idx: return None
            row = rows[idx[key]]
            return row[vc] if vc < len(row) else None

        def r(key):
            if key not in idx: return None
            row = rows[idx[key]]
            return row[rc] if rc < len(row) else None

        revenue = v('매출')
        if not revenue:
            continue

        pnl[str(m)] = {
            "revenue": int(revenue) if revenue else 0,
            "food_cost": int(v('재료비') or 0),
            "food_cost_pct": round((r('재료비') or 0) * 100, 1),
            "labor_cost": int(v('노무비') or 0),
            "labor_cost_pct": round((r('노무비') or 0) * 100, 1),
            "ga_cost": int(v('일반관리비') or 0),
            "ga_cost_pct": round((r('일반관리비') or 0) * 100, 1),
            "op_profit": int(v('영업이익') or 0),
            "op_profit_pct": round((r('영업이익') or 0) * 100, 1),
        }
        print(f"  P&L {m}월: 매출 {revenue:,.0f}원, 영업이익 {pnl[str(m)]['op_profit_pct']}%")

    return pnl

def parse_pmix():
    wb = load_workbook(PMIX_PATH, data_only=True)
    pmix = {}
    for sheet_name in wb.sheetnames:
        try:
            m = int(sheet_name)
        except ValueError:
            continue
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(min_row=1, max_row=60, max_col=15, values_only=True))

        def clean_name(s):
            if not s: return ""
            return str(s).replace("(H.CD)", "").replace("(H.CD", "").strip()

        # 전체 메뉴: col C(2)=상품명, G(6)=판매수량, H(7)=판매금액 (row 4~끝)
        # col J 이론원가 섹션과 무관하게 col C/G/H 데이터 모두 수집
        all_menus = []
        total_qty = 0; total_revenue = 0
        EXCLUDE_CATEGORIES = {'조식', '음료', '추가메뉴'}
        EXCLUDE_KEYWORDS   = ['[아침]', '사이다', '펩시', '맥주', '밀크티', '레몬티', '홍차', '포장용기', '고수추가']
        seen = set()
        current_cat = ''
        for i in range(3, len(rows)):
            row = rows[i]
            cat_cell = row[1]  # col B = 분류
            if cat_cell: current_cat = str(cat_cell).strip()
            if current_cat in EXCLUDE_CATEGORIES: continue
            name = row[2]; qty = row[6]; price = row[3]
            rev = row[7] if isinstance(row[7], (int, float)) and row[7] > 0 else (price * qty if isinstance(price, (int, float)) and isinstance(qty, (int, float)) else None)
            n = clean_name(name)
            if not n: continue
            if any(k in n for k in EXCLUDE_KEYWORDS): continue
            if not (isinstance(qty, (int, float)) and qty > 0): continue
            if not (isinstance(rev, (int, float)) and rev > 0): continue
            if n in seen: continue
            seen.add(n)
            all_menus.append({"name": n, "qty": int(qty), "revenue": int(rev), "category": current_cat})
            total_qty += int(qty)
            total_revenue += int(rev)

        # TOP 10: col J(9)/K(10) 판매수량, col M(12)/N(13) 매출
        sales_top10, revenue_top10 = [], []
        for i in range(4, 14):
            if i >= len(rows): break
            row = rows[i]
            name_s = row[9]; qty = row[10]
            name_r = row[12]; rev = row[13]
            clean = lambda s: str(s).replace("(H.CD)", "").replace("(H.C", "").strip() if s else ""
            if name_s and isinstance(qty, (int, float)) and qty > 0:
                sales_top10.append({"name": clean(name_s), "qty": int(qty)})
            if name_r and isinstance(rev, (int, float)) and rev > 0:
                revenue_top10.append({"name": clean(name_r), "revenue": int(rev)})

        # 예상 이론원가: col J(9), K(10), 행16(idx15)부터
        theory_cost = {}
        for i in range(15, 35):
            if i >= len(rows): break
            row = rows[i]
            label = row[9]; val = row[10]
            if label and isinstance(val, (int, float)):
                theory_cost[str(label)] = round(float(val) * 100, 1)

        if sales_top10 or revenue_top10 or all_menus:
            pmix[str(m)] = {
                "menus_all": all_menus,
                "sales_top10": sales_top10,
                "revenue_top10": revenue_top10,
                "theory_cost": theory_cost,
                "total_revenue": total_revenue,
                "total_qty": total_qty,
            }
            print(f"  P-MIX {m}월: 판매TOP{len(sales_top10)} 매출TOP{len(revenue_top10)}")
    return pmix

def main():
    print(f"엑셀 파일 읽는 중...")
    wb = load_workbook(EXCEL_PATH, data_only=True)

    all_data = {}
    for month in range(1, 13):
        sheet_name = str(month)
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        days = parse_sheet(ws, month)
        if days:
            all_data[sheet_name] = days
            print(f"  {month}월: {len(days)}일 데이터 읽음")

    print(f"\nP&L 파일 읽는 중...")
    pnl_data = parse_pnl()

    print(f"\nP-MIX 파일 읽는 중...")
    pmix_data = parse_pmix()

    output = {
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "store": "서울역",
        "store_code": "4018",
        "months": all_data,
        "pnl": pnl_data,
        "pmix": pmix_data,
    }

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"\n완료! → {OUTPUT_PATH}")

if __name__ == "__main__":
    main()
