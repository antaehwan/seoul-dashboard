"""
영등포 매출 데이터 업데이트 스크립트
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

EXCEL_PATH = os.path.join(_ONEDRIVE, "01. 외식업 신규 ( 26.08~ )", "00. 매출 & 손익", "00. 매출", "H.영등포 월간 매출 _ 26Y.xlsx")
PNL_PATH   = os.path.join(_ONEDRIVE, "01. 외식업 신규 ( 26.08~ )", "00. 매출 & 손익", "01. 실적", "H.영등포 누적 실적 _ 25Y~.xlsx")
PMIX_PATH  = os.path.join(_ONEDRIVE, "01. 외식업 신규 ( 26.08~ )", "00. 매출 & 손익", "01. 실적", "H.영등포 P-MIX _ 26Y.xlsx")
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
    rows = list(ws.iter_rows(min_row=1, max_row=220, max_col=20, values_only=True))

    idx = {}
    sub_labels = {
        '재료비': ['1) 식료재료비', '2) 음료재료비', '3) 기타재료비'],
        '노무비': ['1) 고정비', '2) 변동비', '3) 연차수당', '4) 상여금', '5) 퇴직급여'],
        '일반관리비': [
            '1) 복리후생비', '2) 교육훈련비', '3) 여비교통비', '4) 통신비',
            '5) 수도광열비', '6) 세금과공과', '7) 지급수수료', '8) 임차료',
            '9) 감가상각비', '10) 수선비', '11) 소모품비', '12) 도서인쇄비',
            '13) 보험료', '14) 광고선전비', '15) 기타', '16) 라이선스 수수료',
        ],
    }
    sub_idx = {k: {} for k in sub_labels}
    for i, row in enumerate(rows):
        b, c = row[1], row[2]
        if b == '매출': idx['매출'] = i
        elif b == '재료비': idx['재료비'] = i
        elif b == '노무비': idx['노무비'] = i
        elif b == '일반관리비': idx['일반관리비'] = i
        elif b == '영업이익(매장)': idx['영업이익'] = i
        if c:
            for parent, labels in sub_labels.items():
                if c in labels:
                    sub_idx[parent][c] = i

    pnl = {}
    for m in range(1, 13):
        vc = 4 + (m - 1) * 2
        rc = vc + 1

        def v(i):
            row = rows[i]
            val = row[vc] if vc < len(row) else None
            return int(val) if isinstance(val, (int, float)) else 0

        def r(i):
            row = rows[i]
            val = row[rc] if rc < len(row) else None
            return round(float(val) * 100, 1) if isinstance(val, (int, float)) else 0.0

        if '매출' not in idx: continue
        revenue = rows[idx['매출']][vc] if vc < len(rows[idx['매출']]) else None
        if not revenue: continue

        def sub_detail(parent):
            parent_amt = v(idx[parent]) if parent in idx else 0
            result = []
            for lbl in sub_labels[parent]:
                if lbl not in sub_idx[parent]: continue
                amt = v(sub_idx[parent][lbl])
                pct = round(amt / parent_amt * 100, 1) if parent_amt else 0.0
                result.append({"label": lbl.split(') ', 1)[1], "amount": amt, "pct": pct})
            return result

        pnl[str(m)] = {
            "revenue": int(revenue),
            "food_cost": v(idx['재료비']) if '재료비' in idx else 0,
            "food_cost_pct": r(idx['재료비']) if '재료비' in idx else 0.0,
            "food_cost_detail": sub_detail('재료비'),
            "labor_cost": v(idx['노무비']) if '노무비' in idx else 0,
            "labor_cost_pct": r(idx['노무비']) if '노무비' in idx else 0.0,
            "labor_cost_detail": sub_detail('노무비'),
            "ga_cost": v(idx['일반관리비']) if '일반관리비' in idx else 0,
            "ga_cost_pct": r(idx['일반관리비']) if '일반관리비' in idx else 0.0,
            "ga_cost_detail": sub_detail('일반관리비'),
            "op_profit": v(idx['영업이익']) if '영업이익' in idx else 0,
            "op_profit_pct": r(idx['영업이익']) if '영업이익' in idx else 0.0,
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
        rows = list(ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=16, values_only=True))

        def clean_name(s):
            if not s: return ""
            return str(s).replace("(H.CD)", "").replace("(H.CD", "").strip()

        # ── 왼쪽: 홀 메뉴 (col B~H, index 1~7) ──
        all_menus = []
        total_qty = 0; total_revenue = 0
        EXCLUDE_CATEGORIES = {'추가메뉴'}
        seen = set()
        current_cat = ''
        for i in range(3, len(rows)):
            row = rows[i]
            cat_cell = row[1]
            if cat_cell and isinstance(cat_cell, str): current_cat = cat_cell.strip()
            if current_cat in EXCLUDE_CATEGORIES: continue
            name = row[2]; qty = row[6]; price = row[3]; cost_amt = row[4]
            rev = row[7] if isinstance(row[7], (int, float)) and row[7] > 0 \
                else (price * qty if isinstance(price, (int, float)) and isinstance(qty, (int, float)) else None)
            n = clean_name(name)
            if not n or not isinstance(qty, (int, float)) or not isinstance(price, (int, float)): continue
            if n in seen: continue
            seen.add(n)
            qty_int = int(qty)
            rev_int = int(rev) if isinstance(rev, (int, float)) else 0
            cost_rate = round(cost_amt / (price / 1.1) * 100, 1) \
                if isinstance(cost_amt, (int, float)) and price > 0 else None
            all_menus.append({"name": n, "qty": qty_int, "revenue": rev_int,
                              "category": current_cat, "cost_rate": cost_rate})
            total_qty += qty_int
            total_revenue += rev_int

        # TOP 10: 왼쪽 메뉴 데이터에서 직접 계산
        valid = [m for m in all_menus if m['qty'] > 0 and m['revenue'] > 0]
        sales_top10  = [{"name": m["name"], "qty": m["qty"]}
                        for m in sorted(valid, key=lambda x: -x['qty'])[:10]]
        revenue_top10 = [{"name": m["name"], "revenue": m["revenue"]}
                         for m in sorted(valid, key=lambda x: -x['revenue'])[:10]]

        # 이론원가: 카테고리별 가중평균 원가율
        cat_rev = {}; cat_cm = {}
        for menu in all_menus:
            if menu['cost_rate'] is None or menu['revenue'] <= 0: continue
            cat = menu['category']
            cat_rev[cat] = cat_rev.get(cat, 0) + menu['revenue']
            cat_cm[cat]  = cat_cm.get(cat, 0)  + menu['revenue'] * menu['cost_rate'] / 100
        theory_cost = {cat: round(cat_cm[cat] / rev * 100, 1)
                       for cat, rev in cat_rev.items() if rev > 0}

        # ── 오른쪽: 배달 & 프로모션 (col J~P, index 9~15) ──
        delivery_revenue = 0
        promo_revenue = 0
        current_right_cat = ''
        for row in rows[3:]:
            cat = row[9]
            name = row[10]
            rev = row[15]
            if cat and isinstance(cat, str) and cat not in ('소분류',):
                current_right_cat = cat.strip()
            # 이름 없는 행(소계행) 제외
            if not (isinstance(name, str) and name): continue
            if not isinstance(rev, (int, float)) or rev <= 0: continue
            if current_right_cat == '배달':
                delivery_revenue += rev
            elif current_right_cat == '프로모션':
                promo_revenue += rev

        if all_menus:
            pmix[str(m)] = {
                "menus_all": all_menus,
                "sales_top10": sales_top10,
                "revenue_top10": revenue_top10,
                "theory_cost": theory_cost,
                "total_revenue": total_revenue,
                "total_qty": total_qty,
                "delivery_revenue": int(delivery_revenue),
                "promo_revenue": int(promo_revenue),
            }
            print(f"  P-MIX {m}월: 판매TOP{len(sales_top10)} 매출TOP{len(revenue_top10)} 배달 {int(delivery_revenue):,}원 프로모션 {int(promo_revenue):,}원")
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

    # 월별 고객수·객단가를 P&L에 추가
    for m_str, days in all_data.items():
        if m_str not in pnl_data: continue
        total_count = sum(d['count'] for d in days if d['actual'] > 0)
        total_actual = sum(d['actual'] for d in days)
        avg_spend = round(total_actual / total_count) if total_count else 0
        pnl_data[m_str]['customer_count'] = total_count
        pnl_data[m_str]['avg_spend'] = avg_spend

    print(f"\nP-MIX 파일 읽는 중...")
    pmix_data = parse_pmix()

    output = {
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "store": "영등포",
        "store_code": "영등포",
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
