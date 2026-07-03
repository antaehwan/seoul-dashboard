"""
서울역 매출 데이터 업데이트 스크립트
실행: python update_data.py
"""
import json
import os
from datetime import datetime
from openpyxl import load_workbook

EXCEL_PATH = r"C:\Users\안태환\Desktop\OneDrive\01. 외식업 신규 ( 25.06 ~ )\00. 매출 & 손익\00. 매출\H.서울역 월간 매출 _ 26Y.xlsx"
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

    output = {
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "store": "서울역",
        "store_code": "4018",
        "months": all_data
    }

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"\n완료! → {OUTPUT_PATH}")

if __name__ == "__main__":
    main()
