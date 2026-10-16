import math
from datetime import datetime
import openpyxl
import requests


def update_gold_excel():
  try:
    url = "https://api.chnwt.dev/thai-gold-api/latest"
    response = requests.get(url, timeout=10)

    if response.status_code == 200:
      data = response.json()

      if data.get("status") == "success":
        resp_data = data["response"]
        price_data = resp_data["price"]

        # แปลงค่าเป็น float โดยตัดคอมมาออก
        gold_bar_sell = float(
            str(price_data["gold_bar"]["sell"]).replace(",", "")
        )
        gold_bar_buy = float(
            str(price_data["gold_bar"]["buy"]).replace(",", "")
        )
        gold_ornament_sell = float(
            str(price_data["gold"]["sell"]).replace(",", "")
        )
        gold_ornament_buy = math.ceil(gold_bar_buy * 0.95)

        # 🧮 คำนวณส่วนต่าง (Spread) ซื้อ-ขาย เรียลไทม์
        spread_ornament = gold_ornament_sell - gold_ornament_buy
        spread_bar = gold_bar_sell - gold_bar_buy

        update_str = f"อัปเดตเมื่อ: {resp_data.get('update_date', '')} {resp_data.get('update_time', '')}"

        # โหลดไฟล์ Excel ด้วย openpyxl
        wb = openpyxl.load_workbook("controller.xlsx")
        ws = wb.active

        # 1. อัปเดตราคาปัจจุบันไว้ที่ A1:B5 (สำหรับหน้าเว็บแสดงผล)
        ws["A1"] = "รูปพรรณ ขาย"
        ws["B1"] = f"{gold_ornament_sell:,.2f}"

        ws["A2"] = "รูปพรรณ รับ"
        ws["B2"] = f"{gold_ornament_buy:,.2f}"

        ws["A3"] = "ทองแท่ง ขาย"
        ws["B3"] = f"{gold_bar_sell:,.2f}"

        ws["A4"] = "ทองแท่ง รับ"
        ws["B4"] = f"{gold_bar_buy:,.2f}"

        ws["A5"] = "เวลาอัปเดต"
        ws["B5"] = update_str

        # 2. ค้นหาแถวสุดท้ายที่มีข้อมูลในตารางรายวันฝั่งขวา (เริ่มที่คอลัมน์ H)
        last_row = 2
        while ws[f"H{last_row + 1}"].value is not None:
          last_row += 1

        # เช็กเงื่อนไข: ถ้าราคาทั้ง 4 ช่องเท่าเดิมเป๊ะๆ จะไม่เพิ่มแถวซ้ำ
        should_append = True
        if last_row >= 3:
          last_ornament_sell = float(ws[f"I{last_row}"].value or 0)
          last_ornament_buy = float(ws[f"J{last_row}"].value or 0)
          last_bar_sell = float(ws[f"K{last_row}"].value or 0)
          last_bar_buy = float(ws[f"L{last_row}"].value or 0)

          if (
              last_ornament_sell == gold_ornament_sell
              and last_ornament_buy == gold_ornament_buy
              and last_bar_sell == gold_bar_sell
              and last_bar_buy == gold_bar_buy
          ):
            should_append = False
            print(
                "ราคายังคงเท่าเดิม ไม่ทำการเพิ่มแถวประวัติใหม่ (อัปเดตเฉพาะหน้าเว็บ)"
            )

        # 3. ถ้ามีการเปลี่ยนแปลงราคา ให้เติมแถวใหม่พร้อมคำนวณส่วนต่างลงคอลัมน์ H ถึง N
        if should_append:
          next_row = last_row + 1 if last_row >= 2 else 3
          current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

          ws[f"H{next_row}"] = current_date  # เวลาอัปเดต
          ws[f"I{next_row}"] = gold_ornament_sell  # พรรณ ขาย
          ws[f"J{next_row}"] = gold_ornament_buy  # พรรณ ซื้อ
          ws[f"K{next_row}"] = gold_bar_sell  # แท่ง ขาย
          ws[f"L{next_row}"] = gold_bar_buy  # แท่ง ซื้อ
          ws[f"M{next_row}"] = spread_ornament  # ส่วนต่าง พรรณ
          ws[f"N{next_row}"] = spread_bar  # ส่วนต่าง แท่ง

          print(
              "ตรวจพบราคาเปลี่ยนแปลง! บันทึกประวัติและส่วนต่างลงตารางขวา (แถวที่"
              f" {next_row}) สำเร็จ"
          )

        # บันทึกไฟล์กลับ
        wb.save("controller.xlsx")

      else:
        print("API แจ้งเตือนสถานะ:", data)
    else:
        print(f"เชื่อมต่อ API ไม่สำเร็จ รหัสสถานะ: {response.status_code}")

  except Exception as e:
    print("เกิดข้อผิดพลาด:", e)


if __name__ == "__main__":
  update_gold_excel()
