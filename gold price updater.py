import requests
import pandas as pd
import math  # นำเข้าไลบรารี math สำหรับใช้ปัดเศษขึ้น


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

                # คำนวณลด 5% และปัดเศษขึ้นทันทีด้วย math.ceil()
                gold_ornament_buy = math.ceil(gold_bar_buy * 0.95)

                excel_data = [
                    ["รูปพรรณ ขายออก", f"{gold_ornament_sell:,.2f}"],
                    ["รูปพรรณ รับซื้อ", f"{gold_ornament_buy:,.2f}"],
                    ["ทองแท่ง ขายออก", f"{gold_bar_sell:,.2f}"],
                    ["ทองแท่ง รับซื้อ", f"{gold_bar_buy:,.2f}"],
                    [
                        "เวลาอัปเดต",
                        f"อัปเดตเมื่อ: {resp_data.get('update_date', '')} {resp_data.get('update_time', '')}",
                    ],
                ]

                df = pd.DataFrame(excel_data)
                df.to_excel("controller.xlsx", index=False, header=False)
                print("คำนวณ ปัดเศษขึ้น และอัปเดตไฟล์ controller.xlsx สำเร็จ! - gold price updater.py:45")
            else:
                print("API แจ้งเตือนสถานะ: - gold price updater.py:47", data)
        else:
            print(f"เชื่อมต่อ API ไม่สำเร็จ รหัสสถานะ: {response.status_code} - gold price updater.py:49")

    except Exception as e:
        print("เกิดข้อผิดพลาด: - gold price updater.py:52", e)


if __name__ == "__main__":
    update_gold_excel()
