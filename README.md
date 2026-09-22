# WMS Thai Address Converter

แอป Streamlit สำหรับอัปโหลดไฟล์ Excel จาก WMS แล้วสร้าง:

- `จังหวัด` — ไทย/อังกฤษตามรูปแบบ `adress.xlsx`
- `เขต/อำเภอ` — ไทย/อังกฤษตามรูปแบบ `adress.xlsx`
- `รหัสไปรษณีย์` — ค้นจากจังหวัด + อำเภอ + ตำบล
- `สถานะตรวจสอบ` — ช่วยระบุรายการที่ต้องตรวจสอบ

## ไฟล์ในโปรเจกต์

- `app.py` — ตัวแอป
- `data/adress.xlsx` — address master ของผู้ใช้
- `requirements.txt` — dependencies
- `README.md` — คู่มือ

## วิธีรันบนคอมพิวเตอร์

ติดตั้ง Python 3.10+ แล้วเปิด Terminal ในโฟลเดอร์นี้:

```bash
pip install -r requirements.txt
streamlit run app.py
```

จากนั้นเปิด URL ที่ Streamlit แสดง เช่น `http://localhost:8501`

## รูปแบบ WMS ที่รองรับ

ต้องมีคอลัมน์:

- `Receipt Province`
- `Receipt City`
- `Receipt Area`
- `Consignee Addr`

ระบบจะคงข้อมูล WMS เดิมไว้ทั้งหมด และแทรก 3 คอลัมน์ใหม่หลัง `Receipt Area`

## หมายเหตุเรื่องรหัสไปรษณีย์

ระบบใช้ข้อมูลระดับตำบล เพราะอำเภอหนึ่งอาจมีรหัสไปรษณีย์มากกว่าหนึ่งรหัส หากระบุได้ถึงตำบล ระบบจะใช้รหัสระดับตำบลก่อน หากอำเภอมีรหัสเดียวจึงใช้ระดับอำเภอเป็น fallback หากมีหลายรหัสและระบุตำบลไม่ได้ จะขึ้นสถานะให้ตรวจสอบแทนการเดา

ฐานข้อมูลภูมิศาสตร์ที่แอปเรียกใช้:
https://github.com/thailand-geography-data/thailand-geography-json

