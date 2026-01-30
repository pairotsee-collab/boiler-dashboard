
# Boiler Fuel–Energy Dashboard — Cloud Ready + OneDrive/SharePoint

โค้ดชุดนี้พร้อมให้คุณ Deploy บน **Streamlit Cloud** และดึงไฟล์ Excel ได้ 4 วิธี:
1) อัปโหลด Excel (.xlsx)
2) URL สาธารณะ (OneDrive/SharePoint/GitHub) → direct download URL
3) **OneDrive/SharePoint (Graph API)** → ไม่ต้อง public link
4) พาธภายใน (เฉพาะรันบน LAN/On-Prem)

---
## วิธีใช้งานอย่างย่อ
1. อัปโหลดไฟล์ในโฟลเดอร์นี้ขึ้น GitHub (public) → เปิด https://share.streamlit.io → New app → เลือกไฟล์หลัก `app.py` → Deploy
2. หากใช้ **Graph API** ให้ตั้งค่า Secrets ก่อน (ดูด้านล่าง)

---
## ตั้งค่า Secrets (เฉพาะโหมด Graph)
ไปที่ **Streamlit Cloud → App → Settings → Secrets** แล้วเพิ่มค่า:

```toml
TENANT_ID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
CLIENT_ID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
CLIENT_SECRET = "your-client-secret"
```

### สิทธิ์ที่ต้องการบน Azure AD (Entra ID)
- Application permissions: `Files.Read.All` หรือ `Sites.Read.All`
- ทำ **Admin consent** ให้เรียบร้อย

---
## โครงสร้างคอลัมน์ในไฟล์ Excel (ต้องมีหัวตารางต่อไปนี้)
```
Date
ยอดบรรจุ(ตัน)
Cost fuel (Baht)
ใช้น้ำ m3
ไม้สับ (กก.)
เปลือกมะม่วงหิมพานต์ (กก.)
ไม้เฟอร์นิเจิร์บด (กก.)
```

---
## รันทดสอบบนเครื่อง
```bash
pip install -r requirements.txt
streamlit run app.py --server.address 0.0.0.0 --server.port 8501
```

---
## หมายเหตุ
- โหมด URL สาธารณะ: OneDrive/SharePoint ให้ตั้งแชร์ Anyone และต่อท้าย `?download=1`
- ถ้าข้อมูลอ่อนไหว แนะนำโหมด **Graph API** แทนการเปิด public link
