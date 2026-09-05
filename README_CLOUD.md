# THANN Spa Stock Pro — Cloud Deployment

แพ็กนี้พร้อมนำขึ้นบริการ Cloud ที่รองรับ Flask + PostgreSQL เช่น Render

## วิธี Deploy
1. สร้าง Git repository และอัปโหลดไฟล์ทั้งหมดในโฟลเดอร์นี้
2. ใน Render เลือก New > Blueprint แล้วเลือก repository นี้
3. Render จะสร้าง Web Service และ PostgreSQL ตาม `render.yaml`
4. เปิด URL ที่ Render ให้มา
5. Login:
   - admin / thann1234
   - staff / staff1234

## สำคัญก่อนใช้งานจริง
- เปลี่ยนรหัสผ่านเริ่มต้น
- เปลี่ยน/เก็บ THANN_SECRET_KEY เป็นค่าลับ
- จำกัดสิทธิ์ผู้ใช้ตามบทบาทเพิ่มเติมหากต้องใช้จริงในองค์กร
- ระบบนี้ใช้ฐานข้อมูลกลาง PostgreSQL เมื่อกำหนด DATABASE_URL จึงแชร์ข้อมูลระหว่างเครื่อง/มือถือและสาขาได้
- Export เป็น .xlsx
