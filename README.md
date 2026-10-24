# GC-Autoload
VB code which have to be pasted into .xlam file again for showing as a ribbon in Excel

## การเพิ่ม OID (v3.4)

OID คือเลขออร์เดอร์เดิมจากคอลัมน์ B ของ `AutoLoad-Input` เก็บและเขียนเป็น Text เพื่อรักษาศูนย์นำหน้าและเลขยาว ส่วน Barcode ยังคงเข้ารหัสด้วย `CODE128_LIBRE` ตอน Load Data

ใช้เทมเพลตที่เพิ่ม OID แล้ว โดยปรับตำแหน่งดังนี้:

| ชีต | ตำแหน่ง | ข้อมูล |
|---|---|---|
| AutoLoad-Input | B1 | SKU ที่เลือก |
| AutoLoad-Input | A3 | จำนวนชุดต่อแถว |
| AutoLoad-Input | B3:E3 | จำนวนสำเนา NameA–NameD |
| AutoLoad-Input | F3 | จำนวนสำเนา OID (0 คือไม่เขียน OID) |
| AutoLoad-Input | G3 | จำนวนสำเนา Barcode |
| AutoLoad-Input | H3 | วันที่ |
| AutoLoad-Input | I3 | Platform / รหัสร้าน |
| NormalizedInput | A:H | SKU, NameA, NameB, NameC, NameD, OID, Barcode, Qty |

ข้อมูลต้นทางแถว 5 เป็นต้นไปยังใช้ A=SKU, B=Order ID, C=Qty, D=Raw Input เหมือนเดิม หาก F3:I3 อ้างอิงสูตรจากชีต Database ต้องปรับสูตรในเทมเพลตให้ตรงกับคอลัมน์ที่เพิ่มด้วย ซอร์สนี้อ่านค่าจาก AutoLoad-Input โดยไม่ได้แก้ไข Database หรือสูตรให้

ชีต SKU ต้องมีคอลัมน์ข้อมูลแต่ละชุดตามลำดับ NameA → NameB → NameC → NameD → OID → Barcode โดยจำนวนคอลัมน์แต่ละประเภทตรงกับ B3:G3 และทำซ้ำตาม A3 เช่น เมื่อมีชื่อหนึ่งช่อง, OID หนึ่งช่อง และ Barcode หนึ่งช่องต่อชุด:

```text
N1-A1 | N1-OID | N1-BC | N2-A1 | N2-OID | N2-BC | DATE | SHOP | SKU | FN
```

ชื่อหัวคอลัมน์ต้องไม่ซ้ำกันและตรงกับตัวแปรใน Illustrator โดย OID ใช้ `Nx-OID` และ Barcode ใช้ `Nx-BC` ให้เพิ่มเลข x ตามแต่ละชุด คอลัมน์ DATE, SHOP, SKU, FN อยู่หลังข้อมูลทุกชุดและค้นหาตามชื่อหัวคอลัมน์ จึงย้ายตำแหน่งได้ตามขนาดชุดใหม่

โค้ดไม่แทรกคอลัมน์หรือสร้างหัวตารางใน workbook ให้ เมื่อปรับเทมเพลตแล้วให้เรียก Translate Input → Load Data → Export XML ใหม่ตามลำดับ อย่าใช้ข้อมูล NormalizedInput จากโครงสร้างเดิม Export XML อ่านหัวคอลัมน์และส่งออก OID เป็นข้อความโดยอัตโนมัติ

## การตรวจสอบ OID ก่อนใช้งานจริง

ใช้สำเนา development ใน Excel แยกจาก add-in production แล้ว Compile VBAProject ก่อนทดสอบด้วยข้อมูลสมมติเท่านั้น:

1. ใช้ Order ID แบบ Text เช่น `00012345678901234567` ตรวจว่าคงค่าเดิมทั้ง NormalizedInput!F:G, Nx-OID และ XML ส่วน Nx-BC ยังเป็นข้อความบาร์โค้ดที่เข้ารหัส
2. กำหนด A3=2, NameA=1, OID=1, Barcode=1 และ Qty=3 ตรวจว่ามี 2 ชุดในแถวแรกและ 1 ชุดในแถวถัดไป โดย OID อยู่ก่อน Barcode ทุกชุด และ DATE/SHOP/SKU/FN ไม่ถูกเขียนทับ
3. ทดสอบ OID=0 และ OID=2 โดยเตรียมหัวคอลัมน์ให้ตรงจำนวนสำเนา ตรวจจำนวนและลำดับข้อมูล รวมถึง Barcode=0
4. ทดสอบ `SKU *3pcs || ALPHA *2pcs \ BETA *1pcs` กับ Qty=3 และ Qty=2 โดยแทน SKU ด้วยรหัสสินค้าสมมติที่เลือกไว้ กรณีจำนวนตรงต้องได้ OID เดียวกันทั้งสองรายการ กรณีไม่ตรงต้องไฮไลต์และข้ามกลุ่ม
5. ทดสอบชื่อเดียวไม่ระบุจำนวน, Order ID + SKU ซ้ำหลายแถว, SKU ต่างกันในออร์เดอร์เดียวกัน และ Raw Input ขัดแย้งกัน เพื่อยืนยันการรวม Qty และการเลือก SKU เดิม
6. ทดสอบชื่อไทยและ Unicode เช่น `ชื่อทดสอบ 中文 한글 😀 & < > " '` ตรวจข้อความและ XML escaping รวมถึงการขึ้นบรรทัดด้วย `||`
7. ทดลองเทมเพลตขาด OID หรือวาง OID/Barcode ผิดตำแหน่ง ตรวจว่าหยุดก่อนล้างข้อมูลเดิม และตรวจ Clear Data, Ribbon callbacks และ version label

รายการข้างต้นเป็นขั้นตอนทดสอบที่ต้องทำใน Excel ไม่ใช่ผลทดสอบที่ผ่านแล้ว Repository นี้ยังไม่มี workbook/add-in และเทมเพลต Illustrator สำหรับตรวจครบทั้งระบบ
