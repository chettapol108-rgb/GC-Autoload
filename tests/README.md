# ทดสอบหลาย SKU แยกตาม Platform

ชุดทดสอบนี้ยังไม่ได้รันใน Excel ระหว่างพัฒนา การเข้าถึง VBProject ผ่าน automation ใน session ทดสอบใช้ไม่ได้ Codex จึงไม่มีผล Compile หรือผลรันชุดทดสอบนี้ผ่าน automation ผู้ใช้แจ้งว่าทดสอบผ่านแล้วและอนุมัติ v3.5.0 เมื่อ 2026-10-31 แต่ไม่ได้ระบุว่าได้รันชุดทดสอบนี้ครบทุก assertion หรือไม่

## วิธีรันใน Excel

1. ใช้ Excel session แยกและสำเนา development ของ add-in ห้ามโหลดคู่กับ production เก็บต้นฉบับไว้ครบ ไม่ใช้ข้อมูลออร์เดอร์จริง
2. วาง source ใน `src/` ลงใน standard modules ชื่อเดิมให้ครบ ไม่สร้างโมดูลซ้ำ ไฟล์ source ส่วนใหญ่ไม่มี Attribute metadata จึงอย่าสมมติว่าการ Import จะคืนชื่อโมดูลเดิมโดยอัตโนมัติ ไม่ต้องเปลี่ยน Trust Center เพื่อรันชุดทดสอบผ่าน VBA editor ด้วยมือ
3. Import `tests/modBatchTests.bas` เป็น standard module `modBatchTests` ใช้เฉพาะ development ไม่รวมโมดูลทดสอบใน release
4. เลือก Debug → Compile VBAProject ก่อน จากนั้นรัน `GC_TestBatchProcessing` ระบบสร้าง workbook ใหม่ด้วยข้อมูลสมมติทั้งหมด ไม่มีการเปิดไฟล์ production หรือบันทึก workbook ทับไฟล์ใด
5. กดรับข้อความผลของแต่ละขั้น รวมข้อความ error ที่จงใจสร้าง ชุดทดสอบจะหยุดด้วย `BATCH TEST FAILED` หาก assertion ไม่ตรง หากครบจะแสดง `Batch assertions passed` และทิ้ง workbook สมมติไว้สำหรับทดสอบ XML ต่อ

Assertions ครอบคลุม SP กับ Shopee แบบไม่สนตัวพิมพ์/ช่องว่าง, TT กับ TikTok/Facebook, ALL รวมทุก platform และไม่ตรวจค่าในคอลัมน์ C, การไม่พึ่ง B1/A3:G3, การรวม Order ID + SKU ซ้ำ, SKU ต่างกันในออร์เดอร์เดียว, จำนวน explicit ตรง/ไม่ตรง, จำนวนที่ไม่ระบุ explicit, Raw Input ขัดแย้ง, จำนวนสำเนา NameA/NameB ต่างกันตาม SKU, OID/Barcode เป็นศูนย์, การตัดแถวตาม Set, หัว OID/BC ทั้ง `_` และ `-`, OID ยาวและศูนย์นำหน้า, Unicode, การบล็อกเมื่อเปลี่ยน platform, การตรวจทุกปลายทางก่อนล้าง, SKU ที่หาไม่พบ/ซ้ำใน Database และสถานะหลัง Load ผิดพลาด

## ตรวจด้วยมือเพิ่มเติม

- รัน Export XML กับ workbook สมมติที่ชุดทดสอบทิ้งไว้ เลือกโฟลเดอร์ทดสอบ ต้องได้โฟลเดอร์รอบ SP ใหม่และไฟล์ `TEST-A.xml`, `TEST-B.xml` เท่านั้น การยกเลิกเลือกโฟลเดอร์ต้องไม่สร้างไฟล์ รันซ้ำต้องสร้างโฟลเดอร์ใหม่ ไม่เขียนทับ
- ตรวจ TEST-A.xml ว่าศูนย์นำหน้า OID, ภาษาไทย/จีน/เกาหลี/emoji อยู่ครบ และ `& < > " '` ถูก escape ถูกต้อง ตรวจบาร์โค้ดและการแปลง `||` เป็นขึ้นบรรทัดตามพฤติกรรมเดิม แล้วเปิดกับเทมเพลต Illustrator ที่ตรงกับหัวตัวแปร
- เปลี่ยน I3 เป็น TT ต้องบล็อก Export ก่อน Translate/Load ใหม่ หลังทำครบแล้วต้องได้เฉพาะข้อมูล TikTok/Facebook และ SHOP/FN ใช้ TT ห้ามนำผล SP เก่ามาปะปน
- เลือก I3=ALL แล้ว Translate/Load ใหม่ ต้องรวม Shopee, TikTok, Facebook และ platform ว่าง พร้อม SHOP/FN เป็น ALL และ Export ไปโฟลเดอร์ AutoLoad-ALL-...; ทดสอบเปลี่ยนจาก SP/TT เป็น ALL แล้ว Export ก่อน Translate/Load ใหม่ต้องถูกบล็อก
- ทดสอบ platform ที่ไม่ตรงเลยและ I3 ที่ไม่ใช่ SP/TT/ALL ต้องหยุดและไม่อนุญาตให้ใช้ normalized รอบเก่า การเลือก TT จะรวม platform ว่างตามนิยามค่าที่ไม่ใช่ Shopee
- แก้ H3 หรือสำเนา Database ระหว่าง Translate กับ Load/Export ต้องบังคับ Translate ใหม่ ทดสอบ Set=0, สำเนาติดลบ/ทศนิยม, สำเนาทั้งหมดเป็น 0, ชีตปลายทางหาย/ป้องกัน และหัว OID/BC ผิดตำแหน่ง ต้องหยุดก่อนล้างชีตปลายทาง
- ทดสอบสำเนา OID=2 ด้วยหัว OID ที่ไม่ซ้ำ, Barcode=0 และ NameC/NameD ที่มีหลายสำเนา เตรียมหัวชีตให้ตรงค่าตั้ง ตรวจ SKU/DATE/SHOP/FN ไม่ถูกเขียนทับ
- เพิ่มชีต SKU ที่มีข้อมูลเก่าแต่ไม่มี SKU นั้นใน input รอบที่เลือก ต้องไม่ถูกล้างและไม่ถูก Export
- บันทึกและเปิดสำเนาทดสอบใหม่ ตรวจ CustomProperties คงอยู่ และ Load/Export ตรวจ platform ได้เหมือนเดิม ทดสอบจำนวน SKU มากพอให้ข้อมูลค่าตั้งยาวเกิน 200 ตัวอักษร
- ตรวจปุ่ม Ribbon เดิมทั้ง Translate, Load, Export, Clear Data และ version label ต้องแสดง GC AutoLoad v3.5.0 ไม่มีการสร้าง Ribbon XML ในงานนี้ Clear Data ยังคงอิง B1 และต้อง Load ใหม่ก่อน Export
- เปรียบเทียบ parser กับ workflow เดิมด้วยข้อมูลสมมติ ทั้งชื่อเดี่ยว, หลายชื่อ, หลาย SKU ใน Raw Input และข้อความไทยที่มี marker บรรทัด

ถ้าทดสอบด้วย automation ในภายหลัง ต้องเข้าถึง VBA project object model ได้ อธิบายและขออนุญาตก่อนเปลี่ยนค่าความปลอดภัยใด ๆ การทดสอบด้วย VBA editor ด้วยมือไม่จำเป็นต้องเปิดสิทธิ์ automation นี้
