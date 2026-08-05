/* =====================================================================
   bc-items.sample.js  —  ตัวอย่างรูปแบบไฟล์ข้อมูล Item Card จาก Business Central
   ---------------------------------------------------------------------
   วิธีใช้:
   1) ดึงข้อมูล Item จาก BC (ดูวิธีในแชท / ท้ายไฟล์นี้) ให้ได้เป็น JSON
   2) วาง JSON นั้นลงในตัวแปร window.BC_ITEMS ข้างล่าง
   3) บันทึกไฟล์นี้ชื่อ  bc-items.js  (ตัดคำว่า .sample ออก) ไว้ "โฟลเดอร์เดียวกับ" stock-analysis.html
   4) เปิด stock-analysis.html ใหม่ → มุมขวาบนจะขึ้น "Item Card: BC จริง (Brand=Enrich)"

   หมายเหตุ: หน้าเพจจะคัดเฉพาะ Brand Code = "Enrich" ให้อัตโนมัติ
   (ตั้งค่าได้ที่ตัวแปร BRAND_FILTER ใน stock-analysis.html)
   ===================================================================== */

/* รูปแบบที่รองรับ: ใส่ได้ทั้ง
   (ก) อาร์เรย์ตรง ๆ  [ {...}, {...} ]
   (ข) ผลลัพธ์ OData ดิบ  { "value": [ {...}, {...} ] }   <-- วางทั้งก้อนที่ได้จาก URL ได้เลย
   ชื่อฟิลด์รองรับทั้งแบบ OData web service (No, Description, Brand_Code)
   และแบบ API v2.0 (number, displayName) — ตัวอ่านจะ map ให้เอง
*/
window.BC_ITEMS = [
  {
    "No": "ENR-0001",
    "Description": "Enrich Vitamin C 1000mg (60s)",
    "Base_Unit_of_Measure": "BOTTLE",
    "Brand_Code": "ENRICH",                 // <-- custom field No.80001 (Code[20]) — ค่า 'ENRICH'
    "Lead_Time_Calculation": "35D",         // BC DateFormula เช่น 35D / 6W / 2M — ระบบแปลงเป็นวันให้
    "Safety_Stock_Quantity": 600,
    "Vendor_No": "V001"
  },
  {
    "No": "ENR-0002",
    "Description": "Enrich Marine Collagen 200g",
    "Base_Unit_of_Measure": "JAR",
    "Brand_Code": "Enrich",
    "Lead_Time_Calculation": "50D",
    "Safety_Stock_Quantity": 500,
    "Vendor_No": "V002"
  },
  {
    "No": "OTH-9001",
    "Description": "สินค้าแบรนด์อื่น (จะถูกกรองออก)",
    "Base_Unit_of_Measure": "PCS",
    "Brand_Code": "OtherBrand",
    "Lead_Time_Calculation": "20D",
    "Safety_Stock_Quantity": 100,
    "Vendor_No": "V004"
  }
];

/* =====================================================================
   วิธีดึงข้อมูล Item จาก BC (เลือกทางใดทางหนึ่ง)
   ---------------------------------------------------------------------
   ทาง A — OData ผ่านเบราว์เซอร์ (ง่ายสุด ไม่ต้องตั้ง Azure App):
     1. ใน BC: Search "Web Services" → New → Page 30 (Item Card) หรือ Page 31 (Item List)
        ตั้ง Service Name เช่น "ItemsEnrich" → ติ๊ก Published
        * ถ้า Brand Code เป็นฟิลด์ที่ไม่ได้อยู่ในเพจมาตรฐาน อาจต้องใช้ Page ที่มีฟิลด์นี้
     2. copy คอลัมน์ OData V4 URL แล้วเติม filter ท้าย URL:
        .../ODataV4/Company('RICHEST SUPP CO., LTD')/ItemsEnrich?$filter=Brand_Code eq 'ENRICH'
        (Brand_Code = custom field No.80001; ถ้า OData ตั้งชื่อฟิลด์ต่างไป ให้ดูชื่อคอลัมน์ใน metadata)
     3. เปิด URL นั้นในเบราว์เซอร์ (ระบบให้ล็อกอินบัญชีองค์กร) → จะได้ JSON { "value": [...] }
     4. ก๊อป JSON ทั้งก้อนมาวางแทนค่า window.BC_ITEMS ข้างบน (วาง { "value": [...] } ได้เลย)

   ทาง B — Export ผ่าน Excel:
     เปิด Item List ใน BC → กรอง Brand Code = Enrich → Open in Excel → save
     แล้วบอกผม ผมช่วยแปลงเป็นไฟล์ bc-items.js ให้
   ===================================================================== */
