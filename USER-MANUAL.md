# PROJECT DOCUMENT GENERATOR
## system requirement
- หน้าจอใช้ html
- หน้าบ้านใช้ js
- หลังบ้านใช้  java ที่มี jasper compiler
- database ใช้ sqlite  
- เป็น standalone โปรแกรม

## วิธีการทำงาน
### menu upload template
1.เมื่ออัปโหลดไฟล์ excel template จะถูกแปลงเป็นเนื้อหาของ jrml
2.ใส่รายละเอียดเกี่ยวกับข้อมูลที่จะบันทึกไปใน jrml 

### menu generate document
1.ให้ผู้ใช้เลือก template ที่จะบันทึกไปใน jrml
2.ให้ผู้ใช้กดปุ่มอัปโหลด excel ที่เป็นข้อมูล
3.ผู้ใช้กด preview PDF หรือ export excel โปรแกรมส่ง ข้อมูล และไปดึง jrxml template ที่ผู้ใช้เลือก มา render เป็นไฟล์ที่ต้องการให้ 


โดยการแปลง excel เป็น jrxml ให้ใช้ฟังก์ชันเดิมที่มีในโปรเจคนี้
