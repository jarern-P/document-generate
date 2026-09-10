# การรันโปรแกรม

โปรแกรมนี้ใช้ Maven Wrapper (`mvnw`) เพื่อดาวน์โหลด Maven ให้อัตโนมัติ (ถ้ายังไม่มี)

---

## วิธีที่ 1 — รันโดยตรง (ระหว่างพัฒนา)

```bash
./mvnw clean package
```

สร้างไฟล์ JAR ที่ `target/document-generator.jar`

```bash
java -jar target/document-generator.jar
```

หน้าบ้านจะเปิดขึ้นที่ `http://127.0.0.1:8080/` (จะเปิด браузерอัตโนมัติ หรือจะเปิดเองก็ได้)

ถ้าอยากเลือกพอร์ต หรือเปลี่ยนโฟลเดอร์เก็บข้อมูล:

```bash
java -jar target/document-generator-1.0.0.jar --port=9000 --data=/path/to/data
```

หรือตั้งตัวแปรสิ่งแวดล้อม:

```bash
export DOCGEN_PORT=9000
java -jar target/document-generator-1.0.0.jar
```

---

## วิธีที่ 2 — สร้างไฟล์ JAR ที่รวม dependency แล้ว (สำหรับแจกจ่าย)

```bash
./mvnw clean package
```

จะได้ไฟล์:

```
target/document-generator-1.0.0.jar
```

รันด้วยคำสั่งเดียวกับวิธีที่ 1:

```bash
java -jar target/document-generator-1.0.0.jar
```

ไฟล์ JAR นี้ไม่ต้องใช้ Maven หรือโฟลเดอร์ `target` ร่วม สามารถ copy ไปรันบนเครื่องอื่นได้ (ต้องมี Java 21 ขึ้นไป)

---

## ข้อกำหนด

- **Java 21** (ตรวจสอบด้วย `java -version`)
- **Git / Maven wrapper** — ใช้ `./mvnw` จะดาวน์โหลด Maven ให้อัตโนมัติ

ตรวจสอบ Java:

```bash
java -version
```

ควรแสดงข้อความว่า `21.x.x` (หรืออย่างน้อยก็ต้องเป็น Java 21)

---

## โครงสร้างข้อมูล

- โฟลเดอร์ข้อมูลเริ่มต้นอยู่ที่ `~/.document-generator/`
- ไฟล์ฐานข้อมูลอยู่ที่ `~/.document-generator/docgen.db`
- เปลี่ยนโฟลเดอร์ได้ด้วย `--data=/path/to/data`

---

## การใช้งานเบื้องต้น

1. เปิด `http://127.0.0.1:8080/`
2. **อัปโหลดเทมเพลต**: ลากไฟล์ `.xlsx` ที่เป็น template เข้ามา
3. **บันทึกเทมเพลต**: ตั้งชื่อและบันทึก
4. **สร้างเอกสาร**: เลือกเทมเพลต, อัปโหลดไฟล์ข้อมูล `.xlsx`, กดสร้าง PDF หรือ Excel

---

## การปิดโปรแกรม

กด `Ctrl+C` ใน terminal

## การสร้าง installer

```bash
.\build-app.ps1
```


สร้างไฟล์ installer ในโฟลเดอร์ `dist`