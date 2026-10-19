# วิธีสร้างคอมโพเนนต์ (อ่านก่อนวาง YAML)

## ข้อควรรู้เรื่องการวาง YAML

Power Apps Studio รองรับการ **คัดลอก–วางคอนโทรลเป็น YAML** แต่รูปแบบชื่อชนิดคอนโทรล
(`Control: Classic/Button` ฯลฯ) **เปลี่ยนไปตามเวอร์ชันของ Studio** ไฟล์ในโฟลเดอร์นี้ใช้รูปแบบ
ที่ใช้อยู่ในปัจจุบัน ถ้าวางแล้วขึ้นว่า "ไม่รู้จักชนิดคอนโทรล" ให้ทำดังนี้

1. สร้างคอนโทรลชนิดนั้น 1 ตัวด้วยมือใน Studio
2. เลือกมัน แล้วกด **Ctrl+C**
3. วางลงใน Notepad จะเห็น YAML ของจริง
4. ดูบรรทัด `Control:` ว่าเขียนว่าอะไร แล้วแก้ไฟล์ให้ตรง (แก้ครั้งเดียวใช้ได้ทั้งไฟล์)

**คอมโพเนนต์โดยเฉพาะ** มักวาง YAML ตรง ๆ ไม่ได้ในหลายเวอร์ชัน ให้สร้างโครงด้วยมือตามตาราง
custom properties ที่อยู่หัวไฟล์แต่ละอัน แล้วค่อยวาง YAML เฉพาะส่วน `Children:` เข้าไปในคอมโพเนนต์

## การตั้งค่าที่ "ต้อง" เปิด

คอมโพเนนต์ `cmpSidebar` และ `cmpFilters` ต้องเปิด **Access app scope**
(เลือกคอมโพเนนต์ใน Tree view → แท็บ **Properties** ทางขวา → สวิตช์ **Access app scope** = On)

ถ้าไม่เปิด คอมโพเนนต์จะมองไม่เห็นตัวแปร global (`gbl…`), collection (`col…`) และ data source
ทำให้ทุกสูตรข้างในพัง

> ข้อจำกัด: ถ้าเอาคอมโพเนนต์ไปไว้ใน **component library** สวิตช์นี้จะใช้ไม่ได้
> โปรเจกต์นี้จึงให้เก็บคอมโพเนนต์ไว้ **ในแอปเดียวกัน** ไม่ต้องทำ library

## รูปแบบการนำทาง (สำคัญ)

คอมโพเนนต์เรียก `Navigate()` เองไม่ได้ เราจึงใช้วิธีนี้

1. ปุ่มในแถบเมนูตั้ง `OnSelect` = `Set(gblNavTarget, "Dashboard"); OnNav()`
2. `OnNav` เป็น custom property ชนิด **Event / Behavior (Input)** ของคอมโพเนนต์
3. ทุกหน้าจอที่วางแถบเมนู ตั้งค่าคุณสมบัติ `OnNav` ของอินสแตนซ์เป็น:

```powerappsfl
Switch(gblNavTarget,
    "Timeline",  Navigate(scrTimeline),
    "Dashboard", Navigate(scrDashboard),
    "Board",     Navigate(scrBoard),
    "Table",     Navigate(scrTable),
    "Calendar",  Navigate(scrCalendar),
    "MyTasks",   Navigate(scrMyTasks),
    "Skills",    Navigate(scrSkills),
    "Help",      Navigate(scrHelp),
    "Admin",     Navigate(scrAdmin)
)
```

> ถ้า environment ของคุณไม่มี custom property ชนิด Behavior ให้ทำแบบสำรอง:
> ตั้ง `OnSelect` ของปุ่มเป็น `Set(gblNavTarget, "Dashboard")` เฉย ๆ แล้วใส่สูตร `Switch` ข้างบน
> ไว้ที่ **OnVisible ของหน้าจอ** ไม่ได้ — วิธีสำรองที่ใช้ได้จริงคือ **ไม่ทำแถบเมนูเป็นคอมโพเนนต์**
> แต่สร้างเป็นกลุ่มคอนโทรลธรรมดาในหน้าจอแรก แล้ว Ctrl+C / Ctrl+V ไปทุกหน้าจอ
> (เสียความสะดวกตอนแก้ไข แต่ทำงานได้ทุกเวอร์ชัน)
