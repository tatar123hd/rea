# ส่วนที่ 2 — สร้างลิสต์ SharePoint ใหม่

มี 2 ทางเลือก ทำทางใดทางหนึ่งพอ

- **(ก) อัตโนมัติ** ด้วยสคริปต์ `sharepoint/Provision-ProjectPortfolio.ps1` — เร็ว แม่นยำ ได้ข้อมูลตัวอย่างด้วย
- **(ข) ทำมือทีละคลิก** — ช้ากว่า แต่ไม่ต้องติดตั้งอะไรและเห็นทุกขั้นตอน

---

## 2.1 โครงสร้างลิสต์ทั้งหมด (ใช้ร่วมกันทั้งสองวิธี)

> กฎการตั้งชื่อ: **ชื่อลิสต์และ internal name ของคอลัมน์เป็นภาษาอังกฤษ ไม่มีช่องว่าง ไม่มีอักขระไทย**
> ส่วน **display name** เป็นภาษาไทยได้ (ผู้ใช้เห็นชื่อไทย แต่สูตรอ้างชื่ออังกฤษ)
>
> ⚠️ ข้อควรระวังใหญ่ที่สุดตอนทำมือ: ถ้าสร้างคอลัมน์โดยพิมพ์ชื่อไทยลงไปก่อน SharePoint จะตั้ง
> internal name เป็นรหัสแปลก ๆ เช่น `_x0e0a_...` แล้วสูตรทั้งหมดจะพัง
> **ให้พิมพ์ชื่ออังกฤษตอนสร้างเสมอ แล้วค่อยเปลี่ยน display name เป็นไทยทีหลัง** (ขั้นตอนอยู่ใน §2.3)

### ลิสต์ `Projects`

| Internal name | ชื่อที่แสดง | ชนิด | จำเป็น | Index | หมายเหตุ |
|---|---|---|---|---|---|
| `Title` | ชื่อโปรเจกต์ | Single line of text | ✔ | | คอลัมน์ติดมากับลิสต์ แค่เปลี่ยนชื่อที่แสดง |
| `Department` | แผนก | Single line of text | ✔ | ✔ | มาจากโปรไฟล์ของผู้สร้าง |
| `StatusKey` | สถานะ | Single line of text | ✔ | ✔ | ค่า: `NotStarted` `InProgress` `Done` `Blocked` |
| `StartDate` | วันเริ่ม | Date and Time (**Date Only**) | | ✔ | แอปเขียนให้จากเฟสที่เริ่มเร็วที่สุด |
| `EndDate` | วันสิ้นสุด | Date and Time (**Date Only**) | | ✔ | แอปเขียนให้จากเฟสที่จบช้าที่สุด |
| `OwnerEmail` | อีเมลเจ้าของ | Single line of text | ✔ | ✔ | สำเนาของ Created By ไว้ใช้กรองแบบ delegate ได้ |
| `ColorHex` | สีแถบ | Single line of text | | | เช่น `#0064FF` |
| `Notes` | หมายเหตุ | Multiple lines of text (plain) | | | |
| `Author` | Created By | (ติดมากับลิสต์) | | | ผู้สร้างตัวจริงระดับ SharePoint |

### ลิสต์ `Phases`

| Internal name | ชื่อที่แสดง | ชนิด | จำเป็น | Index |
|---|---|---|---|---|
| `Title` | ชื่อเฟส | Single line of text | ✔ | |
| `ProjectID` | รหัสโปรเจกต์ | Number (0 ตำแหน่งทศนิยม) | ✔ | ✔ |
| `ProjectTitle` | ชื่อโปรเจกต์ | Single line of text | | |
| `Department` | แผนก | Single line of text | | ✔ |
| `PhaseType` | ประเภทเฟส | Single line of text | | ✔ |
| `SortOrder` | ลำดับ | Number | | |
| `StartDate` | วันเริ่ม | Date and Time (**Date Only**) | | ✔ |
| `EndDate` | วันสิ้นสุด | Date and Time (**Date Only**) | | ✔ |
| `PersonAssigned` | ผู้รับผิดชอบ | Person or Group | | |
| `PersonEmail` | อีเมลผู้รับผิดชอบ | Single line of text | | ✔ |
| `StatusKey` | สถานะ | Single line of text | | ✔ |
| `OwnerEmail` | อีเมลเจ้าของ | Single line of text | | ✔ |

### ลิสต์ `Tasks`

| Internal name | ชื่อที่แสดง | ชนิด | จำเป็น | Index |
|---|---|---|---|---|
| `Title` | ชื่องาน | Single line of text | ✔ | |
| `Description` | รายละเอียดงาน | Multiple lines of text (plain, 6 บรรทัด) | | |
| `ProjectID` | รหัสโปรเจกต์ | Number | ✔ | ✔ |
| `ProjectTitle` | ชื่อโปรเจกต์ | Single line of text | | |
| `PhaseID` | รหัสเฟส | Number | ✔ | ✔ |
| `PhaseTitle` | ชื่อเฟส | Single line of text | | |
| `Department` | แผนก | Single line of text | | ✔ |
| `AssignedTo` | ผู้รับผิดชอบ | Person or Group | | |
| `AssignedToEmail` | อีเมลผู้รับผิดชอบ | Single line of text | | ✔ |
| `StartDate` | วันเริ่ม | Date and Time (**Date Only**) | | ✔ |
| `DueDate` | กำหนดส่ง | Date and Time (**Date Only**) | | ✔ |
| `StatusKey` | สถานะ | Single line of text | ✔ | ✔ |
| `PriorityKey` | ความสำคัญ | Single line of text | | ✔ |
| `OwnerEmail` | อีเมลเจ้าของ | Single line of text | | ✔ |

### ลิสต์ `TaskTemplates`

| Internal name | ชื่อที่แสดง | ชนิด | จำเป็น | Index |
|---|---|---|---|---|
| `Title` | ชื่องานที่แนะนำ | Single line of text | ✔ | |
| `PhaseType` | ประเภทเฟส | Single line of text | ✔ | ✔ |
| `SortOrder` | ลำดับ | Number | | |
| `HintText` | คำอธิบาย | Multiple lines of text (plain) | | |

### ลิสต์ `SkillsInterests`

| Internal name | ชื่อที่แสดง | ชนิด | จำเป็น | Index |
|---|---|---|---|---|
| `Title` | ชื่อทักษะ/ความสนใจ | Single line of text | ✔ | |
| `PersonUser` | บุคคล | Person or Group | | |
| `PersonEmail` | อีเมล | Single line of text | ✔ | ✔ |
| `Department` | แผนก | Single line of text | | ✔ |
| `ItemKind` | ประเภท | Single line of text | ✔ | ✔ |
| `Level` | ระดับ | Single line of text | | |

### ลิสต์ `AppAdmins`

| Internal name | ชื่อที่แสดง | ชนิด | จำเป็น | Index |
|---|---|---|---|---|
| `Title` | ชื่อผู้ดูแล | Single line of text | ✔ | |
| `PersonUser` | บุคคล | Person or Group | | |
| `Email` | อีเมล | Single line of text | ✔ | ✔ |

### ทำไม `StatusKey` เป็น Text ไม่ใช่ Choice

- **ความแน่นอนของ delegation** — การเทียบ Choice ต้องใช้ `Status.Value` ซึ่งเป็น sub-field
  SharePoint connector รองรับเฉพาะบางรูปแบบ ส่วน Text ที่ index แล้วใช้ `=` ได้แน่นอน
- **ไม่มีค่าว่างแปลก ๆ** — Choice ที่ลบตัวเลือกทิ้งจะทิ้งค่าเก่าค้างไว้
- **ข้อเสียที่ยอมรับ** — SharePoint ไม่ได้บังคับค่าให้ ถ้ามีคนพิมพ์มั่วผ่านหน้าลิสต์ตรง ๆ
  แอปจะแสดงเป็น "ไม่ระบุ" แทนที่จะพัง ถ้าต้องการบังคับจริงจัง ให้ใส่ **Column validation**
  ที่คอลัมน์: `=OR(StatusKey="NotStarted",StatusKey="InProgress",StatusKey="Done",StatusKey="Blocked")`
  พร้อมข้อความไทยว่า "สถานะต้องเป็นค่าใดค่าหนึ่งที่ระบบกำหนด"
- ผู้ใช้ **ไม่เคยเห็นค่าอังกฤษพวกนี้** — แอปแปลงเป็นไทยทั้งหมด (ยังไม่เริ่ม / กำลังดำเนินการ / เสร็จแล้ว / ติดปัญหา)

### ค่าคงที่ (ใช้ตรงกันทั้งระบบ)

| ชนิด | ค่าในฐานข้อมูล | ข้อความไทยที่ผู้ใช้เห็น |
|---|---|---|
| StatusKey | `NotStarted` | ยังไม่เริ่ม |
| StatusKey | `InProgress` | กำลังดำเนินการ |
| StatusKey | `Done` | เสร็จแล้ว |
| StatusKey | `Blocked` | ติดปัญหา |
| PriorityKey | `High` / `Medium` / `Low` | สูง / ปานกลาง / ต่ำ |
| ItemKind | `Skill` / `Interest` | ทักษะ / ความสนใจ |
| PhaseType | `Plan` `Design` `Develop` `Test` `Launch` `Evaluate` | วางแผน / ออกแบบ / พัฒนา / ทดสอบ / เริ่มใช้งาน / ประเมินผล |

---

## 2.2 วิธี (ก) — รันสคริปต์ PnP PowerShell

### เตรียมเครื่องครั้งเดียว

1. เปิด **PowerShell 7** (ไม่ใช่ Windows PowerShell 5.1 เพราะ PnP.PowerShell รุ่นใหม่ต้องใช้ PS7)
2. ติดตั้งโมดูล:
   ```powershell
   Install-Module PnP.PowerShell -Scope CurrentUser
   ```
3. ลงทะเบียนแอปสำหรับล็อกอิน (ทำครั้งเดียวต่อ tenant — ต้องมีสิทธิ์ให้ consent):
   ```powershell
   Register-PnPEntraIDAppForInteractiveLogin `
       -ApplicationName "PnP PowerShell" `
       -Tenant contoso.onmicrosoft.com
   ```
   จดค่า **client id** ที่ได้ไว้

> ⚠️ ข้อนี้เปลี่ยนบ่อย: ตั้งแต่ปลายปี 2024 PnP.PowerShell **ไม่มีแอปกลางให้ใช้ฟรีแล้ว**
> จึงต้องระบุ `-ClientId` เสมอ และชื่อ cmdlet เคยเป็น `Register-PnPAzureADApp`
> ถ้ารันแล้วไม่เจอคำสั่ง ให้ดูขั้นตอนล่าสุดที่ https://pnp.github.io/powershell/

### รันสคริปต์

```powershell
.\Provision-ProjectPortfolio.ps1 `
    -SiteUrl  "https://contoso.sharepoint.com/sites/ProjectPortfolio" `
    -ClientId "00000000-1111-2222-3333-444444444444" `
    -AdminUpn "somchai@contoso.com"
```

เพิ่ม `-SkipSeedData` ถ้าไม่ต้องการข้อมูลตัวอย่าง

สคริปต์ **รันซ้ำได้** ของที่มีอยู่แล้วจะถูกข้าม ไม่มีการลบข้อมูล

### ตรวจผลหลังรัน

1. เปิด site → เมนูซ้ายต้องเห็นลิสต์ครบ 6 ลิสต์
2. `Projects` → Settings → **Indexed columns** ต้องเห็น Department, StatusKey, StartDate, EndDate, OwnerEmail
3. `Projects` → Settings → **Advanced settings** → Item-level permissions
   ต้องเป็น *Read all items* + *Create items and edit items that were created by the user*
4. มีข้อมูลตัวอย่าง 3 โปรเจกต์ และ TaskTemplates 30 รายการ

---

## 2.3 วิธี (ข) — ทำมือทีละคลิก

ทำซ้ำตามรอบนี้สำหรับทุกลิสต์ในตาราง §2.1 ตัวอย่างด้านล่างใช้ `Projects` เป็นแบบ

### ขั้นที่ 1 — สร้างลิสต์

1. เปิด SharePoint site ที่ต้องการ
2. กด **+ New** ที่แถบด้านบน → เลือก **List**
3. เลือก **Blank list**
4. ช่อง **Name** พิมพ์ `Projects` (ภาษาอังกฤษ ไม่มีช่องว่าง)
5. ช่อง **Description** พิมพ์ `โปรเจกต์หลัก`
6. เปิดสวิตช์ **Show in site navigation**
7. กด **Create**

### ขั้นที่ 2 — เพิ่มคอลัมน์ (ทำซ้ำทุกคอลัมน์ในตาราง)

1. ในหน้าลิสต์ กด **+ Add column**
2. เลือกชนิดตามตาราง:
   - *Single line of text* → **Text**
   - *Number* → **Number**
   - *Date and Time* → **Date and time**
   - *Person or Group* → **Person**
   - *Multiple lines of text* → **Multiple lines of text**
3. ช่อง **Name** พิมพ์ **ชื่ออังกฤษจากคอลัมน์ "Internal name"** เช่น `Department`
4. ถ้าเป็นคอลัมน์วันที่ ให้เลือก **Include time → No** (หรือ **Date only**)
5. ถ้าเป็น Number ให้ตั้ง **Number of decimal places → 0**
6. ถ้าเป็น Multiple lines of text ให้เลือก **Use enhanced rich text → No**
7. ถ้าตารางกำกับว่า "จำเป็น" ให้กด **More options** → **Require that this column contains information → Yes**
8. กด **Save**

### ขั้นที่ 3 — เปลี่ยนชื่อที่แสดงเป็นภาษาไทย (ทำหลังสร้างครบแล้วเท่านั้น)

1. คลิกหัวคอลัมน์ → **Column settings** → **Edit**
2. แก้ช่อง **Name** เป็นภาษาไทยตามตาราง เช่น `แผนก`
3. กด **Save**

> การเปลี่ยนชื่อที่แสดงภายหลัง **ไม่** เปลี่ยน internal name ซึ่งเป็นสิ่งที่สูตรใช้ — ปลอดภัย
> ตรวจ internal name ได้ที่ List settings → คลิกชื่อคอลัมน์ → ดูท้าย URL ที่ `Field=Department`

### ขั้นที่ 4 — สร้าง index

ทำกับทุกคอลัมน์ที่ตาราง §2.1 ติ๊ก ✔ ในช่อง Index

1. ฟันเฟือง (Settings) มุมขวาบน → **List settings**
2. เลื่อนลงไปใต้ตาราง Columns → กด **Indexed columns**
3. กด **Create a new index**
4. **Primary column for this index** เลือกคอลัมน์ที่ต้องการ
5. กด **Create**
6. ทำซ้ำจนครบ

> ถ้าปุ่ม **Indexed columns** หาไม่เจอ: บาง tenant ย้ายไปอยู่ที่
> **List settings → Advanced settings** หรือซ่อนไว้หลัง **More library settings**
> ค้นคำว่า "Add an index to a SharePoint list column" ใน Microsoft Learn เพื่อดูขั้นตอนล่าสุด
>
> ⚠️ **ต้องสร้าง index ตั้งแต่ลิสต์ยังเล็ก** พอเกิน 5,000 รายการแล้วจะสร้าง index ไม่ได้ในเวลาทำการ

### ขั้นที่ 5 — สิทธิ์ระดับรายการ (ทำเฉพาะ Projects, Phases, Tasks)

1. **List settings** → **Advanced settings**
2. หัวข้อ **Item-level Permissions**
   - *Read access* → เลือก **Read all items**
   - *Create and Edit access* → เลือก **Create items and edit items that were created by the user**
3. เลื่อนลงล่างสุด กด **OK**

### ขั้นที่ 6 — สิทธิ์ลิสต์ของผู้ดูแล (ทำเฉพาะ SkillsInterests, AppAdmins, TaskTemplates)

1. สร้างกลุ่มก่อน: **Site settings** → **Site permissions** → **Advanced permissions settings**
   → **Create Group** → ชื่อ `PP_Admins` → สิทธิ์ **Edit** → **Create** → เพิ่มสมาชิก
2. ไปที่ลิสต์ → **List settings** → **Permissions for this list**
3. กด **Stop Inheriting Permissions** → ยืนยัน **OK**
4. ติ๊กกลุ่ม **(ชื่อ site) Members** → **Edit User Permissions** → ติ๊ก **Read** เท่านั้น → **OK**
5. กด **Grant Permissions** → ใส่ `PP_Admins` → **Show options** → เลือก **Edit** → **Share**

> **สำคัญ:** คนที่เป็นผู้ดูแลแอปต้องอยู่ **ทั้ง** ในกลุ่ม `PP_Admins` **และ** ในลิสต์ `AppAdmins`
> ลิสต์ `AppAdmins` คุมแค่ว่าเห็นเมนูผู้ดูแลในแอปหรือไม่ ส่วนกลุ่มคุมสิทธิ์เขียนจริงที่ SharePoint

### ขั้นที่ 7 — ตั้ง time zone ของ site

1. **Site settings** → **Regional settings**
2. **Time Zone** → เลือก **(UTC+07:00) Bangkok, Hanoi, Jakarta**
3. **OK**

### ขั้นที่ 8 — ใส่ข้อมูล TaskTemplates

ลิสต์นี้ต้องมีข้อมูลถึงจะมีงานแนะนำให้เลือก ใส่ด้วยมือผ่าน **Edit in grid view** จะเร็วที่สุด
(รายการทั้ง 30 บรรทัดอยู่ในสคริปต์ PnP ส่วน `$templates` คัดลอกข้อความไทยมาใช้ได้เลย)

1. เปิดลิสต์ `TaskTemplates` → กด **Edit in grid view**
2. พิมพ์ทีละบรรทัด: ชื่องานที่แนะนำ / ประเภทเฟส (`Plan`, `Design`, `Develop`, `Test`, `Launch`, `Evaluate`) / ลำดับ
3. กด **Exit grid view**
