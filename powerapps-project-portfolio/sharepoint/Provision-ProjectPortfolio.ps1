<#
.SYNOPSIS
    สร้างลิสต์ SharePoint Online ทั้งหมดสำหรับแอป Project Portfolio (Projects / Phases / Tasks /
    TaskTemplates / SkillsInterests / AppAdmins) พร้อมคอลัมน์ index สิทธิ์ และข้อมูลตัวอย่าง

.DESCRIPTION
    สคริปต์นี้ทำงานแบบ idempotent: รันซ้ำได้ ถ้ามีลิสต์/คอลัมน์อยู่แล้วจะข้าม ไม่ลบของเดิม
    สคริปต์นี้ "ไม่แตะ" ลิสต์ BAU (L&D TMP) ใด ๆ ทั้งสิ้น

.PARAMETER SiteUrl
    URL ของ SharePoint site ที่จะสร้างลิสต์ใหม่ เช่น https://contoso.sharepoint.com/sites/ProjectPortfolio

.PARAMETER ClientId
    Entra ID (Azure AD) application id ที่ใช้ล็อกอิน PnP PowerShell
    ตั้งแต่ PnP.PowerShell 2.x เป็นต้นไป "ต้อง" ระบุค่านี้ (แอป PnP Management Shell แบบ multi-tenant
    ถูกปิดไปตั้งแต่กันยายน 2024) ถ้ายังไม่มี ให้รันครั้งเดียว:
        Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP Rocks" -Tenant contoso.onmicrosoft.com
    แล้วนำ client id ที่ได้มาใส่
    หมายเหตุ: ชื่อ cmdlet นี้เคยชื่อ Register-PnPAzureADApp ในเวอร์ชันเก่า ถ้าไม่เจอให้ตรวจที่
    https://pnp.github.io/powershell/

.PARAMETER AdminGroupName
    ชื่อ SharePoint group ที่จะมีสิทธิ์ Edit บน SkillsInterests / AppAdmins / TaskTemplates

.PARAMETER AdminUpn
    อีเมล (UPN) ของผู้ดูแลแอปคนแรก จะถูกใส่ทั้งใน SharePoint group และลิสต์ AppAdmins

.PARAMETER SkipSeedData
    ข้ามการใส่ข้อมูลตัวอย่าง (ใช้ตอน deploy ขึ้น production)

.EXAMPLE
    .\Provision-ProjectPortfolio.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/PP" `
        -ClientId "00000000-1111-2222-3333-444444444444" -AdminUpn "somchai@contoso.com"

.NOTES
    ต้องมี: PnP.PowerShell  (Install-Module PnP.PowerShell -Scope CurrentUser)
    สิทธิ์ที่ต้องมี: Site Owner (Full Control) ของ site ปลายทาง
    สคริปต์นี้ยังไม่เคยถูกรันจริงโดยผู้เขียนเอกสาร — กรุณาทดสอบใน site ทดสอบก่อนเสมอ
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string] $SiteUrl,
    [Parameter(Mandatory = $true)][string] $ClientId,
    [Parameter(Mandatory = $true)][string] $AdminUpn,
    [string] $AdminGroupName = "PP_Admins",
    [switch] $SkipSeedData
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

#region ---------- helper functions (idempotent) ----------

function Write-Step { param([string]$Text) Write-Host "`n=== $Text ===" -ForegroundColor Cyan }
function Write-Skip { param([string]$Text) Write-Host "    [ข้าม] $Text" -ForegroundColor DarkGray }
function Write-Done { param([string]$Text) Write-Host "    [ทำแล้ว] $Text" -ForegroundColor Green }

function Ensure-List {
    param([string]$Title, [string]$Description)
    $list = Get-PnPList -Identity $Title -ErrorAction SilentlyContinue
    if ($null -ne $list) { Write-Skip "ลิสต์ $Title มีอยู่แล้ว"; return $list }
    $list = New-PnPList -Title $Title -Template GenericList -OnQuickLaunch
    Set-PnPList -Identity $Title -Description $Description
    Write-Done "สร้างลิสต์ $Title"
    return $list
}

function Ensure-Field {
    <#  Type: Text | Note | Number | DateTime | User | Choice | Boolean  #>
    param(
        [string]$List, [string]$InternalName, [string]$DisplayName, [string]$Type,
        [string[]]$Choices = @(), [switch]$Required, [switch]$Indexed, [switch]$DateOnly,
        [switch]$AddToDefaultView
    )
    $existing = Get-PnPField -List $List -Identity $InternalName -ErrorAction SilentlyContinue
    if ($null -eq $existing) {
        $params = @{
            List         = $List
            DisplayName  = $DisplayName
            InternalName = $InternalName
            Type         = $Type
        }
        if ($AddToDefaultView) { $params["AddToDefaultView"] = $true }
        if ($Type -eq "Choice" -and $Choices.Count -gt 0) { $params["Choices"] = $Choices }
        Add-PnPField @params | Out-Null
        Write-Done "เพิ่มคอลัมน์ $List.$InternalName ($Type)"
    }
    else { Write-Skip "คอลัมน์ $List.$InternalName มีอยู่แล้ว" }

    # ค่าประกอบของคอลัมน์ ตั้งซ้ำได้ไม่มีผลข้างเคียง
    $values = @{}
    if ($Required)  { $values["Required"] = $true }
    if ($Indexed)   { $values["Indexed"]  = $true }
    if ($DateOnly)  { $values["DisplayFormat"] = 0 }   # 0 = DateOnly, 1 = DateTime
    if ($Type -eq "Note") { $values["RichText"] = $false; $values["NumberOfLines"] = 6 }
    if ($values.Count -gt 0) {
        Set-PnPField -List $List -Identity $InternalName -Values $values | Out-Null
    }
}

function Ensure-TitleRename {
    param([string]$List, [string]$NewDisplayName)
    Set-PnPField -List $List -Identity "Title" -Values @{ Title = $NewDisplayName } | Out-Null
}

#endregion

#region ---------- connect ----------

Write-Step "เชื่อมต่อ SharePoint"
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
$web = Get-PnPWeb
Write-Done "เชื่อมต่อ $($web.Title) แล้ว"

#endregion

#region ---------- 1. Projects ----------

Write-Step "ลิสต์ Projects"
Ensure-List -Title "Projects" -Description "โปรเจกต์หลัก (1 แถว = 1 โปรเจกต์)" | Out-Null
Ensure-TitleRename -List "Projects" -NewDisplayName "ชื่อโปรเจกต์"

Ensure-Field -List "Projects" -InternalName "Department"  -DisplayName "แผนก"        -Type Text     -Required -Indexed -AddToDefaultView
Ensure-Field -List "Projects" -InternalName "StatusKey"   -DisplayName "สถานะ"       -Type Text     -Required -Indexed -AddToDefaultView
Ensure-Field -List "Projects" -InternalName "StartDate"   -DisplayName "วันเริ่ม"     -Type DateTime -Indexed -DateOnly -AddToDefaultView
Ensure-Field -List "Projects" -InternalName "EndDate"     -DisplayName "วันสิ้นสุด"   -Type DateTime -Indexed -DateOnly -AddToDefaultView
Ensure-Field -List "Projects" -InternalName "OwnerEmail"  -DisplayName "อีเมลเจ้าของ" -Type Text     -Required -Indexed -AddToDefaultView
Ensure-Field -List "Projects" -InternalName "ColorHex"    -DisplayName "สีแถบ"        -Type Text
Ensure-Field -List "Projects" -InternalName "Notes"       -DisplayName "หมายเหตุ"     -Type Note

# สิทธิ์ระดับรายการ: อ่านได้ทุกแถว แก้ได้เฉพาะแถวที่ตัวเองสร้าง
Set-PnPList -Identity "Projects" -ReadSecurity 1 -WriteSecurity 2
Write-Done "ตั้งสิทธิ์ Projects = อ่านทุกแถว / แก้เฉพาะของตัวเอง"

#endregion

#region ---------- 2. Phases ----------

Write-Step "ลิสต์ Phases"
Ensure-List -Title "Phases" -Description "เฟสของโปรเจกต์ (ลูกของ Projects ผ่าน ProjectID)" | Out-Null
Ensure-TitleRename -List "Phases" -NewDisplayName "ชื่อเฟส"

Ensure-Field -List "Phases" -InternalName "ProjectID"    -DisplayName "รหัสโปรเจกต์"  -Type Number   -Required -Indexed -AddToDefaultView
Ensure-Field -List "Phases" -InternalName "ProjectTitle" -DisplayName "ชื่อโปรเจกต์"  -Type Text     -AddToDefaultView
Ensure-Field -List "Phases" -InternalName "Department"   -DisplayName "แผนก"          -Type Text     -Indexed
Ensure-Field -List "Phases" -InternalName "PhaseType"    -DisplayName "ประเภทเฟส"     -Type Text     -Indexed -AddToDefaultView
Ensure-Field -List "Phases" -InternalName "SortOrder"    -DisplayName "ลำดับ"          -Type Number   -AddToDefaultView
Ensure-Field -List "Phases" -InternalName "StartDate"    -DisplayName "วันเริ่ม"       -Type DateTime -Indexed -DateOnly -AddToDefaultView
Ensure-Field -List "Phases" -InternalName "EndDate"      -DisplayName "วันสิ้นสุด"     -Type DateTime -Indexed -DateOnly -AddToDefaultView
Ensure-Field -List "Phases" -InternalName "PersonAssigned" -DisplayName "ผู้รับผิดชอบ" -Type User    -AddToDefaultView
Ensure-Field -List "Phases" -InternalName "PersonEmail"  -DisplayName "อีเมลผู้รับผิดชอบ" -Type Text -Indexed
Ensure-Field -List "Phases" -InternalName "StatusKey"    -DisplayName "สถานะ"          -Type Text     -Indexed
Ensure-Field -List "Phases" -InternalName "OwnerEmail"   -DisplayName "อีเมลเจ้าของ"   -Type Text     -Indexed

Set-PnPList -Identity "Phases" -ReadSecurity 1 -WriteSecurity 2
Write-Done "ตั้งสิทธิ์ Phases"

#endregion

#region ---------- 3. Tasks ----------

Write-Step "ลิสต์ Tasks"
Ensure-List -Title "Tasks" -Description "งานย่อยในเฟส (ลูกของ Phases ผ่าน PhaseID)" | Out-Null
Ensure-TitleRename -List "Tasks" -NewDisplayName "ชื่องาน"

Ensure-Field -List "Tasks" -InternalName "Description"     -DisplayName "รายละเอียดงาน" -Type Note
Ensure-Field -List "Tasks" -InternalName "ProjectID"       -DisplayName "รหัสโปรเจกต์"   -Type Number   -Required -Indexed -AddToDefaultView
Ensure-Field -List "Tasks" -InternalName "ProjectTitle"    -DisplayName "ชื่อโปรเจกต์"   -Type Text
Ensure-Field -List "Tasks" -InternalName "PhaseID"         -DisplayName "รหัสเฟส"        -Type Number   -Required -Indexed -AddToDefaultView
Ensure-Field -List "Tasks" -InternalName "PhaseTitle"      -DisplayName "ชื่อเฟส"        -Type Text
Ensure-Field -List "Tasks" -InternalName "Department"      -DisplayName "แผนก"           -Type Text     -Indexed
Ensure-Field -List "Tasks" -InternalName "AssignedTo"      -DisplayName "ผู้รับผิดชอบ"   -Type User     -AddToDefaultView
Ensure-Field -List "Tasks" -InternalName "AssignedToEmail" -DisplayName "อีเมลผู้รับผิดชอบ" -Type Text  -Indexed
Ensure-Field -List "Tasks" -InternalName "StartDate"       -DisplayName "วันเริ่ม"        -Type DateTime -Indexed -DateOnly -AddToDefaultView
Ensure-Field -List "Tasks" -InternalName "DueDate"         -DisplayName "กำหนดส่ง"       -Type DateTime -Indexed -DateOnly -AddToDefaultView
Ensure-Field -List "Tasks" -InternalName "StatusKey"       -DisplayName "สถานะ"          -Type Text     -Required -Indexed -AddToDefaultView
Ensure-Field -List "Tasks" -InternalName "PriorityKey"     -DisplayName "ความสำคัญ"      -Type Text     -Indexed
Ensure-Field -List "Tasks" -InternalName "OwnerEmail"      -DisplayName "อีเมลเจ้าของ"   -Type Text     -Indexed

Set-PnPList -Identity "Tasks" -ReadSecurity 1 -WriteSecurity 2
Write-Done "ตั้งสิทธิ์ Tasks"

#endregion

#region ---------- 4. TaskTemplates ----------

Write-Step "ลิสต์ TaskTemplates"
Ensure-List -Title "TaskTemplates" -Description "ตัวอย่างงานมาตรฐานต่อประเภทเฟส (master data อ่านอย่างเดียว)" | Out-Null
Ensure-TitleRename -List "TaskTemplates" -NewDisplayName "ชื่องานที่แนะนำ"

Ensure-Field -List "TaskTemplates" -InternalName "PhaseType"  -DisplayName "ประเภทเฟส" -Type Text   -Required -Indexed -AddToDefaultView
Ensure-Field -List "TaskTemplates" -InternalName "SortOrder"  -DisplayName "ลำดับ"      -Type Number -AddToDefaultView
Ensure-Field -List "TaskTemplates" -InternalName "HintText"   -DisplayName "คำอธิบาย"   -Type Note

#endregion

#region ---------- 5. SkillsInterests ----------

Write-Step "ลิสต์ SkillsInterests"
Ensure-List -Title "SkillsInterests" -Description "ทักษะและความสนใจรายบุคคล (เฉพาะ admin แก้ได้)" | Out-Null
Ensure-TitleRename -List "SkillsInterests" -NewDisplayName "ชื่อทักษะ/ความสนใจ"

Ensure-Field -List "SkillsInterests" -InternalName "PersonUser"  -DisplayName "บุคคล"     -Type User -AddToDefaultView
Ensure-Field -List "SkillsInterests" -InternalName "PersonEmail" -DisplayName "อีเมล"     -Type Text -Required -Indexed -AddToDefaultView
Ensure-Field -List "SkillsInterests" -InternalName "Department"  -DisplayName "แผนก"      -Type Text -Indexed -AddToDefaultView
Ensure-Field -List "SkillsInterests" -InternalName "ItemKind"    -DisplayName "ประเภท"    -Type Text -Required -Indexed -AddToDefaultView
Ensure-Field -List "SkillsInterests" -InternalName "Level"       -DisplayName "ระดับ"      -Type Text

#endregion

#region ---------- 6. AppAdmins ----------

Write-Step "ลิสต์ AppAdmins"
Ensure-List -Title "AppAdmins" -Description "ผู้ดูแลแอป (ใครอยู่ในลิสต์นี้จะเห็นเมนูผู้ดูแล)" | Out-Null
Ensure-TitleRename -List "AppAdmins" -NewDisplayName "ชื่อผู้ดูแล"

Ensure-Field -List "AppAdmins" -InternalName "PersonUser" -DisplayName "บุคคล" -Type User -AddToDefaultView
Ensure-Field -List "AppAdmins" -InternalName "Email"      -DisplayName "อีเมล" -Type Text -Required -Indexed -AddToDefaultView

#endregion

#region ---------- 7. สิทธิ์ระดับลิสต์สำหรับลิสต์ของผู้ดูแล ----------

Write-Step "กลุ่มผู้ดูแลและสิทธิ์ลิสต์"

$group = Get-PnPGroup -Identity $AdminGroupName -ErrorAction SilentlyContinue
if ($null -eq $group) {
    $group = New-PnPGroup -Title $AdminGroupName -Description "ผู้ดูแลแอป Project Portfolio"
    Set-PnPGroupPermissions -Identity $AdminGroupName -AddRole "Edit"
    Write-Done "สร้างกลุ่ม $AdminGroupName"
}
else { Write-Skip "กลุ่ม $AdminGroupName มีอยู่แล้ว" }

try {
    Add-PnPGroupMember -LoginName $AdminUpn -Group $AdminGroupName -ErrorAction Stop
    Write-Done "เพิ่ม $AdminUpn เข้ากลุ่ม $AdminGroupName"
}
catch { Write-Skip "เพิ่มสมาชิกไม่สำเร็จหรือมีอยู่แล้ว: $($_.Exception.Message)" }

foreach ($adminList in @("SkillsInterests", "AppAdmins", "TaskTemplates")) {
    $l = Get-PnPList -Identity $adminList -Includes HasUniqueRoleAssignments
    if (-not $l.HasUniqueRoleAssignments) {
        Set-PnPList -Identity $adminList -BreakRoleInheritance -CopyRoleAssignments
        Write-Done "ตัดการสืบทอดสิทธิ์ของ $adminList"
    }
    else { Write-Skip "$adminList ตัดการสืบทอดสิทธิ์แล้ว" }

    # ผู้ใช้ทั่วไป (Members) เหลือแค่อ่าน / กลุ่มผู้ดูแลแก้ได้
    $members = Get-PnPGroup -AssociatedMemberGroup -ErrorAction SilentlyContinue
    if ($null -ne $members) {
        Set-PnPListPermission -Identity $adminList -Group $members.Title -RemoveRole "Edit"       -ErrorAction SilentlyContinue
        Set-PnPListPermission -Identity $adminList -Group $members.Title -RemoveRole "Contribute" -ErrorAction SilentlyContinue
        Set-PnPListPermission -Identity $adminList -Group $members.Title -AddRole "Read"          -ErrorAction SilentlyContinue
    }
    Set-PnPListPermission -Identity $adminList -Group $AdminGroupName -AddRole "Edit" -ErrorAction SilentlyContinue
    Write-Done "ตั้งสิทธิ์ $adminList (ทุกคนอ่าน / $AdminGroupName แก้ไข)"
}

Write-Host ""
Write-Host "    *** สำคัญ: ใครที่จะเป็นผู้ดูแลแอป ต้องอยู่ทั้งในลิสต์ AppAdmins และในกลุ่ม $AdminGroupName ***" -ForegroundColor Yellow
Write-Host "    ถ้าอยู่แค่ในลิสต์ จะเห็นเมนูผู้ดูแลแต่กดบันทึกแล้วขึ้น error สิทธิ์" -ForegroundColor Yellow

#endregion

#region ---------- 8. ข้อมูลตัวอย่าง ----------

if ($SkipSeedData) {
    Write-Step "ข้ามข้อมูลตัวอย่าง (-SkipSeedData)"
}
else {
    Write-Step "ข้อมูลตัวอย่าง"

    # --- 8.1 TaskTemplates (MECE ต่อประเภทเฟส) ---
    $templates = @(
        @{ PhaseType = "Plan";     Order = 1; Title = "กำหนดขอบเขตและสิ่งที่ส่งมอบ" },
        @{ PhaseType = "Plan";     Order = 2; Title = "ระบุผู้มีส่วนได้ส่วนเสียและผู้อนุมัติ" },
        @{ PhaseType = "Plan";     Order = 3; Title = "ประเมินทรัพยากรและงบประมาณ" },
        @{ PhaseType = "Plan";     Order = 4; Title = "จัดทำแผนเวลาและจุดตรวจสอบ" },
        @{ PhaseType = "Plan";     Order = 5; Title = "ประเมินความเสี่ยงและแผนสำรอง" },

        @{ PhaseType = "Design";   Order = 1; Title = "รวบรวมความต้องการจากผู้ใช้" },
        @{ PhaseType = "Design";   Order = 2; Title = "ออกแบบโครงสร้างและขั้นตอนทำงาน" },
        @{ PhaseType = "Design";   Order = 3; Title = "จัดทำต้นแบบ (prototype)" },
        @{ PhaseType = "Design";   Order = 4; Title = "ทบทวนแบบกับผู้มีส่วนได้ส่วนเสีย" },
        @{ PhaseType = "Design";   Order = 5; Title = "สรุปแบบและขออนุมัติ" },

        @{ PhaseType = "Develop";  Order = 1; Title = "เตรียมสภาพแวดล้อมและเครื่องมือ" },
        @{ PhaseType = "Develop";  Order = 2; Title = "พัฒนาส่วนหลัก" },
        @{ PhaseType = "Develop";  Order = 3; Title = "พัฒนาส่วนเสริมและการเชื่อมต่อ" },
        @{ PhaseType = "Develop";  Order = 4; Title = "ทบทวนงานภายในทีม" },
        @{ PhaseType = "Develop";  Order = 5; Title = "จัดทำเอกสารประกอบ" },

        @{ PhaseType = "Test";     Order = 1; Title = "จัดทำแผนและกรณีทดสอบ" },
        @{ PhaseType = "Test";     Order = 2; Title = "ทดสอบการทำงานตามฟังก์ชัน" },
        @{ PhaseType = "Test";     Order = 3; Title = "ทดสอบร่วมกับผู้ใช้จริง (UAT)" },
        @{ PhaseType = "Test";     Order = 4; Title = "แก้ไขข้อบกพร่องที่พบ" },
        @{ PhaseType = "Test";     Order = 5; Title = "สรุปผลและขออนุมัติขึ้นใช้งาน" },

        @{ PhaseType = "Launch";   Order = 1; Title = "เตรียมแผนการขึ้นใช้งาน" },
        @{ PhaseType = "Launch";   Order = 2; Title = "สื่อสารและอบรมผู้ใช้" },
        @{ PhaseType = "Launch";   Order = 3; Title = "ขึ้นใช้งานจริง" },
        @{ PhaseType = "Launch";   Order = 4; Title = "เฝ้าระวังหลังขึ้นใช้งาน" },
        @{ PhaseType = "Launch";   Order = 5; Title = "ส่งมอบให้ทีมดูแลประจำ" },

        @{ PhaseType = "Evaluate"; Order = 1; Title = "เก็บตัวชี้วัดผลลัพธ์" },
        @{ PhaseType = "Evaluate"; Order = 2; Title = "รวบรวมความคิดเห็นผู้ใช้" },
        @{ PhaseType = "Evaluate"; Order = 3; Title = "เปรียบเทียบผลกับเป้าหมาย" },
        @{ PhaseType = "Evaluate"; Order = 4; Title = "สรุปบทเรียนที่ได้รับ" },
        @{ PhaseType = "Evaluate"; Order = 5; Title = "เสนอแนวทางปรับปรุงรอบถัดไป" }
    )

    $existingTemplates = Get-PnPListItem -List "TaskTemplates" -PageSize 500
    if ($existingTemplates.Count -eq 0) {
        foreach ($t in $templates) {
            Add-PnPListItem -List "TaskTemplates" -Values @{
                Title     = $t.Title
                PhaseType = $t.PhaseType
                SortOrder = $t.Order
            } | Out-Null
        }
        Write-Done "ใส่ TaskTemplates $($templates.Count) รายการ"
    }
    else { Write-Skip "TaskTemplates มีข้อมูลอยู่แล้ว ($($existingTemplates.Count) รายการ)" }

    # --- 8.2 Projects / Phases / Tasks ตัวอย่าง 3 โปรเจกต์ที่ช่วงเวลาทับกัน ---
    $existingProjects = Get-PnPListItem -List "Projects" -PageSize 100
    if ($existingProjects.Count -eq 0) {

        # ยึดเวลาเที่ยงวันเสมอ กันค่าตกไปวันก่อนหน้าเมื่อแปลงเป็น UTC
        $today = (Get-Date).Date.AddHours(12)
        $dept  = "Learning and Development"

        $projectSeed = @(
            @{ Title = "ระบบติดตามโครงการภายใน"; Color = "#0064FF"; Offset = -30; Length = 120; Status = "InProgress"; Type = @("Plan","Design","Develop","Test") },
            @{ Title = "ปรับปรุงหลักสูตรพนักงานใหม่"; Color = "#00A36C"; Offset = -10; Length = 90;  Status = "InProgress"; Type = @("Plan","Design","Launch") },
            @{ Title = "ย้ายคลังเอกสารขึ้น SharePoint"; Color = "#D97706"; Offset = 10;  Length = 150; Status = "NotStarted"; Type = @("Plan","Develop","Test","Launch","Evaluate") }
        )

        foreach ($p in $projectSeed) {
            $pStart = $today.AddDays($p.Offset)
            $pEnd   = $pStart.AddDays($p.Length)

            $projItem = Add-PnPListItem -List "Projects" -Values @{
                Title      = $p.Title
                Department = $dept
                StatusKey  = $p.Status
                StartDate  = $pStart
                EndDate    = $pEnd
                OwnerEmail = $AdminUpn.ToLower()
                ColorHex   = $p.Color
                Notes      = "ข้อมูลตัวอย่างที่สร้างโดยสคริปต์ provisioning"
            }
            $projId = $projItem.Id
            Write-Done "สร้างโปรเจกต์ '$($p.Title)' (ID = $projId)"

            # แบ่งช่วงเวลาโปรเจกต์ออกเป็นเฟสเท่า ๆ กัน
            $phaseCount  = $p.Type.Count
            $phaseLength = [int]($p.Length / $phaseCount)
            $order = 1

            foreach ($phaseType in $p.Type) {
                $phStart = $pStart.AddDays(($order - 1) * $phaseLength)
                $phEnd   = $phStart.AddDays($phaseLength - 1)

                $phaseItem = Add-PnPListItem -List "Phases" -Values @{
                    Title        = "เฟส $order : $phaseType"
                    ProjectID    = $projId
                    ProjectTitle = $p.Title
                    Department   = $dept
                    PhaseType    = $phaseType
                    SortOrder    = $order
                    StartDate    = $phStart
                    EndDate      = $phEnd
                    PersonEmail  = $AdminUpn.ToLower()
                    StatusKey    = $(if ($phEnd -lt $today) { "Done" } elseif ($phStart -le $today) { "InProgress" } else { "NotStarted" })
                    OwnerEmail   = $AdminUpn.ToLower()
                }
                $phaseId = $phaseItem.Id

                # ดึงงานแนะนำ 3 อันแรกของเฟสประเภทนี้มาเป็นงานตัวอย่าง
                $suggested = $templates | Where-Object { $_.PhaseType -eq $phaseType } | Select-Object -First 3
                $taskIndex = 0
                foreach ($s in $suggested) {
                    $tStart = $phStart.AddDays($taskIndex * 3)
                    $tDue   = $tStart.AddDays(4)
                    Add-PnPListItem -List "Tasks" -Values @{
                        Title           = $s.Title
                        Description     = "งานตัวอย่างในเฟส $phaseType ของโปรเจกต์ $($p.Title)"
                        ProjectID       = $projId
                        ProjectTitle    = $p.Title
                        PhaseID         = $phaseId
                        PhaseTitle      = "เฟส $order : $phaseType"
                        Department      = $dept
                        AssignedToEmail = $AdminUpn.ToLower()
                        StartDate       = $tStart
                        DueDate         = $tDue
                        StatusKey       = $(if ($tDue -lt $today) { "Done" } elseif ($tStart -le $today) { "InProgress" } else { "NotStarted" })
                        PriorityKey     = @("High","Medium","Low")[$taskIndex % 3]
                        OwnerEmail      = $AdminUpn.ToLower()
                    } | Out-Null
                    $taskIndex++
                }
                $order++
            }
        }

        # งานหนึ่งรายการที่เลยกำหนดแล้ว เพื่อทดสอบการไฮไลต์ "เลยกำหนด"
        $firstProj = (Get-PnPListItem -List "Projects" -PageSize 1)[0]
        $firstPhase = (Get-PnPListItem -List "Phases" -PageSize 1)[0]
        Add-PnPListItem -List "Tasks" -Values @{
            Title           = "งานทดสอบสถานะเลยกำหนด"
            ProjectID       = $firstProj.Id
            ProjectTitle    = $firstProj["Title"]
            PhaseID         = $firstPhase.Id
            PhaseTitle      = $firstPhase["Title"]
            Department      = $dept
            AssignedToEmail = $AdminUpn.ToLower()
            StartDate       = $today.AddDays(-20)
            DueDate         = $today.AddDays(-5)
            StatusKey       = "InProgress"
            PriorityKey     = "High"
            OwnerEmail      = $AdminUpn.ToLower()
        } | Out-Null
        Write-Done "ใส่ข้อมูลตัวอย่าง 3 โปรเจกต์พร้อมเฟสและงาน"
    }
    else { Write-Skip "Projects มีข้อมูลอยู่แล้ว ($($existingProjects.Count) รายการ)" }

    # --- 8.3 AppAdmins ---
    $existingAdmins = Get-PnPListItem -List "AppAdmins" -PageSize 100
    if ($existingAdmins.Count -eq 0) {
        Add-PnPListItem -List "AppAdmins" -Values @{
            Title = $AdminUpn
            Email = $AdminUpn.ToLower()
        } | Out-Null
        Write-Done "เพิ่ม $AdminUpn ในลิสต์ AppAdmins"
    }
    else { Write-Skip "AppAdmins มีข้อมูลอยู่แล้ว" }

    # --- 8.4 SkillsInterests ---
    $existingSkills = Get-PnPListItem -List "SkillsInterests" -PageSize 100
    if ($existingSkills.Count -eq 0) {
        foreach ($s in @(
            @{ Name = "Power Apps";        Kind = "Skill";    Level = "สูง" },
            @{ Name = "SharePoint Online"; Kind = "Skill";    Level = "ปานกลาง" },
            @{ Name = "การออกแบบหลักสูตร"; Kind = "Skill";    Level = "สูง" },
            @{ Name = "การวิเคราะห์ข้อมูล"; Kind = "Interest"; Level = "" }
        )) {
            Add-PnPListItem -List "SkillsInterests" -Values @{
                Title       = $s.Name
                PersonEmail = $AdminUpn.ToLower()
                Department  = "Learning and Development"
                ItemKind    = $s.Kind
                Level       = $s.Level
            } | Out-Null
        }
        Write-Done "ใส่ข้อมูลตัวอย่าง SkillsInterests"
    }
    else { Write-Skip "SkillsInterests มีข้อมูลอยู่แล้ว" }
}

#endregion

Write-Step "เสร็จสิ้น"
Write-Host @"
ขั้นตอนถัดไป:
  1. ตรวจว่า index ถูกสร้างครบ: List settings -> Indexed columns ของแต่ละลิสต์
  2. ตั้ง time zone ของ site เป็น (UTC+07:00) Bangkok ที่ Site settings -> Regional settings
  3. ขอให้เจ้าของลิสต์ BAU เพิ่ม index ให้ StartDate / EndDate / Status (ดูข้อความในไฟล์ 01)
  4. เปิด Power Apps Studio แล้วทำตามไฟล์ 03-beginner-guide-th.md
"@ -ForegroundColor Cyan

Disconnect-PnPOnline
