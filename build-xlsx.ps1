# build-xlsx.ps1
# One-off script that generates GPS-Ecosystem.xlsx from the seed data below.
# Seed data is mirrored from the original prototype's ecosystem-data.jsx so day-1
# editors start from a fully-populated workbook, not a blank one.
#
# Usage:
#   PowerShell -ExecutionPolicy Bypass -File .\build-xlsx.ps1
#
# Output:
#   GPS-Ecosystem.xlsx (next to this script)

$ErrorActionPreference = 'Stop'

# --- Seed data (mirrors GPS object in ecosystem-data.jsx) -----------------------
# Type: App | Data | Service | (blank) — what kind of component it is
# Flow: In | Out | Bidirectional | (blank) — direction of data flow vs GPS Stadium

$items = @(
    # Audiences -- people, not components. Type/Flow blank.
    @{ Section='Audiences'; Name='Educators & Teachers'; Short='Teachers'; Description='Classroom teachers and school support staff using GPS tools for instruction planning, MTSS, and student support. Primary consumers of Teacher Navigator.'; Future=$false; Type=''; Flow=''; Dashboards=''; DocURL=''; Order=1 }
    @{ Section='Audiences'; Name='School & District Leaders'; Short='Admins'; Description='Principals, district administrators, and LEA staff monitoring performance, guiding improvement, and managing compliance.'; Future=$false; Type=''; Flow=''; Dashboards=''; DocURL=''; Order=2 }
    @{ Section='Audiences'; Name='SCDE Staff'; Short='SCDE'; Description='State agency program offices and leadership, accountability, funding, research, and policy teams.'; Future=$false; Type=''; Flow=''; Dashboards=''; DocURL=''; Order=3 }
    @{ Section='Audiences'; Name='Public'; Short='Public'; Description='Policymakers, families, students, community members, and other stakeholders accessing public-facing data, report cards, and dashboards.'; Future=$false; Type=''; Flow=''; Dashboards=''; DocURL=''; Order=4 }

    # Tools & Apps -- all are apps users open; flow varies
    @{ Section='Tools & Apps'; Name='Teacher Navigator'; Short='Teacher Navigator'; Description='Student-level decision-support app for classroom teachers. Student, Classroom, Grade Level, and Section disaggregations of student performance dashboards. Expected growth projections for students; grouping feature for interventions.'; Future=$false; Type='App'; Flow='Out'; Dashboards='Student View; Class View; Grade View; Group View'; DocURL=''; Order=1 }
    @{ Section='Tools & Apps'; Name='Admin Navigator'; Short='Admin Navigator'; Description='School and district leader app with aggregate data and drill-down dashboards. Allows for analysis of enrollment, attendance, behavior, assessments, grades, and accountability indicators.'; Future=$false; Type='App'; Flow='Out'; Dashboards='Profile; Chronic Absenteeism; Enrollment; Behavior; Assessment; Course Grades; On-Track & CCR; Predictions'; DocURL='https://scdoe.sharepoint.com/:w:/s/X-TeamGPSGrowingPathwaysforStudents_Group/IQBu-joU23DHT4uYeTMIV_9_AbUuGDtnaZ08lrJPb-vsaM0?e=OoI1nh'; Order=2 }
    @{ Section='Tools & Apps'; Name='PDC Dashboards'; Short='Podium+'; Description='District-facing analytics dashboards delivered by the Palmetto Data Collaborative (PDC). Separate product from the Navigators - uses the same data foundation but accessed by LEA staff for ad-hoc and comparative visualizations.'; Future=$false; Type='App'; Flow='Out'; Dashboards=''; DocURL=''; Order=3 }
    @{ Section='Tools & Apps'; Name='Member Center'; Short='SCDE Navigator'; Description='App portal for SCDE and LEA users. Content spans HR functions and Grant Reporting to Data Reports and Support Requests.'; Future=$true; Type='App'; Flow='Bidirectional'; Dashboards=''; DocURL=''; Order=4 }

    # Outputs -- downstream uses of GPS data; mostly Out
    @{ Section='Outputs'; Name='School Report Cards'; Short=''; Description='School accountability reports published annually.'; Future=$false; Type='Data'; Flow='Out'; Dashboards=''; DocURL=''; Order=1 }
    @{ Section='Outputs'; Name='Funding Data'; Short=''; Description='Financial and funding visualizations and reports, driven by GPS analytics.'; Future=$false; Type='Data'; Flow='Out'; Dashboards=''; DocURL=''; Order=2 }
    @{ Section='Outputs'; Name='Research Requests'; Short=''; Description='Custom data extracts for approved external research partners and policy studies.'; Future=$false; Type='Data'; Flow='Out'; Dashboards=''; DocURL=''; Order=3 }
    @{ Section='Outputs'; Name='Program Office Reports'; Short=''; Description='Aggregated analytics (reports and visualizations) consumed by SCDE program offices for program management, evaluations, and federal reporting.'; Future=$false; Type='Data'; Flow='Out'; Dashboards=''; DocURL=''; Order=4 }
    @{ Section='Outputs'; Name='App Platform'; Short=''; Description='A planned app-store model: vetted SCDE offices or partners publish data services on the GPS platform.'; Future=$true; Type='Service'; Flow='Bidirectional'; Dashboards=''; DocURL=''; Order=5 }
    @{ Section='Outputs'; Name='State Navigator'; Short=''; Description='Planned state-leadership reports and dashboards: aggregating GPS analytics for cross-LEA comparison, accountability, and strategic planning.'; Future=$true; Type='App'; Flow='Out'; Dashboards=''; DocURL=''; Order=6 }

    # Data Services -- mix of apps and headless services; flow varies
    @{ Section='Data Services'; Name='Classlink'; Short=''; Description='Single sign-on and rostering provider. Pushes identity/roster data into connected apps.'; Future=$false; Type='Service'; Flow='In'; Dashboards=''; DocURL=''; Order=1 }
    @{ Section='Data Services'; Name='Palmetto Pathways'; Short='Palm. Pathways'; Description='Student graduation pathways and credential tracking application. Includes workflows for school counselors and school/district administrators.'; Future=$false; Type='App'; Flow='Bidirectional'; Dashboards=''; DocURL=''; Order=2 }
    @{ Section='Data Services'; Name='GiftED'; Short=''; Description='Gifted and talented program-management service. Sends program eligibility and participation data into GPS.'; Future=$false; Type='Service'; Flow='In'; Dashboards=''; DocURL=''; Order=3 }
    @{ Section='Data Services'; Name='Everyday Labs'; Short=''; Description='Attendance intervention partner. Consumes attendance data from GPS and delivers attendance recovery interventions to schools.'; Future=$false; Type='Service'; Flow='Bidirectional'; Dashboards=''; DocURL=''; Order=4 }
    @{ Section='Data Services'; Name='EdPlan SC'; Short=''; Description='Individualized Education Plan (IEP) management service. Bi-directional data exchange with GPS.'; Future=$false; Type='App'; Flow='Bidirectional'; Dashboards=''; DocURL=''; Order=5 }
    @{ Section='Data Services'; Name='Assessment Rostering'; Short=''; Description='Automated rostering service for state vendor assessments.'; Future=$false; Type='Service'; Flow='Bidirectional'; Dashboards=''; DocURL=''; Order=6 }

    # Inputs -- all data sources flowing into GPS
    @{ Section='Inputs'; Name='PowerSchool SIS'; Short=''; Description='Statewide Student Information System - the canonical source for enrollment, demographics, grades, attendance, and schedules. PowerSchool is an external vendor system integrated with the Ed-Fi API.'; Future=$false; Type='Data'; Flow='In'; Dashboards=''; DocURL=''; Order=1 }
    @{ Section='Inputs'; Name='Students Assessments'; Short=''; Description='Student assessment results (SC READY, SC PASS, end-of-course, etc.) loaded into GPS for analytics.'; Future=$false; Type='Data'; Flow='In'; Dashboards=''; DocURL=''; Order=2 }
    @{ Section='Inputs'; Name='EdPlanSC'; Short=''; Description='Individualized Education Plan (IEP) management service. Note: EdPlanSC also appears as a Data Service because it both produces and consumes data.'; Future=$false; Type='Data'; Flow='In'; Dashboards=''; DocURL=''; Order=3 }
    @{ Section='Inputs'; Name='iReady'; Short=''; Description='Diagnostic and instructional assessment data from district adoptions of iReady.'; Future=$false; Type='Data'; Flow='In'; Dashboards=''; DocURL=''; Order=4 }
    @{ Section='Inputs'; Name='Early Childhood Programs'; Short=''; Description='Early childhood program enrollment and outcomes data.'; Future=$false; Type='Data'; Flow='In'; Dashboards=''; DocURL=''; Order=5 }
    @{ Section='Inputs'; Name='ESTF Data'; Short=''; Description='Educational Scholarship Trust Fund data. A SCDE program evaluated via GPS.'; Future=$false; Type='Data'; Flow='In'; Dashboards=''; DocURL=''; Order=6 }
    @{ Section='Inputs'; Name='Educator Workforce'; Short=''; Description='Educator pipeline, workforce, and credentialing data - inclusive of SCEducator, SCLead, and Educator Preparation program data.'; Future=$true; Type='Data'; Flow='In'; Dashboards=''; DocURL=''; Order=7 }
    @{ Section='Inputs'; Name='Financial Reporting'; Short=''; Description='Transparent LEA and state-level financial data - connecting spending to student outcomes.'; Future=$true; Type='Data'; Flow='In'; Dashboards=''; DocURL=''; Order=8 }
    @{ Section='Inputs'; Name='CTE Credentials'; Short=''; Description='Career and technical pathway data - completions, credentials, and workforce outcomes.'; Future=$true; Type='Data'; Flow='In'; Dashboards=''; DocURL=''; Order=9 }

    # Foundation -- these ARE GPS; Type/Flow don't apply (left blank)
    @{ Section='Foundation'; Name='Stadium Data Warehouse'; Short='Stadium'; Description='The centralized cloud longitudinal data warehouse where all GPS data lands and is analytically modeled.'; Future=$false; Type=''; Flow=''; Dashboards=''; DocURL=''; Order=1 }
    @{ Section='Foundation'; Name='Ed-Fi Integration Layer'; Short='Ed-Fi'; Description='The Ed-Fi Data Standard and API - the interoperability model that defines how educational data flows between systems. Sits between systems/applications and Stadium.'; Future=$false; Type=''; Flow=''; Dashboards=''; DocURL=''; Order=2 }
    @{ Section='Foundation'; Name='Data Standards & APIs'; Short='APIs'; Description='The shared APIs, data contracts, definitions, and naming conventions that let services exchange data consistently. The "rules of the road" that sit on top of Ed-Fi.'; Future=$false; Type=''; Flow=''; Dashboards=''; DocURL=''; Order=3 }
    @{ Section='Foundation'; Name='Security & Governance'; Short='Security'; Description='Cross-cutting access controls, privacy policies, audit, and data governance practices. Not a separate system, as it wraps every other foundation component.'; Future=$false; Type=''; Flow=''; Dashboards=''; DocURL=''; Order=4 }

    # Governance -- meta-level roles, not components. Type/Flow blank.
    @{ Section='Governance'; Name='GPS Leadership'; Short=''; Description='GPS Leadership keeps time - accountable for ecosystem coherence across SCDE divisions and partners.'; Future=$false; Type=''; Flow=''; Dashboards=''; DocURL=''; Order=1 }
    @{ Section='Governance'; Name='X-Team'; Short=''; Description='X-Team sets the score - shared standards, interoperability, governance across the GPS portfolio.'; Future=$false; Type=''; Flow=''; Dashboards=''; DocURL=''; Order=2 }
    @{ Section='Governance'; Name='SCDE Offices'; Short=''; Description='Divisions own their instruments: applications, data, budgets, vendors, roadmaps.'; Future=$false; Type=''; Flow=''; Dashboards=''; DocURL=''; Order=3 }
)

$settings = [ordered]@{
    'Title'                = 'GPS Ecosystem'
    'Definition'           = 'A coordinated, statewide system of data, tools, and applications that deliver decision-support to educators, leaders, and policymakers across South Carolina K-12 to guide student success.'
    'AsterisksNote'        = '* In current/future development'
    'AudiencesSubtitle'    = 'Audiences GPS exists to serve - consumers of the ecosystem.'
    'ToolsSubtitle'        = 'What people open. User-facing apps and downstream products that consume GPS data.'
    'DataServicesSubtitle' = 'Services that produce and consume data via the foundation, plus the integrated applications.'
    'FoundationSubtitle'   = 'The shared technical substrate. Stadium is the warehouse; Ed-Fi defines interoperability; APIs and Security wrap everything.'
    'GovernanceSubtitle'   = 'Who keeps the ecosystem coherent across SCDE divisions and partners.'
}

$sectionOptions = 'Audiences,Tools & Apps,Data Services,Foundation,Governance,Outputs,Inputs'
$typeOptions    = 'App,Data,Service'
$flowOptions    = 'In,Out,Bidirectional'

# --- Excel COM workbook generation ---------------------------------------------

$outPath = Join-Path $PSScriptRoot 'GPS-Ecosystem.xlsx'
if (Test-Path $outPath) { Remove-Item $outPath -Force }

Write-Host "Building $outPath ..."

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Add()
    $items_ws = $wb.Worksheets.Item(1)
    $items_ws.Name = 'Items'
    $settings_ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $items_ws)
    $settings_ws.Name = 'Settings'

    # --- Items sheet ---
    $headers = @('Section','Name','Short','Description','Future','Type','Flow','Dashboards','DocURL','Order')
    for ($c = 0; $c -lt $headers.Count; $c++) {
        $items_ws.Cells.Item(1, $c + 1).Value2 = $headers[$c]
    }
    $hdr = $items_ws.Range($items_ws.Cells.Item(1,1), $items_ws.Cells.Item(1, $headers.Count))
    $hdr.Font.Bold = $true
    $hdr.Interior.Color = 3092271
    $hdr.Font.Color = 16777215
    $hdr.RowHeight = 22
    $hdr.HorizontalAlignment = -4131

    $row = 2
    foreach ($it in $items) {
        $items_ws.Cells.Item($row,  1).Value2 = $it.Section
        $items_ws.Cells.Item($row,  2).Value2 = $it.Name
        $items_ws.Cells.Item($row,  3).Value2 = $it.Short
        $items_ws.Cells.Item($row,  4).Value2 = $it.Description
        $items_ws.Cells.Item($row,  5).Value2 = if ($it.Future) { 'TRUE' } else { 'FALSE' }
        $items_ws.Cells.Item($row,  6).Value2 = $it.Type
        $items_ws.Cells.Item($row,  7).Value2 = $it.Flow
        $items_ws.Cells.Item($row,  8).Value2 = $it.Dashboards
        $items_ws.Cells.Item($row,  9).Value2 = $it.DocURL
        $items_ws.Cells.Item($row, 10).Value2 = [string]$it.Order
        $row++
    }

    # Column widths (10 columns now)
    $items_ws.Columns.Item(1).ColumnWidth  = 16   # Section
    $items_ws.Columns.Item(2).ColumnWidth  = 28   # Name
    $items_ws.Columns.Item(3).ColumnWidth  = 18   # Short
    $items_ws.Columns.Item(4).ColumnWidth  = 70   # Description
    $items_ws.Columns.Item(5).ColumnWidth  = 9    # Future
    $items_ws.Columns.Item(6).ColumnWidth  = 9    # Type
    $items_ws.Columns.Item(7).ColumnWidth  = 14   # Flow
    $items_ws.Columns.Item(8).ColumnWidth  = 40   # Dashboards
    $items_ws.Columns.Item(9).ColumnWidth  = 50   # DocURL
    $items_ws.Columns.Item(10).ColumnWidth = 8    # Order

    $items_ws.Columns.Item(4).WrapText = $true
    $items_ws.Columns.Item(8).WrapText = $true

    # Data validation: Section (column A)
    $sectionRange = $items_ws.Range("A2:A500")
    $sectionRange.Validation.Delete()
    [void]$sectionRange.Validation.Add(3, 1, 1, $sectionOptions)
    $sectionRange.Validation.IgnoreBlank = $true
    $sectionRange.Validation.InCellDropdown = $true
    $sectionRange.Validation.ErrorTitle = 'Invalid section'
    $sectionRange.Validation.ErrorMessage = "Pick one of: $sectionOptions"

    # Data validation: Future (column E)
    $futureRange = $items_ws.Range("E2:E500")
    $futureRange.Validation.Delete()
    [void]$futureRange.Validation.Add(3, 1, 1, 'TRUE,FALSE')
    $futureRange.Validation.IgnoreBlank = $true
    $futureRange.Validation.InCellDropdown = $true

    # Data validation: Type (column F)
    $typeRange = $items_ws.Range("F2:F500")
    $typeRange.Validation.Delete()
    [void]$typeRange.Validation.Add(3, 1, 1, $typeOptions)
    $typeRange.Validation.IgnoreBlank = $true
    $typeRange.Validation.InCellDropdown = $true
    $typeRange.Validation.ErrorTitle = 'Invalid type'
    $typeRange.Validation.ErrorMessage = "Pick one of: $typeOptions (or leave blank)"

    # Data validation: Flow (column G)
    $flowRange = $items_ws.Range("G2:G500")
    $flowRange.Validation.Delete()
    [void]$flowRange.Validation.Add(3, 1, 1, $flowOptions)
    $flowRange.Validation.IgnoreBlank = $true
    $flowRange.Validation.InCellDropdown = $true
    $flowRange.Validation.ErrorTitle = 'Invalid flow'
    $flowRange.Validation.ErrorMessage = "Pick one of: $flowOptions (or leave blank)"

    # Freeze top row
    $items_ws.Activate()
    $items_ws.Application.ActiveWindow.SplitRow = 1
    $items_ws.Application.ActiveWindow.FreezePanes = $true

    # --- Settings sheet ---
    $settings_ws.Cells.Item(1, 1).Value2 = 'Key'
    $settings_ws.Cells.Item(1, 2).Value2 = 'Value'
    $shdr = $settings_ws.Range($settings_ws.Cells.Item(1,1), $settings_ws.Cells.Item(1, 2))
    $shdr.Font.Bold = $true
    $shdr.Interior.Color = 3092271
    $shdr.Font.Color = 16777215
    $shdr.RowHeight = 22

    $srow = 2
    foreach ($k in $settings.Keys) {
        $settings_ws.Cells.Item($srow, 1).Value2 = $k
        $settings_ws.Cells.Item($srow, 2).Value2 = $settings[$k]
        $srow++
    }
    $settings_ws.Columns.Item(1).ColumnWidth = 24
    $settings_ws.Columns.Item(2).ColumnWidth = 100
    $settings_ws.Columns.Item(2).WrapText = $true

    # xlOpenXMLWorkbook = 51
    $wb.SaveAs($outPath, 51)
    $wb.Close($false)
    Write-Host "Wrote $outPath"
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect() | Out-Null
    [System.GC]::WaitForPendingFinalizers() | Out-Null
}
