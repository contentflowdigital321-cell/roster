Attribute VB_Name = "RosterModule"
Option Explicit

' ================================================================
' SENIOR CARE CENTER - ROSTER PLANNING TOOL (Excel / VBA)
' ================================================================
' Sheets required (auto-created if missing): "Staff Settings", "Roster", "Dashboard"
' Run via the buttons on the "Home" tab, or directly: GenerateOptimalRoster / RefreshDashboard
' ================================================================

' ----------------------------- CONFIG -----------------------------

Private Const SHEET_STAFF As String = "Staff Settings"
Private Const SHEET_ROSTER As String = "Roster"
Private Const SHEET_DASHBOARD As String = "Dashboard"

Private Const SHIFT_START_MIN As Integer = 7 * 60   ' 07:00
Private Const SHIFT_END_MIN As Integer = 16 * 60    ' 16:00
Private Const BLOCK_MIN As Integer = 30
Private Const ROSTER_FIRST_TIME_COL As Long = 3      ' Column C

Private Const LUNCH_DURATION_MIN As Integer = 60
' Facility-wide rule: no two employees, of any role, may ever start lunch in the same slot.
' Each of the 4 slots is used by at most one person, so at most 4 employees total can be
' scheduled for lunch across the whole center.
Private Const MAX_TOTAL_LUNCH_CAPACITY As Long = 4
Private Const ERR_ROSTER As Long = vbObjectError + 513 ' custom "constraint impossible" error

Private Type EmployeeRec
    EName As String
    ERole As String
    LunchStart As Integer   ' -1 = not assigned yet
    LunchSlotLabel As String
End Type

' ----------------------------- LOOKUP TABLES -----------------------------

' Named RoleNames (not "Roles") because VBA is case-insensitive: a same-named local
' variable "roles" would otherwise shadow this function and "Roles()" would be parsed
' as indexing into that (empty) local variable instead of calling this function -
' raising "Subscript out of range" at every call site. Do not rename back to "Roles".
Private Function RoleNames() As Variant
    RoleNames = Array("Nurse", "Care Staff", "Kitchen", "Escort")
End Function

' The four permitted lunch start times, per spec. Each lunch is exactly 60 minutes
' (two consecutive 30-minute blocks). Latest end is 14:00, safely inside the
' "must end by 2:30 PM" window.
Private Function LunchSlotIds() As Variant
    LunchSlotIds = Array("11:30", "12:00", "12:30", "13:00")
End Function

Private Function LunchSlotStarts() As Variant
    LunchSlotStarts = Array(11 * 60 + 30, 12 * 60, 12 * 60 + 30, 13 * 60)
End Function

Private Function LunchBlockLabels() As Variant
    LunchBlockLabels = Array("11:30", "12:00", "12:30", "13:00", "13:30")
End Function

' ----------------------------- COLORS -----------------------------

Private Function ColorCharcoal() As Long: ColorCharcoal = RGB(38, 50, 56): End Function
Private Function ColorWorkBg() As Long: ColorWorkBg = RGB(220, 240, 220): End Function
Private Function ColorWorkText() As Long: ColorWorkText = RGB(37, 96, 41): End Function
Private Function ColorLunchBg() As Long: ColorLunchBg = RGB(255, 231, 204): End Function
Private Function ColorLunchText() As Long: ColorLunchText = RGB(183, 94, 0): End Function
Private Function ColorOk() As Long: ColorOk = RGB(220, 240, 220): End Function
Private Function ColorWarning() As Long: ColorWarning = RGB(255, 243, 205): End Function
Private Function ColorErr() As Long: ColorErr = RGB(250, 219, 216): End Function

' ----------------------------- MAIN ENTRY POINTS -----------------------------

Public Sub GenerateOptimalRoster()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim wb As Workbook
    Set wb = ThisWorkbook
    EnsureTemplateSheets wb

    Dim staffSheet As Worksheet
    Set staffSheet = wb.Sheets(SHEET_STAFF)

    Dim employees() As EmployeeRec
    Dim empCount As Long
    Dim readWarnings() As String
    Dim readWarnCount As Long
    ReadActiveStaff staffSheet, employees, empCount, readWarnings, readWarnCount

    If empCount = 0 Then
        MsgBox "No active employees were found in ""Staff Settings"". Add staff and mark ""Is Active"" TRUE before generating a roster.", vbExclamation, "No Active Employees"
        GoTo Finish
    End If

    Dim assignWarnings() As String
    Dim assignWarnCount As Long
    AssignAllLunches employees, empCount, assignWarnings, assignWarnCount

    Dim rosterSheet As Worksheet
    Set rosterSheet = wb.Sheets(SHEET_ROSTER)
    WriteRosterSheet rosterSheet, employees, empCount
    FormatRosterSheet rosterSheet, empCount

    Dim allNotes() As String
    Dim allCount As Long
    MergeStringArrays readWarnings, readWarnCount, assignWarnings, assignWarnCount, allNotes, allCount
    RefreshDashboardInternal wb, allNotes, allCount

    MsgBox "Roster successfully generated for " & empCount & " active employee(s)." & vbCrLf & _
           "See the Dashboard sheet for coverage metrics and validation warnings.", vbInformation, "Roster Generated"

Finish:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    If Err.Number = ERR_ROSTER Then
        MsgBox Err.Description, vbCritical, "Cannot Generate Roster - Constraint Violation"
    Else
        MsgBox "Something went wrong while generating the roster:" & vbCrLf & Err.Description, vbCritical, "Unexpected Error"
    End If
    Resume Finish
End Sub

' Re-validates and redraws the Dashboard from whatever is currently in the Roster sheet,
' without touching Staff Settings or Roster. Useful after manual edits to the roster grid.
Public Sub RefreshDashboard()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    Dim wb As Workbook
    Set wb = ThisWorkbook
    EnsureTemplateSheets wb

    Dim emptyNotes(1 To 1) As String
    RefreshDashboardInternal wb, emptyNotes, 0

    Application.ScreenUpdating = True
    MsgBox "Dashboard metrics and validation warnings have been updated.", vbInformation, "Dashboard Refreshed"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Could not refresh the dashboard:" & vbCrLf & Err.Description, vbCritical, "Unexpected Error"
End Sub

Public Sub InitializeSheetTemplates()
    On Error GoTo ErrHandler
    EnsureTemplateSheets ThisWorkbook
    MsgBox "Sheet templates are ready. Fill in ""Staff Settings"", then click ""Generate Optimal Roster"".", vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Could not initialize sheet templates:" & vbCrLf & Err.Description, vbCritical
End Sub

' ----------------------------- SHEET SETUP -----------------------------

Private Sub EnsureTemplateSheets(wb As Workbook)
    Dim staffSheet As Worksheet
    On Error Resume Next
    Set staffSheet = wb.Sheets(SHEET_STAFF)
    On Error GoTo 0
    If staffSheet Is Nothing Then
        Set staffSheet = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        staffSheet.Name = SHEET_STAFF
    End If
    If staffSheet.Cells(1, 1).Value = "" Then
        staffSheet.Cells(1, 1).Value = "Employee Name"
        staffSheet.Cells(1, 2).Value = "Role"
        staffSheet.Cells(1, 3).Value = "Is Active"
        staffSheet.Cells(2, 1).Value = "Jane Smith": staffSheet.Cells(2, 2).Value = "Nurse": staffSheet.Cells(2, 3).Value = True
        staffSheet.Cells(3, 1).Value = "Alex Chen": staffSheet.Cells(3, 2).Value = "Care Staff": staffSheet.Cells(3, 3).Value = True
        staffSheet.Cells(4, 1).Value = "Maria Lopez": staffSheet.Cells(4, 2).Value = "Kitchen": staffSheet.Cells(4, 3).Value = True
    End If
    FormatStaffSettingsSheet staffSheet

    Dim rosterSheet As Worksheet
    On Error Resume Next
    Set rosterSheet = wb.Sheets(SHEET_ROSTER)
    On Error GoTo 0
    If rosterSheet Is Nothing Then
        Set rosterSheet = wb.Sheets.Add(After:=staffSheet)
        rosterSheet.Name = SHEET_ROSTER
    End If
    If rosterSheet.Cells(1, 1).Value = "" Then
        WriteRosterHeaderOnly rosterSheet
    End If

    Dim dashSheet As Worksheet
    On Error Resume Next
    Set dashSheet = wb.Sheets(SHEET_DASHBOARD)
    On Error GoTo 0
    If dashSheet Is Nothing Then
        Set dashSheet = wb.Sheets.Add(After:=rosterSheet)
        dashSheet.Name = SHEET_DASHBOARD
    End If
End Sub

Private Sub FormatStaffSettingsSheet(staffSheet As Worksheet)
    With staffSheet.Range("A1:C1")
        .Interior.Color = ColorCharcoal()
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    staffSheet.Columns(1).ColumnWidth = 26
    staffSheet.Columns(2).ColumnWidth = 16
    staffSheet.Columns(3).ColumnWidth = 12

    Dim roleList As String
    roleList = Join(RoleNames(), ",")

    With staffSheet.Range("B2:B300").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=roleList
    End With
    With staffSheet.Range("C2:C300").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="TRUE,FALSE"
    End With
End Sub

Private Sub WriteRosterHeaderOnly(rosterSheet As Worksheet)
    rosterSheet.Cells.Clear
    Dim headers() As String
    headers = BuildRosterHeaders()
    Dim totalCols As Long
    totalCols = UBound(headers) - LBound(headers) + 1

    ' Force text format BEFORE writing "07:00"-style strings, otherwise Excel silently
    ' converts them to time serial numbers instead of keeping them as literal text -
    ' which would break every later comparison against these header labels.
    rosterSheet.Range(rosterSheet.Cells(1, ROSTER_FIRST_TIME_COL), rosterSheet.Cells(1, totalCols)).NumberFormat = "@"

    Dim c As Long
    For c = LBound(headers) To UBound(headers)
        rosterSheet.Cells(1, c + 1).Value = headers(c)
    Next c
End Sub

Private Function BuildRosterHeaders() As String()
    Dim nBlocks As Long
    nBlocks = (SHIFT_END_MIN - SHIFT_START_MIN) \ BLOCK_MIN
    Dim headers() As String
    ReDim headers(0 To nBlocks + 1)
    headers(0) = "Employee Name"
    headers(1) = "Role"
    Dim t As Integer, i As Long
    i = 2
    For t = SHIFT_START_MIN To SHIFT_END_MIN - 1 Step BLOCK_MIN
        headers(i) = MinutesToLabel(t)
        i = i + 1
    Next t
    BuildRosterHeaders = headers
End Function

Private Function MinutesToLabel(totalMinutes As Integer) As String
    Dim h As Integer, m As Integer
    h = totalMinutes \ 60
    m = totalMinutes Mod 60
    MinutesToLabel = Format(h, "00") & ":" & Format(m, "00")
End Function

' ----------------------------- READ STAFF -----------------------------

Private Sub ReadActiveStaff(staffSheet As Worksheet, ByRef employees() As EmployeeRec, ByRef empCount As Long, _
                             ByRef warnings() As String, ByRef warnCount As Long)
    Dim lastRow As Long
    lastRow = staffSheet.Cells(staffSheet.Rows.Count, 1).End(xlUp).Row

    empCount = 0
    warnCount = 0
    ReDim employees(1 To 1)
    ReDim warnings(1 To 1)
    If lastRow < 2 Then Exit Sub

    ReDim employees(1 To lastRow)
    ReDim warnings(1 To lastRow)

    Dim roles As Variant
    roles = RoleNames()

    Dim r As Long
    For r = 2 To lastRow
        Dim nm As String
        nm = Trim(CStr(staffSheet.Cells(r, 1).Value))
        If nm = "" Then GoTo NextRow

        Dim actVal As Variant
        actVal = staffSheet.Cells(r, 3).Value
        Dim isActive As Boolean
        If VarType(actVal) = vbBoolean Then
            isActive = CBool(actVal)
        Else
            Dim s As String
            s = UCase(Trim(CStr(actVal)))
            isActive = (s = "TRUE" Or s = "YES" Or s = "ACTIVE")
        End If
        If Not isActive Then GoTo NextRow

        Dim rl As String
        rl = Trim(CStr(staffSheet.Cells(r, 2).Value))
        Dim matchedRole As String
        matchedRole = ""
        Dim i As Long
        For i = LBound(roles) To UBound(roles)
            If LCase(roles(i)) = LCase(rl) Then
                matchedRole = roles(i)
                Exit For
            End If
        Next i

        If matchedRole = "" Then
            warnCount = warnCount + 1
            warnings(warnCount) = "[WARNING] Row " & r & " (""" & nm & """) has an unrecognized role """ & rl & _
                """ and was excluded from the roster. Valid roles: " & Join(roles, ", ") & "."
            GoTo NextRow
        End If

        empCount = empCount + 1
        employees(empCount).EName = nm
        employees(empCount).ERole = matchedRole
        employees(empCount).LunchStart = -1
        employees(empCount).LunchSlotLabel = ""
NextRow:
    Next r

    If empCount = 0 Then
        ReDim employees(1 To 1)
    Else
        ReDim Preserve employees(1 To empCount)
    End If
    If warnCount = 0 Then
        ReDim warnings(1 To 1)
    Else
        ReDim Preserve warnings(1 To warnCount)
    End If
End Sub

' ----------------------------- LUNCH ASSIGNMENT -----------------------------

' Assigns every active employee (across all roles) to one of the 4 lunch slots so that no
' two employees ever start lunch in the same slot. Since only 4 slots exist, at most 4
' employees total can be scheduled - this is a hard facility-wide cap, not a per-role one.
' Sorts `employees` in place by role priority then name, then mutates each record with
' LunchStart / LunchSlotLabel. Raises ERR_ROSTER if empCount > MAX_TOTAL_LUNCH_CAPACITY.
Private Sub AssignAllLunches(ByRef employees() As EmployeeRec, empCount As Long, ByRef warnings() As String, ByRef warnCount As Long)
    warnCount = 0
    ReDim warnings(1 To 1)
    If empCount = 0 Then Exit Sub

    If empCount > MAX_TOTAL_LUNCH_CAPACITY Then
        Err.Raise ERR_ROSTER, "AssignAllLunches", _
            "There are " & empCount & " active employees, but only " & MAX_TOTAL_LUNCH_CAPACITY & _
            " lunch slots exist and each slot may be used by only one person - no two staff may start lunch " & _
            "at the same time. Reduce the active headcount to " & MAX_TOTAL_LUNCH_CAPACITY & " or fewer, mark " & _
            "extra staff inactive, or add more lunch slots, then try again."
    End If

    ' Sort by role priority then name so the solver's placement is deterministic and roles
    ' that matter most for coverage (Nurse first, per spec) are considered first when ties occur.
    SortEmployeesByRoleThenName employees, empCount

    Dim slotIds As Variant: slotIds = LunchSlotIds()
    Dim slotStarts As Variant: slotStarts = LunchSlotStarts()
    Dim numSlots As Long: numSlots = UBound(slotIds) - LBound(slotIds) + 1

    Dim used() As Boolean
    ReDim used(0 To numSlots - 1)
    Dim current() As Long
    ReDim current(1 To empCount)
    Dim best() As Long
    ReDim best(1 To empCount)
    Dim bestViolations As Long
    bestViolations = 2147483647

    SolveLunchSlotsRecurse employees, empCount, 1, 0, used, current, best, bestViolations, numSlots

    Dim i As Long
    For i = 1 To empCount
        employees(i).LunchStart = slotStarts(best(i))
        employees(i).LunchSlotLabel = slotIds(best(i))
    Next i
End Sub

' Finds the assignment of employees to distinct lunch slots that minimizes same-role
' adjacent-slot pairs. Adjacent slots (e.g. 12:00 and 12:30) overlap by 30 minutes, so if
' they're the only two people in that role, the role goes fully unstaffed for that half
' hour. The search space is tiny (at most 4 employees over 4 slots, <=24 permutations), so
' this small backtracking search finds the true optimum instead of relying on a fixed greedy
' fill order, which can miss a zero-violation arrangement that does exist (e.g. 2 Nurses +
' 2 Care Staff can always be split into two non-overlapping pairs, but a naive sequential
' fill can accidentally leave the second role with an adjacent, overlapping pair).
Private Sub SolveLunchSlotsRecurse(employees() As EmployeeRec, empCount As Long, i As Long, violations As Long, _
                                    ByRef used() As Boolean, ByRef current() As Long, ByRef best() As Long, _
                                    ByRef bestViolations As Long, numSlots As Long)
    If violations >= bestViolations Then Exit Sub

    If i > empCount Then
        bestViolations = violations
        Dim k As Long
        For k = 1 To empCount
            best(k) = current(k)
        Next k
        Exit Sub
    End If

    Dim s As Long
    For s = 0 To numSlots - 1
        If Not used(s) Then
            Dim added As Long
            added = 0
            Dim j As Long
            For j = 1 To i - 1
                If employees(j).ERole = employees(i).ERole And Abs(current(j) - s) = 1 Then added = added + 1
            Next j
            used(s) = True
            current(i) = s
            SolveLunchSlotsRecurse employees, empCount, i + 1, violations + added, used, current, best, bestViolations, numSlots
            used(s) = False
            If bestViolations = 0 Then Exit Sub ' already optimal, stop searching
        End If
    Next s
End Sub

Private Sub SortEmployeesByRoleThenName(ByRef employees() As EmployeeRec, empCount As Long)
    Dim roles As Variant: roles = RoleNames()
    Dim a As Long, b As Long
    Dim tmp As EmployeeRec
    For a = 2 To empCount
        tmp = employees(a)
        b = a - 1
        Do While b >= 1 And CompareEmpRec(employees(b), tmp, roles) > 0
            employees(b + 1) = employees(b)
            b = b - 1
        Loop
        employees(b + 1) = tmp
    Next a
End Sub

Private Function CompareEmpRec(a As EmployeeRec, b As EmployeeRec, roles As Variant) As Long
    Dim ra As Long, rb As Long
    ra = RoleIndex(a.ERole, roles)
    rb = RoleIndex(b.ERole, roles)
    If ra <> rb Then
        CompareEmpRec = ra - rb
    Else
        CompareEmpRec = StrComp(a.EName, b.EName, vbTextCompare)
    End If
End Function

' ----------------------------- ROSTER SHEET -----------------------------

' Assumes `employees` is already sorted by role then name (AssignAllLunches sorts it in
' place before this is called).
Private Sub WriteRosterSheet(rosterSheet As Worksheet, employees() As EmployeeRec, empCount As Long)
    rosterSheet.Cells.Clear

    Dim headers() As String
    headers = BuildRosterHeaders()
    Dim totalCols As Long
    totalCols = UBound(headers) - LBound(headers) + 1

    ' See note in WriteRosterHeaderOnly: must set Text format before writing "HH:MM" strings.
    rosterSheet.Range(rosterSheet.Cells(1, ROSTER_FIRST_TIME_COL), rosterSheet.Cells(1, totalCols)).NumberFormat = "@"

    Dim c As Long
    For c = LBound(headers) To UBound(headers)
        rosterSheet.Cells(1, c + 1).Value = headers(c)
    Next c

    Dim nBlocks As Long
    nBlocks = (SHIFT_END_MIN - SHIFT_START_MIN) \ BLOCK_MIN
    Dim blockStarts() As Integer
    ReDim blockStarts(0 To nBlocks - 1)
    Dim t As Integer, bi As Long
    bi = 0
    For t = SHIFT_START_MIN To SHIFT_END_MIN - 1 Step BLOCK_MIN
        blockStarts(bi) = t
        bi = bi + 1
    Next t

    If empCount > 0 Then
        rosterSheet.Range(rosterSheet.Cells(2, ROSTER_FIRST_TIME_COL), rosterSheet.Cells(empCount + 1, totalCols)).NumberFormat = "@"
    End If

    Dim row As Long
    For row = 1 To empCount
        Dim emp As EmployeeRec
        emp = employees(row)
        rosterSheet.Cells(row + 1, 1).Value = emp.EName
        rosterSheet.Cells(row + 1, 2).Value = emp.ERole
        For bi = 0 To nBlocks - 1
            Dim onLunch As Boolean
            onLunch = (emp.LunchStart >= 0) And (blockStarts(bi) >= emp.LunchStart) And (blockStarts(bi) < emp.LunchStart + LUNCH_DURATION_MIN)
            rosterSheet.Cells(row + 1, ROSTER_FIRST_TIME_COL + bi).Value = IIf(onLunch, "LUNCH", "WORK")
        Next bi
    Next row
End Sub

Private Function RoleIndex(role As String, roles As Variant) As Long
    Dim i As Long
    For i = LBound(roles) To UBound(roles)
        If roles(i) = role Then
            RoleIndex = i
            Exit Function
        End If
    Next i
    RoleIndex = 999
End Function

Private Sub FormatRosterSheet(rosterSheet As Worksheet, empCount As Long)
    Dim headers() As String
    headers = BuildRosterHeaders()
    Dim totalCols As Long
    totalCols = UBound(headers) - LBound(headers) + 1

    With rosterSheet.Range(rosterSheet.Cells(1, 1), rosterSheet.Cells(1, totalCols))
        .Interior.Color = ColorCharcoal()
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    rosterSheet.Activate
    rosterSheet.Cells(2, ROSTER_FIRST_TIME_COL).Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True

    rosterSheet.Columns(1).ColumnWidth = 22
    rosterSheet.Columns(2).ColumnWidth = 14
    Dim c As Long
    For c = ROSTER_FIRST_TIME_COL To totalCols
        rosterSheet.Columns(c).ColumnWidth = 7
    Next c

    If empCount > 0 Then
        Dim dataRange As Range
        Set dataRange = rosterSheet.Range(rosterSheet.Cells(1, 1), rosterSheet.Cells(empCount + 1, totalCols))
        dataRange.Borders.LineStyle = xlContinuous
        dataRange.Borders.Color = RGB(208, 208, 208)

        rosterSheet.Range(rosterSheet.Cells(2, 1), rosterSheet.Cells(empCount + 1, 1)).Font.Bold = True
        rosterSheet.Range(rosterSheet.Cells(2, 1), rosterSheet.Cells(empCount + 1, totalCols)).HorizontalAlignment = xlCenter
        rosterSheet.Range(rosterSheet.Cells(2, 1), rosterSheet.Cells(empCount + 1, 1)).HorizontalAlignment = xlLeft

        Dim timeRange As Range
        Set timeRange = rosterSheet.Range(rosterSheet.Cells(2, ROSTER_FIRST_TIME_COL), rosterSheet.Cells(empCount + 1, totalCols))
        timeRange.FormatConditions.Delete

        Dim fc1 As FormatCondition
        Set fc1 = timeRange.FormatConditions.Add(Type:=xlTextString, String:="WORK", TextOperator:=xlContains)
        fc1.Interior.Color = ColorWorkBg()
        fc1.Font.Color = ColorWorkText()

        Dim fc2 As FormatCondition
        Set fc2 = timeRange.FormatConditions.Add(Type:=xlTextString, String:="LUNCH", TextOperator:=xlContains)
        fc2.Interior.Color = ColorLunchBg()
        fc2.Font.Color = ColorLunchText()
    End If

    rosterSheet.Rows(1).RowHeight = 22
End Sub

' ----------------------------- DASHBOARD -----------------------------

Private Sub RefreshDashboardInternal(wb As Workbook, notes() As String, noteCount As Long)
    Dim rosterSheet As Worksheet: Set rosterSheet = wb.Sheets(SHEET_ROSTER)
    Dim dashSheet As Worksheet: Set dashSheet = wb.Sheets(SHEET_DASHBOARD)

    Dim roleTotals() As Long
    Dim coverage() As Long
    Dim scanWarnings() As String
    Dim scanWarnCount As Long
    ScanRosterForValidation rosterSheet, roleTotals, coverage, scanWarnings, scanWarnCount

    Dim allWarnings() As String
    Dim allCount As Long
    MergeStringArrays notes, noteCount, scanWarnings, scanWarnCount, allWarnings, allCount

    WriteDashboard dashSheet, roleTotals, coverage, allWarnings, allCount
End Sub

Private Sub MergeStringArrays(a() As String, aCount As Long, b() As String, bCount As Long, ByRef result() As String, ByRef resultCount As Long)
    resultCount = aCount + bCount
    If resultCount = 0 Then
        ReDim result(1 To 1)
        Exit Sub
    End If
    ReDim result(1 To resultCount)
    Dim i As Long
    For i = 1 To aCount
        result(i) = a(i)
    Next i
    For i = 1 To bCount
        result(aCount + i) = b(i)
    Next i
End Sub

' Reads the Roster sheet's WORK/LUNCH grid directly (single source of truth) and
' recomputes per-role, per-block lunch coverage plus rule violations. Running this
' after a manual edit to the Roster grid re-validates it exactly the same way.
Private Sub ScanRosterForValidation(rosterSheet As Worksheet, ByRef roleTotals() As Long, ByRef coverage() As Long, _
                                     ByRef warnings() As String, ByRef warnCount As Long)
    Dim roles As Variant: roles = RoleNames()
    Dim blockLabels As Variant: blockLabels = LunchBlockLabels()
    Dim nRoles As Long: nRoles = UBound(roles) - LBound(roles) + 1
    Dim nBlocks As Long: nBlocks = UBound(blockLabels) - LBound(blockLabels) + 1

    ReDim roleTotals(0 To nRoles - 1)
    ReDim coverage(0 To nRoles - 1, 0 To nBlocks - 1)
    warnCount = 0
    ReDim warnings(1 To 1)

    Dim lastRow As Long: lastRow = rosterSheet.Cells(rosterSheet.Rows.Count, 1).End(xlUp).Row
    Dim lastCol As Long: lastCol = rosterSheet.Cells(1, rosterSheet.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Or lastCol < ROSTER_FIRST_TIME_COL Then
        warnCount = 1
        ReDim warnings(1 To 1)
        warnings(1) = "[INFO] Roster has not been generated yet. Click ""Generate Optimal Roster"" first."
        Exit Sub
    End If

    Dim blockCol() As Long
    ReDim blockCol(0 To nBlocks - 1)
    Dim bi As Long, c As Long
    For bi = 0 To nBlocks - 1
        blockCol(bi) = -1
        For c = ROSTER_FIRST_TIME_COL To lastCol
            If CStr(rosterSheet.Cells(1, c).Value) = blockLabels(bi) Then
                blockCol(bi) = c
                Exit For
            End If
        Next c
    Next bi

    ' label -> comma-joined names of employees whose lunch appears to START in that slot
    ' (first LUNCH block encountered, chronologically) - used for the "no two people start
    ' lunch at once" rule.
    Dim startedAt() As String
    ReDim startedAt(0 To nBlocks - 1)
    Dim r As Long
    For r = 0 To nBlocks - 1: startedAt(r) = "": Next r

    For r = 2 To lastRow
        Dim nm As String
        nm = CStr(rosterSheet.Cells(r, 1).Value)
        Dim rl As String
        rl = CStr(rosterSheet.Cells(r, 2).Value)
        Dim ri As Long
        ri = -1
        Dim i As Long
        For i = 0 To nRoles - 1
            If roles(i) = rl Then
                ri = i
                Exit For
            End If
        Next i
        If ri = -1 Then GoTo NextR

        roleTotals(ri) = roleTotals(ri) + 1
        Dim inferredStart As Long
        inferredStart = -1
        For bi = 0 To nBlocks - 1
            If blockCol(bi) > 0 Then
                If rosterSheet.Cells(r, blockCol(bi)).Value = "LUNCH" Then
                    coverage(ri, bi) = coverage(ri, bi) + 1
                    If inferredStart = -1 Then inferredStart = bi
                End If
            End If
        Next bi
        If inferredStart >= 0 Then
            If Len(startedAt(inferredStart)) > 0 Then
                startedAt(inferredStart) = startedAt(inferredStart) & ", " & nm
            Else
                startedAt(inferredStart) = nm
            End If
        End If
NextR:
    Next r

    For bi = 0 To nBlocks - 1
        Dim cnt As Long: cnt = 0
        If Len(startedAt(bi)) > 0 Then cnt = UBound(Split(startedAt(bi), ",")) + 1
        If cnt > 1 Then
            warnCount = warnCount + 1
            ReDim Preserve warnings(1 To warnCount)
            warnings(warnCount) = "[ERROR] " & cnt & " employees are all starting lunch at " & blockLabels(bi) & _
                " (" & startedAt(bi) & ") - only one person may start lunch at a given time. Reassign one of them " & _
                "or regenerate the roster."
        End If
    Next bi

    For i = 0 To nRoles - 1
        If roleTotals(i) = 0 Then GoTo NextRole
        For bi = 0 To nBlocks - 1
            Dim onLunch As Long: onLunch = coverage(i, bi)
            Dim working As Long: working = roleTotals(i) - onLunch
            Dim lbl As String: lbl = blockLabels(bi)

            If working = 0 Then
                warnCount = warnCount + 1
                ReDim Preserve warnings(1 To warnCount)
                If roleTotals(i) >= 2 Then
                    warnings(warnCount) = "[ERROR] Role """ & roles(i) & """ is completely unstaffed during the " & lbl & _
                        " half-hour - minimum staffing floor violated. With " & roleTotals(i) & " employees in this role, " & _
                        "the roster generator should never produce this on its own; the roster was likely edited " & _
                        "manually - consider regenerating."
                Else
                    warnings(warnCount) = "[INFO] Role """ & roles(i) & """ has only one staff member, so it is unstaffed " & _
                        "during their own lunch (" & lbl & "). This is expected for single-person roles."
                End If
            End If
        Next bi
NextRole:
    Next i

    If warnCount = 0 Then ReDim warnings(1 To 1)
End Sub

Private Sub WriteDashboard(dashSheet As Worksheet, roleTotals() As Long, coverage() As Long, warnings() As String, warnCount As Long)
    ' Break apart any merged cells from a previous run before clearing - Cells.Clear does not
    ' reliably unmerge, and the row layout below shifts as the warning count changes.
    dashSheet.Cells.UnMerge
    dashSheet.Cells.Clear
    dashSheet.Cells.FormatConditions.Delete

    dashSheet.Columns(1).ColumnWidth = 26
    dashSheet.Columns(2).ColumnWidth = 18
    Dim c As Long
    For c = 3 To 8
        dashSheet.Columns(c).ColumnWidth = 12
    Next c

    With dashSheet.Range(dashSheet.Cells(1, 1), dashSheet.Cells(1, 6))
        .Merge
        .Value = "Senior Care Center - Roster Dashboard"
        .Interior.Color = ColorCharcoal()
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlLeft
    End With
    dashSheet.Rows(1).RowHeight = 28

    Dim roles As Variant: roles = RoleNames()
    With dashSheet.Range(dashSheet.Cells(3, 1), dashSheet.Cells(3, 2))
        .Value = Array("Role", "Active Staff")
        .Interior.Color = ColorCharcoal()
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    Dim i As Long
    For i = 0 To UBound(roles)
        Dim rowN As Long: rowN = 4 + i
        dashSheet.Cells(rowN, 1).Value = roles(i)
        dashSheet.Cells(rowN, 2).Formula = "=COUNTIFS('" & SHEET_STAFF & "'!B:B,""" & roles(i) & """,'" & SHEET_STAFF & "'!C:C,TRUE)"
    Next i

    Dim totalRow As Long: totalRow = 4 + (UBound(roles) - LBound(roles) + 1)
    dashSheet.Cells(totalRow, 1).Value = "Total Active Staff"
    dashSheet.Cells(totalRow, 1).Font.Bold = True
    dashSheet.Cells(totalRow, 2).Formula = "=COUNTIF('" & SHEET_STAFF & "'!C:C,TRUE)"
    dashSheet.Cells(totalRow, 2).Font.Bold = True

    Dim statusLevel As String: statusLevel = WorstLevel(warnings, warnCount)
    Dim statusRow As Long: statusRow = totalRow + 2
    dashSheet.Cells(statusRow, 1).Value = "Overall Roster Status:"
    dashSheet.Cells(statusRow, 1).Font.Bold = True
    dashSheet.Cells(statusRow, 2).Value = statusLevel
    dashSheet.Cells(statusRow, 2).Font.Bold = True
    Select Case statusLevel
        Case "ERROR": dashSheet.Cells(statusRow, 2).Interior.Color = ColorErr()
        Case "WARNING": dashSheet.Cells(statusRow, 2).Interior.Color = ColorWarning()
        Case Else: dashSheet.Cells(statusRow, 2).Interior.Color = ColorOk()
    End Select

    dashSheet.Cells(statusRow + 1, 1).Value = "Last Refreshed:"
    dashSheet.Cells(statusRow + 1, 1).Font.Bold = True
    dashSheet.Cells(statusRow + 1, 2).Value = Now
    dashSheet.Cells(statusRow + 1, 2).NumberFormat = "yyyy-mm-dd hh:mm"

    Dim warnRow As Long: warnRow = statusRow + 3
    With dashSheet.Range(dashSheet.Cells(warnRow, 1), dashSheet.Cells(warnRow, 6))
        .Merge
        .Value = "Validation Warnings"
        .Interior.Color = ColorCharcoal()
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    Dim nextRow As Long: nextRow = warnRow + 1
    If warnCount = 0 Then
        With dashSheet.Range(dashSheet.Cells(nextRow, 1), dashSheet.Cells(nextRow, 6))
            .Merge
            .Value = "No issues detected."
            .Interior.Color = ColorOk()
        End With
        nextRow = nextRow + 1
    Else
        Dim w As Long
        For w = 1 To warnCount
            Dim bgColor As Long
            If Left(warnings(w), 7) = "[ERROR]" Then
                bgColor = ColorErr()
            ElseIf Left(warnings(w), 9) = "[WARNING]" Then
                bgColor = ColorWarning()
            Else
                bgColor = RGB(238, 238, 238)
            End If
            With dashSheet.Range(dashSheet.Cells(nextRow, 1), dashSheet.Cells(nextRow, 6))
                .Merge
                .Value = warnings(w)
                .Interior.Color = bgColor
                .WrapText = True
                .VerticalAlignment = xlTop
            End With
            nextRow = nextRow + 1
        Next w
    End If

    nextRow = nextRow + 1
    With dashSheet.Range(dashSheet.Cells(nextRow, 1), dashSheet.Cells(nextRow, 6))
        .Merge
        .Value = "Lunch-Hour Coverage (headcount on lunch per 30-min block)"
        .Interior.Color = ColorCharcoal()
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    nextRow = nextRow + 1

    Dim blockLabels As Variant: blockLabels = LunchBlockLabels()
    Dim nBlocks As Long: nBlocks = UBound(blockLabels) - LBound(blockLabels) + 1

    ' Same text-vs-time-serial trap as the Roster headers - force text before writing "11:30" etc.
    dashSheet.Range(dashSheet.Cells(nextRow, 2), dashSheet.Cells(nextRow, 1 + nBlocks)).NumberFormat = "@"

    dashSheet.Cells(nextRow, 1).Value = "Role"
    Dim bi As Long
    For bi = 0 To nBlocks - 1
        dashSheet.Cells(nextRow, 2 + bi).Value = blockLabels(bi)
    Next bi
    dashSheet.Range(dashSheet.Cells(nextRow, 1), dashSheet.Cells(nextRow, 1 + nBlocks)).Font.Bold = True
    dashSheet.Range(dashSheet.Cells(nextRow, 1), dashSheet.Cells(nextRow, 1 + nBlocks)).Interior.Color = RGB(236, 239, 241)
    nextRow = nextRow + 1

    For i = 0 To UBound(roles)
        If roleTotals(i) > 0 Then
            dashSheet.Cells(nextRow, 1).Value = roles(i)
            For bi = 0 To nBlocks - 1
                dashSheet.Cells(nextRow, 2 + bi).Value = coverage(i, bi)
            Next bi
            nextRow = nextRow + 1
        End If
    Next i

    With dashSheet.Range(dashSheet.Cells(1, 1), dashSheet.Cells(nextRow, 8))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(208, 208, 208)
    End With
End Sub

Private Function WorstLevel(warnings() As String, warnCount As Long) As String
    Dim i As Long
    For i = 1 To warnCount
        If Left(warnings(i), 7) = "[ERROR]" Then
            WorstLevel = "ERROR"
            Exit Function
        End If
    Next i
    For i = 1 To warnCount
        If Left(warnings(i), 9) = "[WARNING]" Then
            WorstLevel = "WARNING"
            Exit Function
        End If
    Next i
    WorstLevel = "OK"
End Function
