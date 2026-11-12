Attribute VB_Name = "UPdateAudit"
Option Explicit
'==========================================================================================
' UPdateAudit - one-click metadata, calculation-error audit, brand shift analysis and the
'               weighted Conversion / Closing build-up for the UPdate workbook.
'
' HOW TO USE
'   1. Open the workbook. Alt+F11 > File > Import File... > choose UPdateAudit.bas
'      (or Insert > Module and paste this whole file, minus the first "Attribute" line).
'   2. Paste / refresh your Exec and trend sheets as usual.
'   3. Alt+F8 > RunFullAudit > Run.   (Save the workbook as .xlsm to keep the macro.)
'
' WHAT IT BUILDS (deleted and rebuilt on every run - nothing else is touched)
'   Audit Summary            counts, top findings, clearest brand shifts
'   Weighted Conv & Closing  LIVE formulas that rebuild Conversion / Closing, Target and
'                            W Target from the current Exec sheet + mix/rate split + calculator
'   Calc Errors              every failed check: cell, reported, re-derived, explanation
'   Brand x Item             brand x item x metric: current month vs prior month vs YTD vs target
'   Trend Shifts             last month vs prior month vs trailing-12 for every trend-sheet row
'   Meta - Sheets / Meta - Columns / Meta - Rows   metadata for every sheet, column and row
'
' SHEETS ARE FOUND BY LAYOUT, NOT BY NAME, so next month's renamed sheets still work:
'   Exec sheet  = cell C2 starts with "From Date".  Current month = single-month sheet with the
'                 latest period in B2; prior month = the next single-month sheet; YTD = B2 <> B3.
'   Trend sheet = any other sheet with a row (top 12 rows) holding >= 6 YYYYMM periods.
'==========================================================================================

Private Const OUT_SUMMARY As String = "Audit Summary"
Private Const OUT_WEIGHTED As String = "Weighted Conv & Closing"
Private Const OUT_ERRORS As String = "Calc Errors"
Private Const OUT_BRAND As String = "Brand x Item"
Private Const OUT_TREND As String = "Trend Shifts"
Private Const OUT_MS As String = "Meta - Sheets"
Private Const OUT_MC As String = "Meta - Columns"
Private Const OUT_MR As String = "Meta - Rows"

' Exec sheet column numbers -----------------------------------------------------------------
Private Const cF As Long = 6, cG As Long = 7, cH As Long = 8, cI As Long = 9, cJ As Long = 10
Private Const cL As Long = 12, cM As Long = 13, cN As Long = 14, cP As Long = 16, cQ As Long = 17
Private Const cS As Long = 19, cT As Long = 20, cU As Long = 21, cV As Long = 22, cW As Long = 23
Private Const cY As Long = 25, cZ As Long = 26, cAA As Long = 27, cAC As Long = 29, cAD As Long = 30
Private Const cAF As Long = 32, cAG As Long = 33, cAH As Long = 34, cAI As Long = 35, cAJ As Long = 36
Private Const cAL As Long = 38, cAM As Long = 39, cAN As Long = 40
Private Const cAP As Long = 42, cAQ As Long = 43, cAR As Long = 44
Private Const cAT As Long = 46, cAU As Long = 47, cAV As Long = 48, cAW As Long = 49, cAX As Long = 50
Private Const cAZ As Long = 52, cBA As Long = 53, cBB As Long = 54
Private Const EXEC_READ_COLS As Long = 60

' Findings buffer ----------------------------------------------------------------------------
Private gF() As Variant
Private gN As Long
Private gSilent As Boolean
Private gStep As String

' Exec sheets: slot 1 = current month, 2 = prior month, 3 = YTD ------------------------------
Private gData(1 To 3) As Variant
Private gName(1 To 3) As String
Private gRows(1 To 3) As Long
Private gHave(1 To 3) As Boolean
Private gPeriod(1 To 3) As String
Private gLab() As String
Private gBlk() As Long
Private gBB() As String
Private gSB() As String
Private gSP() As String
Private gNBlk(1 To 3) As Long
Private gBS() As Long
Private gBE() As Long

' Trend sheets -------------------------------------------------------------------------------
Private tCount As Long
Private tName(1 To 30) As String
Private tFrom(1 To 30) As Long
Private tTo(1 To 30) As Long
Private tCap(1 To 30) As Long
Private tDays(1 To 30) As Long
Private tStart(1 To 30) As Long
Private tLabCol(1 To 30) As Long

' Summary bits -------------------------------------------------------------------------------
Private gErrCount As Long
Private gLatestMonth As Long

'==========================================================================================
'  ENTRY POINTS
'==========================================================================================
Public Sub RunFullAudit()
    If Not gSilent Then
        If MsgBox("Rebuild the audit sheets?" & vbCrLf & vbCrLf & _
                  OUT_SUMMARY & ", " & OUT_WEIGHTED & ", " & OUT_ERRORS & ", " & OUT_BRAND & ", " & _
                  OUT_TREND & ", " & OUT_MS & ", " & OUT_MC & ", " & OUT_MR & vbCrLf & vbCrLf & _
                  "Existing copies of ONLY these sheets are replaced.", vbOKCancel + vbQuestion, "UPdate audit") <> vbOK Then Exit Sub
    End If
    Dim t0 As Double, oldCalc As Long
    t0 = Timer
    On Error GoTo Fail
    SetAppState True, oldCalc

    Status "deleting previous audit sheets"
    DeleteOutputSheets
    gN = 0
    ReDim gF(1 To 1000)
    gErrCount = 0
    gLatestMonth = 0

    Status "finding Exec sheets"
    DetectExecSheets
    If Not gHave(1) Then
        SetAppState False, oldCalc
        MsgBox "No Exec sheet found (cell C2 must start with 'From Date').", vbExclamation, "UPdate audit"
        Exit Sub
    End If
    Status "finding trend sheets"
    DetectTrendSheets

    Status "building metadata"
    BuildMetadata

    Dim s As Long
    For s = 1 To 3
        If gHave(s) Then
            Status "checking rows on " & gName(s)
            ExecRowChecks s
            Status "checking hierarchy sums on " & gName(s)
            ExecHierarchyChecks s
            Status "checking brand splits on " & gName(s)
            ExecSplitChecks s
            Status "checking repeated lines on " & gName(s)
            ExecDuplicateChecks s
            Status "reconciling brands on " & gName(s)
            ExecBrandRecon s
        End If
    Next s
    Status "cross-sheet checks"
    CrossSheetChecks
    Status "trend shifts"
    BuildTrendShifts
    Status "writing Calc Errors"
    WriteFindings
    Status "brand x item"
    BuildBrandItem
    Status "weighted conversion and closing"
    BuildWeighted
    Status "summary"
    BuildSummary

    OrderOutputSheets
    SetAppState False, oldCalc
    Application.Calculate
    Status ""
    If Not gSilent Then
        MsgBox "Audit complete in " & Format(Timer - t0, "0") & " s." & vbCrLf & _
               gN & " findings written to '" & OUT_ERRORS & "'." & vbCrLf & _
               "Start with the '" & OUT_SUMMARY & "' sheet.", vbInformation, "UPdate audit"
    End If
    Exit Sub
Fail:
    Dim msg As String
    msg = "Audit stopped while " & gStep & ": " & Err.Description & " (error " & Err.Number & ")"
    On Error Resume Next
    SetAppState False, oldCalc
    Status ""
    If gSilent Then
        Debug.Print msg
        Err.Raise vbObjectError + 1, "UPdateAudit", msg
    Else
        MsgBox msg, vbCritical, "UPdate audit"
    End If
End Sub

' Same as RunFullAudit but without message boxes (for scheduled / automated runs).
Public Sub RunFullAuditSilent()
    gSilent = True
    RunFullAudit
    gSilent = False
End Sub

' Last step reached (handy if a run stops part-way).
Public Function AuditLastStep() As String
    AuditLastStep = gStep
End Function

'==========================================================================================
'  SMALL HELPERS
'==========================================================================================
Private Sub SetAppState(ByVal starting As Boolean, ByRef oldCalc As Long)
    On Error Resume Next
    If starting Then
        oldCalc = Application.Calculation
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        Application.EnableEvents = False
        Application.Calculation = -4135          ' xlCalculationManual
    Else
        If oldCalc <> 0 Then Application.Calculation = oldCalc
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Application.EnableEvents = True
    End If
End Sub

Private Sub Status(ByVal s As String)
    If s <> "" Then gStep = s
    On Error Resume Next
    If s = "" Then
        Application.StatusBar = False
    Else
        Application.StatusBar = "UPdate audit: " & s & " ..."
    End If
End Sub

Private Function IsOutputName(ByVal nm As String) As Boolean
    Select Case nm
        Case OUT_SUMMARY, OUT_WEIGHTED, OUT_ERRORS, OUT_BRAND, OUT_TREND, OUT_MS, OUT_MC, OUT_MR, _
             "READ ME", "Key Findings"
            IsOutputName = True
        Case Else
            IsOutputName = False
    End Select
End Function

Private Sub DeleteOutputSheets()
    Dim nm As Variant
    For Each nm In Array(OUT_SUMMARY, OUT_WEIGHTED, OUT_ERRORS, OUT_BRAND, OUT_TREND, OUT_MS, OUT_MC, OUT_MR)
        If SheetExists(CStr(nm)) Then ThisWorkbook.Worksheets(CStr(nm)).Delete
    Next nm
End Sub

Private Function SheetExists(ByVal nm As String) As Boolean
    Dim ws As Worksheet
    SheetExists = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = nm Then
            SheetExists = True
            Exit Function
        End If
    Next ws
End Function

Private Function NewSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = nm
    ws.Cells.Font.Name = "Arial"
    ws.Cells.Font.Size = 10
    Set NewSheet = ws
End Function

Private Sub OrderOutputSheets()
    On Error Resume Next
    Dim order As Variant, i As Long
    order = Array(OUT_SUMMARY, OUT_WEIGHTED, OUT_ERRORS, OUT_BRAND, OUT_TREND)
    For i = UBound(order) To 0 Step -1
        If SheetExists(CStr(order(i))) Then ThisWorkbook.Worksheets(CStr(order(i))).Move Before:=ThisWorkbook.Worksheets(1)
    Next i
    ThisWorkbook.Worksheets(OUT_SUMMARY).Activate
End Sub

Private Function LastRowOf(ByVal ws As Worksheet) As Long
    LastRowOf = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
End Function

Private Function LastColOf(ByVal ws As Worksheet) As Long
    LastColOf = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
End Function

Private Function ReadBlock(ByVal ws As Worksheet, ByVal nR As Long, ByVal nC As Long) As Variant
    Dim v As Variant, a() As Variant
    If nR < 2 Then nR = 2
    If nC < 2 Then nC = 2
    v = ws.Range(ws.Cells(1, 1), ws.Cells(nR, nC)).Value
    ReadBlock = v
End Function

Private Function IsNum(ByVal v As Variant) As Boolean
    If IsError(v) Then
        IsNum = False
    Else
        Select Case VarType(v)
            Case vbDouble, vbInteger, vbLong, vbSingle, vbCurrency, vbDecimal
                IsNum = True
            Case Else
                IsNum = False
        End Select
    End If
End Function

Private Function Txt(ByVal v As Variant) As String
    If IsError(v) Then
        Txt = ""
    ElseIf VarType(v) = vbString Then
        Txt = Trim$(v)
    Else
        Txt = ""
    End If
End Function

Private Function IsPeriod(ByVal v As Variant) As Boolean
    IsPeriod = False
    If IsNum(v) Then
        If v >= 200001 And v <= 209912 Then
            If v = Int(v) Then
                If (v Mod 100) >= 1 And (v Mod 100) <= 12 Then IsPeriod = True
            End If
        End If
    End If
End Function

Private Function IsErrText(ByVal v As Variant) As Boolean
    Select Case Txt(v)
        Case "#DIV/0!", "#VALUE!", "#REF!", "#N/A", "#NAME?", "#NUM!", "#NULL!": IsErrText = True
        Case Else: IsErrText = False
    End Select
End Function

Private Function ErrName(ByVal v As Variant) As String
    If IsErrText(v) Then
        ErrName = Txt(v)
        Exit Function
    End If
    Dim n As Long, s As String
    On Error Resume Next
    s = CStr(v)                      ' an Excel error value converts to "Error 2007" etc.
    n = Val(Mid$(s, InStr(s, " ") + 1))
    If Err.Number <> 0 Then
        Err.Clear
        n = 0
    End If
    On Error GoTo 0
    Select Case n
        Case 2007: ErrName = "#DIV/0!"
        Case 2015: ErrName = "#VALUE!"
        Case 2023: ErrName = "#REF!"
        Case 2042: ErrName = "#N/A"
        Case 2029: ErrName = "#NAME?"
        Case 2036: ErrName = "#NUM!"
        Case 2000: ErrName = "#NULL!"
        Case Else: ErrName = "#ERROR"
    End Select
End Function

Private Function ColL(ByVal c As Long) As String
    Dim s As String, m As Long
    Do While c > 0
        m = (c - 1) Mod 26
        s = Chr$(65 + m) & s
        c = (c - m - 1) \ 26
    Loop
    ColL = s
End Function

Private Function BrandOf(ByVal s As String) As String
    Select Case LCase$(Trim$(s))
        Case "budget": BrandOf = "Budget"
        Case "ffw", "first for women": BrandOf = "First For Women"
        Case "dial", "dialdirect": BrandOf = "Dialdirect"
        Case "a&g", "auto & general": BrandOf = "Auto & General"
        Case "virseker": BrandOf = "Virseker"
        Case "tih multibrand": BrandOf = "TIH Multibrand"
        Case Else: BrandOf = ""
    End Select
End Function

Private Function BrandList() As Variant
    BrandList = Array("Budget", "First For Women", "Dialdirect", "Auto & General", "Virseker")
End Function

Private Function CGet(ByVal col As Collection, ByVal key As String, ByRef found As Boolean) As Variant
    On Error GoTo NotFound
    CGet = col.Item(key)
    found = True
    Exit Function
NotFound:
    found = False
    CGet = Empty
End Function

Private Sub CPut(ByVal col As Collection, ByVal key As String, ByVal v As Variant)
    On Error Resume Next
    col.Add v, key
    Err.Clear
End Sub

Private Function Max2(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then Max2 = a Else Max2 = b
End Function

Private Function Min2(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then Min2 = a Else Min2 = b
End Function

Private Sub SetTab(ByVal ws As Worksheet, ByVal clr As Long)
    On Error Resume Next
    ws.Tab.Color = clr
End Sub

Private Sub NoGrid()
    If gSilent Then Exit Sub        ' cosmetic only; skipped in unattended runs
    On Error Resume Next
    NoGrid
End Sub

Private Sub Freeze(ByVal ws As Worksheet, ByVal r As Long)
    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(r, 1).Select
    ActiveWindow.FreezePanes = True
End Sub

Private Function SafeDiv(ByVal a As Double, ByVal b As Double) As Double
    If b = 0 Then SafeDiv = 0 Else SafeDiv = a / b
End Function

Private Function QS(ByVal nm As String) As String
    QS = "'" & Replace(nm, "'", "''") & "'!"
End Function

Private Function FmtV(ByVal v As Variant) As String
    If IsNum(v) Then
        If Abs(v) < 5 And v <> Int(v) Then
            FmtV = Format(v, "0.0%")
        Else
            FmtV = Format(v, "#,##0")
        End If
    Else
        FmtV = Txt(v)
    End If
End Function

Private Sub StyleHeader(ByVal ws As Worksheet, ByVal r As Long, ByVal nC As Long)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, nC))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 120)
        .WrapText = True
        .VerticalAlignment = -4108      ' xlCenter
    End With
End Sub

Private Sub Banner(ByVal ws As Worksheet, ByVal r As Long, ByVal text As String, ByVal nC As Long)
    ws.Cells(r, 1).Value = text
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, nC))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 120)
    End With
End Sub

' Writes a 2-D array (1-based) with a header row, freezes panes and adds a filter.
Private Sub DumpTable(ByVal ws As Worksheet, ByVal topRow As Long, hdr As Variant, data As Variant, ByVal nRows As Long)
    Dim nC As Long, j As Long
    nC = UBound(hdr) - LBound(hdr) + 1
    For j = 1 To nC
        ws.Cells(topRow, j).Value = hdr(LBound(hdr) + j - 1)
    Next j
    StyleHeader ws, topRow, nC
    If nRows > 0 Then
        ws.Range(ws.Cells(topRow + 1, 1), ws.Cells(topRow + nRows, nC)).Value = data
    End If
    On Error Resume Next
    ws.Range(ws.Cells(topRow, 1), ws.Cells(topRow + IIf(nRows > 0, nRows, 1), nC)).AutoFilter
    On Error GoTo 0
    Freeze ws, topRow + 1
End Sub

'==========================================================================================
'  FINDINGS
'==========================================================================================
Private Sub AddF(ByVal sev As String, ByVal cat As String, ByVal sh As String, ByVal ref As String, _
                 ByVal item As String, ByVal brand As String, ByVal check As String, _
                 ByVal reported As Variant, ByVal expected As Variant, ByVal note As String)
    Dim d As Variant, i As Long, key As String
    If IsNum(reported) And IsNum(expected) Then d = reported - expected Else d = Empty
    ' merge identical defects repeated in several copies of a block
    key = sh & "|" & cat & "|" & item & "|" & brand & "|" & check & "|" & CStr(reported) & "|" & CStr(expected)
    For i = gN To IIf(gN > 400, gN - 400, 1) Step -1
        If gF(i)(11) = key Then
            gF(i)(3) = gF(i)(3) & ", " & ref
            Exit Sub
        End If
    Next i
    gN = gN + 1
    If gN > UBound(gF) Then
        ReDim Preserve gF(1 To UBound(gF) * 2)
    End If
    gF(gN) = Array(sev, cat, sh, ref, item, brand, check, reported, expected, d, note, key)
End Sub

Private Sub WriteFindings()
    Dim ws As Worksheet, data() As Variant, i As Long, j As Long, k As Long, sev As Variant, n As Long
    Set ws = NewSheet(OUT_ERRORS)
    Dim hdr As Variant
    hdr = Array("ID", "Severity", "Category", "Sheet", "Cell / Row", "Item", "Brand", "Check", _
                "Reported", "Expected / Re-derived", "Difference", "Explanation / Impact")
    If gN = 0 Then
        DumpTable ws, 1, hdr, Empty, 0
        ws.Cells(2, 1).Value = "No findings."
        Exit Sub
    End If
    ReDim data(1 To gN, 1 To 12)
    n = 0
    For Each sev In Array("High", "Medium", "Low", "Info")
        For i = 1 To gN
            If gF(i)(0) = sev Then
                n = n + 1
                data(n, 1) = "F" & Format(n, "0000")
                For j = 0 To 10
                    data(n, j + 2) = gF(i)(j)
                Next j
            End If
        Next i
    Next sev
    DumpTable ws, 1, hdr, data, n
    ws.Columns(1).ColumnWidth = 8
    ws.Columns(2).ColumnWidth = 9
    ws.Columns(3).ColumnWidth = 20
    ws.Columns(4).ColumnWidth = 20
    ws.Columns(5).ColumnWidth = 20
    ws.Columns(6).ColumnWidth = 34
    ws.Columns(7).ColumnWidth = 15
    ws.Columns(8).ColumnWidth = 55
    ws.Columns(9).ColumnWidth = 13
    ws.Columns(10).ColumnWidth = 13
    ws.Columns(11).ColumnWidth = 13
    ws.Columns(12).ColumnWidth = 90
    ws.Range(ws.Cells(2, 9), ws.Cells(n + 1, 11)).NumberFormat = "#,##0.0000;-#,##0.0000;0"
    ws.Range(ws.Cells(2, 5), ws.Cells(n + 1, 5)).WrapText = True
    ws.Range(ws.Cells(2, 8), ws.Cells(n + 1, 8)).WrapText = True
    ws.Range(ws.Cells(2, 12), ws.Cells(n + 1, 12)).WrapText = True
    For k = 2 To n + 1
        Select Case ws.Cells(k, 2).Value
            Case "High": ws.Cells(k, 2).Interior.Color = RGB(248, 203, 173)
            Case "Medium": ws.Cells(k, 2).Interior.Color = RGB(255, 230, 153)
            Case "Low": ws.Cells(k, 2).Interior.Color = RGB(242, 242, 242)
            Case "Info": ws.Cells(k, 2).Interior.Color = RGB(221, 235, 247)
        End Select
    Next k
    SetTab ws, RGB(192, 0, 0)
End Sub

'==========================================================================================
'  EXEC SHEETS: detection and parsing
'==========================================================================================
Private Sub DetectExecSheets()
    Dim ws As Worksheet, b2 As Variant, b3 As Variant
    Dim curP As Double, prvP As Double, ytdP As Double
    Dim curWs As Worksheet, prvWs As Worksheet, ytdWs As Worksheet
    Dim s As Long
    For s = 1 To 3
        gHave(s) = False
        gName(s) = ""
    Next s
    curP = 0: prvP = 0: ytdP = 0
    ' current month = latest single-month sheet
    For Each ws In ThisWorkbook.Worksheets
        If Not IsOutputName(ws.Name) Then
            If Left$(Txt(ws.Range("C2").Value), 9) = "From Date" Then
                b2 = ws.Range("B2").Value: b3 = ws.Range("B3").Value
                If IsNum(b2) And IsNum(b3) Then
                    If b2 = b3 Then
                        If b2 > curP Then
                            curP = b2
                            Set curWs = ws
                        End If
                    ElseIf b3 > ytdP Then
                        ytdP = b3: Set ytdWs = ws
                    End If
                End If
            End If
        End If
    Next ws
    If curWs Is Nothing Then Exit Sub
    For Each ws In ThisWorkbook.Worksheets
        If Not IsOutputName(ws.Name) And ws.Name <> curWs.Name Then
            If Left$(Txt(ws.Range("C2").Value), 9) = "From Date" Then
                b2 = ws.Range("B2").Value: b3 = ws.Range("B3").Value
                If IsNum(b2) And IsNum(b3) Then
                    If b2 = b3 And b2 < curP And b2 > prvP Then
                        prvP = b2
                        Set prvWs = ws
                    End If
                End If
            End If
        End If
    Next ws

    Dim maxR As Long
    maxR = LastRowOf(curWs)
    If Not prvWs Is Nothing Then If LastRowOf(prvWs) > maxR Then maxR = LastRowOf(prvWs)
    If Not ytdWs Is Nothing Then If LastRowOf(ytdWs) > maxR Then maxR = LastRowOf(ytdWs)
    ReDim gLab(1 To 3, 1 To maxR)
    ReDim gBlk(1 To 3, 1 To maxR)
    ReDim gBB(1 To 3, 1 To maxR)
    ReDim gSB(1 To 3, 1 To maxR)
    ReDim gSP(1 To 3, 1 To maxR)
    ReDim gBS(1 To 3, 0 To 400)
    ReDim gBE(1 To 3, 0 To 400)

    LoadExec 1, curWs
    If Not prvWs Is Nothing Then LoadExec 2, prvWs
    If Not ytdWs Is Nothing Then LoadExec 3, ytdWs
End Sub

Private Sub LoadExec(ByVal s As Long, ByVal ws As Worksheet)
    gName(s) = ws.Name
    gRows(s) = LastRowOf(ws)
    gData(s) = ReadBlock(ws, gRows(s), EXEC_READ_COLS)
    gHave(s) = True
    gPeriod(s) = CStr(ws.Range("B2").Value) & IIf(ws.Range("B2").Value = ws.Range("B3").Value, "", "-" & CStr(ws.Range("B3").Value))
    ParseExec s
End Sub

Private Sub ParseExec(ByVal s As Long)
    Dim arr As Variant, r As Long, lab As String, tg As String, blk As Long
    Dim brand As String, parent As String, rb As String, isBrandRow As Boolean
    arr = gData(s)
    blk = 0: brand = "": parent = ""
    gBS(s, 0) = 1
    For r = 1 To gRows(s)
        lab = Txt(arr(r, 3))
        tg = Txt(arr(r, 2))
        gLab(s, r) = "": gBB(s, r) = "": gSB(s, r) = "": gSP(s, r) = ""
        If Left$(lab, 9) = "From Date" Then
            gBE(s, blk) = r - 1
            blk = blk + 1
            gBS(s, blk) = r
            brand = "": parent = ""
            gBlk(s, r) = blk
        ElseIf Left$(lab, 7) = "To Date" Then
            gBlk(s, r) = blk
        Else
            gBlk(s, r) = blk
            If Right$(lab, 11) = "Brand Total" Then brand = BrandOf(Replace(lab, " - Brand Total", ""))
            rb = BrandOf(tg)
            isBrandRow = (rb <> "" And lab <> "" And BrandOf(lab) = rb)
            If lab <> "" And Not isBrandRow Then parent = lab
            gLab(s, r) = lab
            gBB(s, r) = brand
            If isBrandRow Then
                gSB(s, r) = rb
                gSP(s, r) = parent
            End If
        End If
    Next r
    gBE(s, blk) = gRows(s)
    gNBlk(s) = blk
End Sub

Private Function EV(ByVal s As Long, ByVal r As Long, ByVal c As Long) As Variant
    EV = gData(s)(r, c)
End Function

Private Function RowHasNum(ByVal s As Long, ByVal r As Long) As Boolean
    Dim c As Long
    RowHasNum = False
    For c = cF To cBB
        If IsNum(gData(s)(r, c)) Then
            RowHasNum = True
            Exit Function
        End If
    Next c
End Function

' First row in block b with this label (non brand-split row), after row 'after'. 0 = not found.
Private Function FindInBlock(ByVal s As Long, ByVal b As Long, ByVal lab As String, ByVal after As Long) As Long
    Dim r As Long
    FindInBlock = 0
    If b < 0 Or b > gNBlk(s) Then Exit Function
    For r = gBS(s, b) To gBE(s, b)
        If r > after Then
            If gLab(s, r) = lab And gSB(s, r) = "" Then
                FindInBlock = r
                Exit Function
            End If
        End If
    Next r
End Function

Private Function ItemName(ByVal s As Long, ByVal r As Long) As String
    If gSB(s, r) <> "" Then ItemName = gSP(s, r) & " > " & gLab(s, r) Else ItemName = gLab(s, r)
End Function

Private Function RowBrand(ByVal s As Long, ByVal r As Long) As String
    If gBB(s, r) <> "" Then RowBrand = gBB(s, r) Else RowBrand = gSB(s, r)
End Function

Private Function ProjFactor(ByVal s As Long) As Double
    Dim v() As Double, n As Long, r As Long, i As Long, j As Long, tmp As Double
    ReDim v(1 To gRows(s))
    For r = 4 To gRows(s)
        If IsNum(EV(s, r, cF)) And IsNum(EV(s, r, cG)) Then
            If EV(s, r, cF) > 2000 Then
                n = n + 1
                v(n) = EV(s, r, cG) / EV(s, r, cF)
            End If
        End If
    Next r
    If n = 0 Then
        ProjFactor = 1
        Exit Function
    End If
    For i = 2 To n
        tmp = v(i): j = i - 1
        Do While j >= 1
            If v(j) <= tmp Then Exit Do
            v(j + 1) = v(j): j = j - 1
        Loop
        v(j + 1) = tmp
    Next i
    If n Mod 2 = 1 Then ProjFactor = v((n + 1) \ 2) Else ProjFactor = (v(n \ 2) + v(n \ 2 + 1)) / 2
End Function

'==========================================================================================
'  EXEC CHECKS
'==========================================================================================
Private Sub ExecRowChecks(ByVal s As Long)
    Dim sh As String, fac As Double, r As Long, g As Long, it As String, br As String
    Dim VA As Variant, VP As Variant, VT As Variant, VD As Variant, VV As Variant, VN As Variant
    Dim RA As Variant, RT As Variant, RV As Variant, RW As Variant, RWV As Variant, RN As Variant
    Dim A As Variant, P As Variant, T As Variant, D As Variant, V As Variant, W As Variant, WV As Variant
    Dim c As Long, cells As String
    sh = gName(s)
    fac = ProjFactor(s)
    VA = Array(cF, cS, cAF, cAT): VP = Array(cG, cT, cAG, cAU): VT = Array(cH, cU, cAH, cAV)
    VD = Array(cI, cV, cAI, cAW): VV = Array(cJ, cW, cAJ, cAX)
    VN = Array("Contacts", "Leads", "Gross Sales", "NB Net Sales")
    RA = Array(cL, cY, cAL, cAP, cAZ): RT = Array(cM, cZ, cAM, cAQ, cBA): RV = Array(cN, cAA, cAN, cAR, cBB)
    RW = Array(cP, cAC, 0, 0, 0): RWV = Array(cQ, cAD, 0, 0, 0)
    RN = Array("Conversion %", "Gross Closing %", "Gross Effectiveness %", "NB Net CXD %", "Net Effectiveness %")

    For r = 4 To gRows(s)
        If gLab(s, r) = "" Then
            If RowHasNum(s, r) And Left$(Txt(EV(s, r, 3)), 4) <> "From" And Left$(Txt(EV(s, r, 3)), 2) <> "To" Then
                cells = ""
                For c = cF To cBB
                    If IsNum(EV(s, r, c)) Then cells = cells & IIf(cells = "", "", ", ") & ColL(c) & r & "=" & EV(s, r, c)
                Next c
                AddF "Medium", "Stray value", sh, "Row " & r, "(no label)", "", "Number without a row label", cells, Empty, _
                     "Orphan number between blocks - a leftover / hard-typed value that belongs to no line. Delete or label it."
            End If
            GoTo NextRow
        End If
        it = ItemName(s, r): br = RowBrand(s, r)

        ' ---- volumes -----------------------------------------------------------------
        For g = 0 To 3
            A = EV(s, r, VA(g)): P = EV(s, r, VP(g)): T = EV(s, r, VT(g)): D = EV(s, r, VD(g)): V = EV(s, r, VV(g))
            If IsNum(P) And IsNum(T) And IsNum(D) Then
                If Abs((P - T) - D) > 1.01 Then AddF "High", "Arithmetic", sh, ColL(VD(g)) & r, it, br, _
                    VN(g) & ": Tar Dif = Projecting - Target", D, P - T, "Target difference does not equal Projecting minus Target."
            End If
            If IsNum(D) And IsNum(T) And IsNum(V) Then
                If T <> 0 Then
                    If Abs(D / T - V) > 0.0006 Then AddF "High", "Arithmetic", sh, ColL(VV(g)) & r, it, br, _
                        VN(g) & ": Tar Var = Tar Dif / Target", V, D / T, "Target variance % does not equal Tar Dif / Target."
                End If
            End If
            If IsNum(A) And IsNum(P) Then
                If A >= 20 Then
                    If Abs(P - A * fac) > Max2(3.01, 0.01 * A) Then
                        AddF "High", "Projection", sh, ColL(VP(g)) & r, it, br, VN(g) & ": Projecting = Actual x " & Format(fac, "0.00000"), _
                             P, Round(A * fac, 0), "Projected value is inconsistent with the run-rate factor used on every other line - " & _
                             "the Actual (or the Projecting) on this line is pulled from a different row."
                    End If
                End If
            End If
            If IsNum(T) Then
                If T < 0 Then AddF "Medium", "Sanity", sh, ColL(VT(g)) & r, it, br, VN(g) & ": target must be >= 0", T, 0, _
                    "Negative target - the % variance next to it is meaningless."
            End If
            If IsNum(A) Then
                If A < 0 Then AddF "Medium", "Sanity", sh, ColL(VA(g)) & r, it, br, VN(g) & ": actual must be >= 0", A, 0, _
                    "Negative sales (cancellations exceed new business on a tiny base) - ratios on this line are distorted."
            End If
        Next g

        ' ---- rates -------------------------------------------------------------------
        For g = 0 To 4
            A = EV(s, r, RA(g)): T = EV(s, r, RT(g)): V = EV(s, r, RV(g))
            If IsNum(A) And IsNum(T) And IsNum(V) Then
                If T <> 0 Then
                    If Abs(A / T - 1 - V) > 0.0006 Then AddF "High", "Arithmetic", sh, ColL(RV(g)) & r, it, br, _
                        RN(g) & ": Tar Var = Actual / Target - 1", V, A / T - 1, "Rate variance does not equal Actual / Target - 1."
                End If
            End If
            If RW(g) > 0 Then
                W = EV(s, r, RW(g)): WV = EV(s, r, RWV(g))
                If IsNum(A) And IsNum(W) And IsNum(WV) Then
                    If W <> 0 Then
                        If Abs(A / W - 1 - WV) > 0.0006 Then AddF "High", "Arithmetic", sh, ColL(RWV(g)) & r, it, br, _
                            RN(g) & ": W Tar Var = Actual / W Target - 1", WV, A / W - 1, "Weighted-target variance does not equal Actual / W Target - 1."
                    End If
                End If
                If IsNum(W) And IsNum(T) Then
                    If W = 0 And T > 0 Then
                        AddF "High", "Weighted target", sh, ColL(RW(g)) & r, it, br, _
                             RN(g) & ": W Target collapsed to 0 while Target = " & Format(T, "0.0%"), W, T, _
                             "Weighted target is 0 although a plan target exists: W variance is unusable and this line feeds 0% into " & _
                             "every parent weighted target (understates the parent W Target)."
                    ElseIf W > 0 And T > 0 Then
                        If W / T < 0.5 Or W / T > 2 Then
                            AddF "High", "Weighted target", sh, ColL(RW(g)) & r, it, br, _
                                 RN(g) & ": W Target " & Format(W, "0.00%") & " vs Target " & Format(T, "0.00%") & " (x" & Format(W / T, "0.00") & ")", W, T, _
                                 "Weighted target is implausibly far from plan - child lines with volume but a 0 / missing target are " & _
                                 "averaged in at 0%. The W Tar Var on this line is not a real performance signal."
                        End If
                    End If
                End If
            End If
        Next g

        ' ---- re-derived ratios -------------------------------------------------------
        If IsNum(EV(s, r, cS)) And IsNum(EV(s, r, cAF)) And IsNum(EV(s, r, cY)) Then
            If EV(s, r, cS) > 0 Then
                If Abs(EV(s, r, cAF) / EV(s, r, cS) - EV(s, r, cY)) > 0.0006 Then AddF "High", "Arithmetic", sh, "Y" & r, it, br, _
                    "Gross Closing = Gross Sales / Leads", EV(s, r, cY), EV(s, r, cAF) / EV(s, r, cS), "Closing % does not equal Gross Sales / Leads."
            End If
        End If
        If IsNum(EV(s, r, cU)) And IsNum(EV(s, r, cAH)) And IsNum(EV(s, r, cZ)) Then
            If EV(s, r, cU) > 0 Then
                If Abs(EV(s, r, cAH) / EV(s, r, cU) - EV(s, r, cZ)) > 0.0006 Then AddF "High", "Arithmetic", sh, "Z" & r, it, br, _
                    "Closing Target = Sales Target / Leads Target", EV(s, r, cZ), EV(s, r, cAH) / EV(s, r, cU), "Closing target is not Sales target / Leads target."
            End If
        End If
        If IsNum(EV(s, r, cAF)) And IsNum(EV(s, r, cAT)) And IsNum(EV(s, r, cAP)) Then
            If EV(s, r, cAF) > 0 Then
                If Abs(1 - EV(s, r, cAT) / EV(s, r, cAF) - EV(s, r, cAP)) > 0.0006 Then AddF "High", "Arithmetic", sh, "AP" & r, it, br, _
                    "NB Net CXD = 1 - Net Sales / Gross Sales", EV(s, r, cAP), 1 - EV(s, r, cAT) / EV(s, r, cAF), "CXD % does not equal 1 - Net / Gross."
            End If
        End If
        If IsNum(EV(s, r, cAH)) And IsNum(EV(s, r, cAV)) And IsNum(EV(s, r, cAQ)) Then
            If EV(s, r, cAH) > 0 Then
                If Abs(1 - EV(s, r, cAV) / EV(s, r, cAH) - EV(s, r, cAQ)) > 0.0006 Then AddF "Medium", "Arithmetic", sh, "AQ" & r, it, br, _
                    "CXD Target = 1 - Net Target / Gross Target", EV(s, r, cAQ), 1 - EV(s, r, cAV) / EV(s, r, cAH), "CXD target inconsistent with net and gross sales targets."
            End If
        End If

        ' ---- sanity ------------------------------------------------------------------
        If IsNum(EV(s, r, cL)) Then If EV(s, r, cL) > 1.000000001 Then AddF "High", "Sanity", sh, "L" & r, it, br, "Conversion % above 100%", EV(s, r, cL), 1, _
            "More leads than contacts: numerator and denominator come from different populations - the rate is not usable."
        If IsNum(EV(s, r, cAL)) Then If EV(s, r, cAL) > 1.000000001 Then AddF "High", "Sanity", sh, "AL" & r, it, br, "Gross Effectiveness % above 100%", EV(s, r, cAL), 1, _
            "More sales than contacts - numerator and denominator come from different populations."
        If IsNum(EV(s, r, cAZ)) Then If EV(s, r, cAZ) > 1.000000001 Then AddF "High", "Sanity", sh, "AZ" & r, it, br, "Net Effectiveness % above 100%", EV(s, r, cAZ), 1, _
            "More net sales than contacts - numerator and denominator come from different populations."
        If IsNum(EV(s, r, cY)) Then If EV(s, r, cY) > 1.000000001 Then AddF "Low", "Sanity", sh, "Y" & r, it, br, "Gross Closing % above 100%", EV(s, r, cY), 1, _
            "More sales than leads (usually timing of sales vs leads on a small line)."
        If (IsNum(EV(s, r, cL)) Or IsNum(EV(s, r, cY))) And Not IsNum(EV(s, r, cF)) And Not IsNum(EV(s, r, cS)) Then
            AddF "Medium", "Completeness", sh, "Row " & r, it, br, "Rates shown but Contacts / Leads volumes are blank", Empty, Empty, _
                 "Conversion / closing % shown with no underlying volumes - cannot be verified or re-weighted."
        End If
NextRow:
    Next r
End Sub

Private Function HierarchyRules() As Variant
    HierarchyRules = Array( _
        "Total Personal Lines Incl MBO|Direct;Non-Direct", _
        "Direct|Direct (Excl Outbound);DM Multibrand Outbound", _
        "Direct (Incl Outbound)|Direct (Excl Outbound);DM Multibrand Outbound", _
        "Direct (Excl Outbound)|Traditional Combined;Competitions Combined;OQ;Direct Marketing;Affiliates;Hippo;Better Compare;Cover Compare;Direct Other Clean Up", _
        "DM Multibrand Outbound|DM - Branded Outbound;DM - Multibrand Outbound;Referral", _
        "Traditional Combined|Generic - Traditional;E-Commerce - Traditional", _
        "Competitions Combined|Generic - Competition;E-Commerce - Competition", _
        "Generic|Generic - Traditional;Generic - Competition", _
        "E-Commerce|E-Commerce - Traditional;E-Commerce - Competition", _
        "OQ|OQ - PNV;OQ - PV;OQ - End to End", _
        "Hippo|Hippo - Branded;Hippo - Unconverted;Hippo - Unbranded", _
        "Non-Direct|Affinities;Bank;Dealer Channel;Broker;Sales Partnerships;Property Partners;Tied Agency;Other Cleanup", _
        "Dealer Channel|Multibrand;Dealer Groups;Independent Dealers;Virtual Ops;Zambesi Dealers;Provincial Groups;Dealer Cleanup", _
        "Broker|Traditional Broker;Dealer Broker;OCC Broker;Broker Bank", _
        "Property Partners|Balwin;Landsdowne;Waterfall", _
        "Tied Agency|Tied Agency Academy;Tied Agency F2F;Tied Agency Virtual;Agency F2F;Agency Lite;Agency Virtual", _
        "Bank|Nedbank Budget;Nedbank FB;Nedbank BB;NGIB;SBIB;ABSA", _
        "Affinities|Insurance Exchange;OEM;Fidelity ADT;1Life Motor and Home;Virgin Money MQ;Affinities Other;Kandua;Namola", _
        "OEM|MyToyota Insurance;Volkswagen Financial Services;Audi Financial Services;Toyota Financial Services")
End Function

Private Sub ExecHierarchyChecks(ByVal s As Long)
    Dim rules As Variant, b As Long, k As Long, parts As Variant, parent As String, kids As Variant
    Dim prow As Long, kr As Long, anywhere As Boolean, nPresent As Long, i As Long
    Dim present() As Long, missing As String, presentNames As String
    Dim sc As Variant, sn As Variant, j As Long, pv As Variant, sm As Double, sev As String, material As Boolean
    sc = Array(cF, cS, cAF, cAT, cH, cU, cAH, cAV)
    sn = Array("Contacts", "Leads", "Gross Sales", "NB Net Sales", "Contacts Target", "Leads Target", "Gross Sales Target", "Net Sales Target")
    rules = HierarchyRules()
    For b = 1 To gNBlk(s)
        For k = LBound(rules) To UBound(rules)
            parts = Split(rules(k), "|")
            parent = parts(0)
            kids = Split(parts(1), ";")
            prow = FindInBlock(s, b, parent, 0)
            If prow = 0 Then GoTo NextRule
            anywhere = (parent = "Total Personal Lines Incl MBO" Or parent = "Direct" Or parent = "Direct (Incl Outbound)" Or parent = "Non-Direct")
            ReDim present(0 To UBound(kids))
            nPresent = 0: missing = "": presentNames = ""
            For i = 0 To UBound(kids)
                If anywhere Then kr = FindInBlock(s, b, CStr(kids(i)), 0) Else kr = FindInBlock(s, b, CStr(kids(i)), prow)
                If kr > 0 Then
                    present(nPresent) = kr: nPresent = nPresent + 1
                    presentNames = presentNames & IIf(presentNames = "", "", ", ") & kids(i)
                ElseIf FindInBlock(s, b, CStr(kids(i)), 0) = 0 Then
                    missing = missing & IIf(missing = "", "", ", ") & kids(i)
                End If
            Next i
            If nPresent < 2 Then GoTo NextRule
            For j = 0 To UBound(sc)
                pv = EV(s, prow, sc(j))
                If IsNum(pv) Then
                    sm = 0
                    For i = 0 To nPresent - 1
                        If IsNum(EV(s, present(i), sc(j))) Then sm = sm + EV(s, present(i), sc(j))
                    Next i
                    If Abs(pv - sm) > 1.01 Then
                        material = Abs(pv - sm) > Max2(5, 0.01 * Abs(pv))
                        If material Then
                            If missing = "" Then sev = "High" Else sev = "Medium"
                        Else
                            sev = "Low"
                        End If
                        If pv = 0 And sm > 0 Then sev = "High"
                        AddF sev, "Hierarchy sum", gName(s), ColL(CLng(sc(j))) & prow, parent, gBB(s, prow), _
                             sn(j) & ": " & parent & " = sum of " & presentNames, pv, sm, _
                             "Parent line does not equal the sum of its child lines in the same block" & _
                             IIf(missing <> "", " (children not shown in this block: " & missing & ")", "") & "."
                    End If
                End If
            Next j
NextRule:
        Next k
    Next b
End Sub

Private Sub ExecSplitChecks(ByVal s As Long)
    Dim keys As Collection, grp() As String, gFirst() As Long, nG As Long
    Set keys = New Collection
    Dim r As Long, key As String, found As Boolean, idx As Variant, i As Long, rr As Long
    Dim prow As Long, rowsArr As Variant, sc As Variant, sn As Variant, j As Long, pv As Variant, sm As Double, bk As Long
    ReDim grp(1 To 200): ReDim gFirst(1 To 200)
    For r = 4 To gRows(s)
        If gSB(s, r) <> "" Then
            key = gBlk(s, r) & "|" & gSP(s, r)
            idx = CGet(keys, key, found)
            If Not found Then
                nG = nG + 1
                If nG > UBound(grp) Then
                    ReDim Preserve grp(1 To nG * 2)
                    ReDim Preserve gFirst(1 To nG * 2)
                End If
                grp(nG) = CStr(r): gFirst(nG) = r
                keys.Add nG, key
            Else
                grp(idx) = grp(idx) & "," & r
            End If
        End If
    Next r
    sc = Array(cF, cS, cAF, cAT, cH, cU, cAH, cAV)
    sn = Array("Contacts", "Leads", "Gross Sales", "NB Net Sales", "Contacts Target", "Leads Target", "Gross Sales Target", "Net Sales Target")
    For i = 1 To nG
        r = gFirst(i)
        prow = 0
        bk = gBlk(s, r)
        For rr = gBS(s, bk) To r - 1
            If gLab(s, rr) = gSP(s, r) And gSB(s, rr) = "" Then prow = rr
        Next rr
        If prow = 0 Then GoTo NextG
        rowsArr = Split(grp(i), ",")
        For j = 0 To UBound(sc)
            pv = EV(s, prow, sc(j))
            If IsNum(pv) Then
                sm = 0
                For rr = 0 To UBound(rowsArr)
                    If IsNum(EV(s, CLng(rowsArr(rr)), sc(j))) Then sm = sm + EV(s, CLng(rowsArr(rr)), sc(j))
                Next rr
                If Abs(pv - sm) > 1.01 Then
                    AddF IIf(Abs(pv - sm) > 0.01 * Abs(pv), "High", "Medium"), "Brand split sum", gName(s), ColL(CLng(sc(j))) & prow, gSP(s, r), "", _
                         sn(j) & ": " & gSP(s, r) & " = sum of its brand rows " & rowsArr(0) & "-" & rowsArr(UBound(rowsArr)), pv, sm, _
                         "The brand rows underneath do not add up to the line they split (difference " & Format(pv - sm, "+#,##0;-#,##0") & _
                         ") - volume not allocated to any brand, or a wrong link."
                End If
            End If
        Next j
NextG:
    Next i
End Sub

Private Sub ExecDuplicateChecks(ByVal s As Long)
    Dim first As Collection, r As Long, key As String, found As Boolean, b As Variant
    Set first = New Collection
    Dim cols As Variant, nms As Variant, j As Long, x As Variant, y As Variant, tol As Double
    cols = Array(cF, cS, cAF, cAT, cG, cL, cY)
    nms = Array("Contacts", "Leads", "Gross Sales", "NB Net Sales", "Contacts Projecting", "Conversion %", "Gross Closing %")
    For r = 4 To gRows(s)
        If gLab(s, r) <> "" And gBB(s, r) = "" Then
            key = gSP(s, r) & "|" & gLab(s, r)
            b = CGet(first, key, found)
            If Not found Then
                first.Add r, key
            Else
                For j = 0 To UBound(cols)
                    x = EV(s, CLng(b), cols(j)): y = EV(s, r, cols(j))
                    If cols(j) = cL Or cols(j) = cY Then tol = 0.0006 Else tol = 0.5
                    If IsNum(x) And IsNum(y) Then
                        If Abs(x - y) > tol Then
                            AddF "High", "Copy / link error", gName(s), ColL(CLng(cols(j))) & r, ItemName(s, r), gSB(s, r), _
                                 nms(j) & ": row " & r & " should equal row " & b & " (same line repeated)", y, x, _
                                 "The same line appears twice with different values - one cell is linked to the wrong source row."
                        End If
                    End If
                Next j
            End If
        End If
    Next r
End Sub

Private Sub ExecBrandRecon(ByVal s As Long)
    Dim allB As Collection, perB As Collection, labs() As String, nL As Long
    Set allB = New Collection
    Set perB = New Collection
    Dim r As Long, lab As String, found As Boolean, v As Variant, i As Long, k As Long
    Dim brands As Variant, bRows(0 To 4) As Long, nb As Long, arow As Long
    Dim cols As Variant, nms As Variant, j As Long, sm As Double, d As Double, isMB As Boolean, sev As String
    ReDim labs(1 To 500)
    For r = 4 To gRows(s)
        lab = gLab(s, r)
        If lab <> "" And gSB(s, r) = "" Then
            If gBB(s, r) <> "" Then
                If Right$(lab, 11) = "Brand Total" Then lab = "Total Personal Lines Incl MBO"
                v = CGet(perB, gBB(s, r) & "|" & lab, found)
                If Not found Then perB.Add r, gBB(s, r) & "|" & lab
                v = CGet(perB, "#" & lab, found)
                If Not found Then
                    perB.Add 1, "#" & lab
                    nL = nL + 1
                    If nL > UBound(labs) Then
                        ReDim Preserve labs(1 To nL * 2)
                    End If
                    labs(nL) = lab
                End If
            Else
                CPut allB, lab, r
            End If
        End If
    Next r
    brands = BrandList()
    cols = Array(cF, cS, cAF, cAT, cH, cAH)
    nms = Array("Contacts", "Leads", "Gross Sales", "NB Net Sales", "Contacts Target", "Gross Sales Target")
    For i = 1 To nL
        v = CGet(allB, labs(i), found)
        If Not found Then GoTo NextL
        arow = v
        nb = 0
        For k = 0 To 4
            v = CGet(perB, brands(k) & "|" & labs(i), found)
            If found Then
                bRows(nb) = v
                nb = nb + 1
            End If
        Next k
        If nb < 5 Then GoTo NextL
        isMB = (labs(i) = "Total Personal Lines Incl MBO" Or labs(i) = "Direct" Or labs(i) = "DM Multibrand Outbound" Or labs(i) = "DM - Multibrand Outbound")
        For j = 0 To UBound(cols)
            If IsNum(EV(s, arow, cols(j))) Then
                sm = 0
                For k = 0 To 4
                    If IsNum(EV(s, bRows(k), cols(j))) Then sm = sm + EV(s, bRows(k), cols(j))
                Next k
                d = EV(s, arow, cols(j)) - sm
                If Abs(d) > 1.01 Then
                    If isMB Then
                        sev = "Info"
                    ElseIf Abs(d) > Max2(5, 0.01 * Abs(EV(s, arow, cols(j)))) Then
                        sev = "High"
                    Else
                        sev = "Low"
                    End If
                    AddF sev, "Brand reconciliation", gName(s), ColL(CLng(cols(j))) & arow, labs(i), "All brands", _
                         nms(j) & ": all-brand row " & arow & " vs sum of the 5 brand pages", EV(s, arow, cols(j)), sm, _
                         IIf(isMB, "Gap = Multibrand (TIH) volume not allocated to a brand - expected, but brand pages do not add to the group total.", _
                         "Brand pages do not add up to the all-brand line - volume on the group line not allocated to any brand page (or double counted).")
                End If
            End If
        Next j
NextL:
    Next i
End Sub

'==========================================================================================
'  TREND SHEETS
'==========================================================================================
Private Sub DetectTrendSheets()
    Dim ws As Worksheet, arr As Variant, r As Long, c As Long, nP As Long, nR As Long, nC As Long
    Dim pr(1 To 12) As Long, nPR As Long, nSmall As Long, cnt2 As Long, cnt3 As Long
    tCount = 0
    For Each ws In ThisWorkbook.Worksheets
        If Not IsOutputName(ws.Name) And Left$(Txt(ws.Range("C2").Value), 9) <> "From Date" Then
            nR = LastRowOf(ws): nC = LastColOf(ws)
            If nR >= 6 And nC >= 6 Then
                If nR > 12 Then nR = 12
                arr = ReadBlock(ws, nR, nC)
                nPR = 0
                For r = 1 To nR
                    nP = 0
                    For c = 1 To nC
                        If IsPeriod(arr(r, c)) Then nP = nP + 1
                    Next c
                    If nP >= 6 Then
                        nPR = nPR + 1
                        pr(nPR) = r
                    End If
                Next r
                If nPR > 0 And tCount < 30 Then
                    tCount = tCount + 1
                    tName(tCount) = ws.Name
                    tFrom(tCount) = pr(1)
                    If nPR > 1 Then tTo(tCount) = pr(2) Else tTo(tCount) = pr(1)
                    tCap(tCount) = tTo(tCount) + 1
                    tDays(tCount) = 0
                    If tCap(tCount) <= nR Then
                        If InStr(1, Txt(arr(tCap(tCount), 3)), "Working Days", vbTextCompare) > 0 Then tDays(tCount) = tCap(tCount)
                    End If
                    If tDays(tCount) = 0 And tFrom(tCount) > 1 Then
                        nSmall = 0
                        For c = 1 To nC
                            If IsNum(arr(tFrom(tCount) - 1, c)) Then
                                If arr(tFrom(tCount) - 1, c) >= 1 And arr(tFrom(tCount) - 1, c) <= 400 Then nSmall = nSmall + 1
                            End If
                        Next c
                        If nSmall >= 5 Then tDays(tCount) = tFrom(tCount) - 1
                    End If
                    tStart(tCount) = Max2(tCap(tCount), tDays(tCount)) + 1
                    ' label column: C unless column B holds more labels
                    arr = ReadBlock(ws, LastRowOf(ws), 3)
                    cnt2 = 0: cnt3 = 0
                    For r = tStart(tCount) To UBound(arr, 1)
                        If Txt(arr(r, 2)) <> "" Then cnt2 = cnt2 + 1
                        If Txt(arr(r, 3)) <> "" Then cnt3 = cnt3 + 1
                    Next r
                    If cnt2 > cnt3 Then tLabCol(tCount) = 2 Else tLabCol(tCount) = 3
                End If
            End If
        End If
    Next ws
End Sub

Private Function IsMonthlyCol(arr As Variant, ByVal t As Long, ByVal c As Long) As Boolean
    Dim f As Variant, tt As Variant, cap As String
    IsMonthlyCol = False
    f = arr(tFrom(t), c): tt = arr(tTo(t), c)
    If IsPeriod(f) And IsPeriod(tt) Then
        If f = tt Then
            cap = Txt(arr(tCap(t), c))
            If InStr(cap, "Act") = 0 And InStr(cap, "Tar") = 0 And InStr(cap, "Var") = 0 And InStr(cap, "Diff") = 0 Then IsMonthlyCol = True
        End If
    End If
End Function

Private Sub BuildTrendShifts()
    Dim ws As Worksheet, out As Worksheet, arr As Variant, t As Long, nR As Long, nC As Long
    Dim mc() As Long, mp() As Long, nm As Long, c As Long, r As Long
    Dim buf() As Variant, nb As Long, j As Long
    Dim lab As String, tg As String, hasNum As Boolean, blank As Long, anyVal As Boolean
    Dim section As String, brand As String, b As String, rowBrand As String
    Dim vals() As Double, vm() As Long, nv As Long, i As Long
    Dim last As Double, prior As Double, mu As Double, sd As Double, z As Variant, stepZ As Variant
    Dim isRate As Boolean, chg As Variant, chgAvg As Variant, material As Boolean, big As Boolean, flag As Boolean
    Dim nPerVals As Long, nNeg As Long, negTxt As String
    Dim lo As String, nT As Long, a3 As Double, b12 As Double, sd12 As Double, over As String, nOver As Long, worstC As Long, worstV As Double
    Dim hdr As Variant

    ReDim buf(1 To 17, 1 To 1000)
    nb = 0
    For t = 1 To tCount
        Set ws = ThisWorkbook.Worksheets(tName(t))
        nR = LastRowOf(ws): nC = LastColOf(ws)
        arr = ReadBlock(ws, nR, nC)
        ReDim mc(1 To nC): ReDim mp(1 To nC)
        nm = 0
        For c = 1 To nC
            If IsMonthlyCol(arr, t, c) Then
                nm = nm + 1: mc(nm) = c: mp(nm) = arr(tFrom(t), c)
                If mp(nm) > gLatestMonth Then gLatestMonth = mp(nm)
            End If
        Next c
        If nm < 8 Then GoTo NextSheet
        ReDim vals(1 To nm): ReDim vm(1 To nm)
        section = "": brand = "": blank = 0
        For r = tStart(t) To nR
            anyVal = False: hasNum = False
            For c = 1 To nC
                If Not IsEmpty(arr(r, c)) Then
                    If Txt(arr(r, c)) <> "" Or IsNum(arr(r, c)) Or IsError(arr(r, c)) Then anyVal = True
                    If c > tLabCol(t) And IsNum(arr(r, c)) Then hasNum = True
                End If
                If anyVal And hasNum Then Exit For
            Next c
            If Not anyVal Then
                blank = blank + 1
                GoTo NextR
            End If
            lab = Txt(arr(r, tLabCol(t)))
            If tLabCol(t) = 3 Then tg = Txt(arr(r, 2)) Else tg = ""
            b = BrandOf(lab)
            If b = "" Then b = BrandOf(tg)
            If lab <> "" And Not hasNum Then
                If b <> "" Then
                    brand = b
                ElseIf blank >= 2 Then
                    brand = ""
                End If
                section = lab
            End If
            If lab <> "" And hasNum And tLabCol(t) = 3 And BrandOf(tg) <> "" Then
                brand = BrandOf(tg)
                section = lab
            End If
            blank = 0
            If lab = "" Or Not hasNum Then GoTo NextR
            rowBrand = BrandOf(lab)
            If rowBrand = "" Then rowBrand = brand

            ' skip rows that just repeat the period headers (e.g. 202607 in every month column)
            nPerVals = 0
            For i = 1 To nm
                If IsPeriod(arr(r, mc(i))) Then nPerVals = nPerVals + 1
            Next i
            If nPerVals > nm \ 2 Then GoTo NextR
            nv = 0
            For i = 1 To nm
                If IsNum(arr(r, mc(i))) Then
                    nv = nv + 1
                    vals(nv) = arr(r, mc(i))
                    vm(nv) = i
                End If
            Next i
            lo = LCase$(lab)
            isRate = (InStr(lo, "conversion") > 0 Or InStr(lo, "closing") > 0 Or InStr(lo, "effectiveness") > 0 Or _
                      InStr(lo, "cxd") > 0 Or InStr(lo, "ratio") > 0 Or InStr(lo, "distribution") > 0 Or InStr(lo, "%") > 0)
            ' rate > 100% (conversion / effectiveness only)
            If InStr(lo, "conversion") > 0 Or InStr(lo, "effectiveness") > 0 Then
                over = "": nOver = 0: worstV = 0: worstC = 0
                For i = 1 To nv
                    If vals(i) > 1.0001 Then
                        nOver = nOver + 1
                        If nOver <= 8 Then over = over & IIf(over = "", "", ", ") & mp(vm(i))
                        If vals(i) > worstV Then
                            worstV = vals(i)
                            worstC = mc(vm(i))
                        End If
                    End If
                Next i
                If nOver > 0 Then AddF "High", "Sanity (time series)", tName(t), ColL(worstC) & r, section & " > " & lab, rowBrand, _
                    lab & " above 100% in " & nOver & " month(s): " & over & IIf(nOver > 8, " ...", ""), worstV, 1, _
                    "Rate > 100% means numerator and denominator are from different populations."
            End If
            ' negative volumes on a volume line
            If lab = "Contacts" Or lab = "Leads" Or lab = "Gross Sales" Or lab = "Net Sales" Then
                nNeg = 0: negTxt = ""
                For i = 1 To nv
                    If vals(i) < 0 Then
                        nNeg = nNeg + 1
                        If nNeg <= 8 Then negTxt = negTxt & IIf(negTxt = "", "", ", ") & mp(vm(i))
                    End If
                Next i
                If nNeg > 0 Then
                    AddF "Medium", "Sanity (time series)", tName(t), "Row " & r, section & " > " & lab, rowBrand, _
                         "Negative " & lab & " in " & nNeg & " month(s): " & negTxt, nNeg, 0, _
                         "Negative volume - typically net cancellations booked to a closed / tiny line."
                End If
            End If
            If nv < 8 Then GoTo NextR
            last = vals(nv): prior = vals(nv - 1)
            nT = Min2(12, nv - 1)
            mu = 0
            For i = nv - nT To nv - 1: mu = mu + vals(i): Next i
            mu = mu / nT
            sd = 0
            For i = nv - nT To nv - 1: sd = sd + (vals(i) - mu) ^ 2: Next i
            sd = Sqr(sd / nT)
            If sd > 0 Then z = (last - mu) / sd Else z = Empty
            stepZ = Empty
            If nv >= 15 Then
                a3 = (vals(nv) + vals(nv - 1) + vals(nv - 2)) / 3
                b12 = 0
                For i = nv - 14 To nv - 3: b12 = b12 + vals(i): Next i
                b12 = b12 / 12
                sd12 = 0
                For i = nv - 14 To nv - 3: sd12 = sd12 + (vals(i) - b12) ^ 2: Next i
                sd12 = Sqr(sd12 / 12)
                If sd12 > 0 Then stepZ = (a3 - b12) / sd12
            End If
            If isRate Then
                chg = last - prior: chgAvg = last - mu
                material = (Abs(chgAvg) >= 0.03 Or Abs(chg) >= 0.03)
                big = True
                For i = nv - nT To nv
                    If Abs(vals(i)) > 1.5 Then big = False
                Next i
            Else
                If prior <> 0 Then chg = last / prior - 1 Else chg = Empty
                If mu <> 0 Then chgAvg = last / mu - 1 Else chgAvg = Empty
                material = False
                If Not IsEmpty(chg) Then If Abs(chg) >= 0.15 Then material = True
                If Not IsEmpty(chgAvg) Then If Abs(chgAvg) >= 0.2 Then material = True
                big = (Abs(mu) >= 20)
            End If
            flag = False
            If big And material Then
                If Not IsEmpty(z) Then If Abs(z) >= 2 Then flag = True
                If Not IsEmpty(stepZ) Then If Abs(stepZ) >= 2 Then flag = True
            End If
            nb = nb + 1
            If nb > UBound(buf, 2) Then
                ReDim Preserve buf(1 To 17, 1 To nb * 2)
            End If
            buf(1, nb) = tName(t): buf(2, nb) = r: buf(3, nb) = section
            buf(4, nb) = IIf(rowBrand = "", "All / n.a.", rowBrand): buf(5, nb) = lab
            buf(6, nb) = IIf(isRate, "rate", "volume"): buf(7, nb) = mp(vm(nv))
            buf(8, nb) = last: buf(9, nb) = prior: buf(10, nb) = mu: buf(11, nb) = sd
            buf(12, nb) = chg: buf(13, nb) = chgAvg: buf(14, nb) = z: buf(15, nb) = stepZ
            buf(16, nb) = IIf(flag, "YES", "")
            If mp(vm(nv)) = gLatestMonth And gHave(1) Then
                If CStr(mp(vm(nv))) = CStr(EV(1, 2, 2)) Then buf(17, nb) = "Latest month may be the run-rate PROJECTION in this sheet"
            End If
NextR:
        Next r
NextSheet:
    Next t

    Set out = NewSheet(OUT_TREND)
    hdr = Array("Sheet", "Row", "Section", "Brand", "Item / Metric", "Type", "Last month", "Last value", "Prior value", _
                "Trailing-12 mean", "Trailing-12 std", "Change vs prior (% vol / pp rate)", "Change vs T12 mean (% vol / pp rate)", _
                "Z-score vs T12", "Step z (last 3 vs prior 12)", "Clear shift?", "Note")
    If nb = 0 Then
        DumpTable out, 1, hdr, Empty, 0
        Exit Sub
    End If
    Dim data() As Variant
    ReDim data(1 To nb, 1 To 17)
    For i = 1 To nb
        For j = 1 To 17
            data(i, j) = buf(j, i)
        Next j
    Next i
    DumpTable out, 1, hdr, data, nb
    out.Range(out.Cells(2, 8), out.Cells(nb + 1, 11)).NumberFormat = "#,##0.00"
    out.Range(out.Cells(2, 12), out.Cells(nb + 1, 13)).NumberFormat = "0.0%"
    out.Range(out.Cells(2, 14), out.Cells(nb + 1, 15)).NumberFormat = "0.0"
    For i = 1 To nb
        If data(i, 6) = "rate" Then out.Range(out.Cells(i + 1, 8), out.Cells(i + 1, 11)).NumberFormat = "0.0%"
        If data(i, 16) = "YES" Then out.Range(out.Cells(i + 1, 1), out.Cells(i + 1, 17)).Interior.Color = RGB(255, 230, 153)
    Next i
    out.Columns("A:Q").ColumnWidth = 12
    out.Columns(3).ColumnWidth = 30
    out.Columns(5).ColumnWidth = 30
    out.Range(out.Cells(1, 1), out.Cells(1, 17)).RowHeight = 30
    SetTab out, RGB(255, 192, 0)
End Sub

'==========================================================================================
'  CROSS-SHEET CHECKS (periods, labels, working days, July 'Act' = projection)
'==========================================================================================
Private Sub CrossSheetChecks()
    Dim t As Long, ws As Worksheet, arr As Variant, nC As Long, c As Long, cap As String
    Dim f As Variant, tt As Variant, months As Long, fy As Long, fyEnd As Long, latest As Long
    Dim curP As Long, daysTxt As String, daysVals As Collection, dv As Variant, distinct As Long, found As Boolean
    Set daysVals = New Collection
    Dim r As Long, lab As String, nR As Long, i As Long
    If IsNum(EV(1, 2, 2)) Then curP = EV(1, 2, 2)
    For t = 1 To tCount
        Set ws = ThisWorkbook.Worksheets(tName(t))
        nC = LastColOf(ws): nR = LastRowOf(ws)
        arr = ReadBlock(ws, Min2(nR, Max2(tStart(t) + 40, 60)), nC)
        latest = 0
        For c = 1 To nC
            If IsMonthlyCol(arr, t, c) Then If arr(tFrom(t), c) > latest Then latest = arr(tFrom(t), c)
        Next c
        For c = 1 To nC
            cap = Txt(arr(tCap(t), c))
            f = arr(tFrom(t), c): tt = arr(tTo(t), c)
            If InStr(1, cap, "YTD", vbTextCompare) > 0 And IsPeriod(f) And IsPeriod(tt) Then
                months = (tt \ 100 - f \ 100) * 12 + (tt Mod 100 - f Mod 100) + 1
                If months > 12 Then
                    AddF "High", "Period / label", tName(t), ColL(c) & tFrom(t) & ":" & ColL(c) & tTo(t), cap, "", _
                         "Column labelled '" & cap & "' covers " & f & " to " & tt & " (" & months & " months)", f & "-" & tt, "max 12 months", _
                         "A YTD column spanning more than 12 months - it double-counts a month (e.g. includes the new month's projection). " & _
                         "YTD actual, target and variance are overstated."
                End If
                i = InStr(cap, "F")
                Do While i > 0
                    If IsNumeric(Mid$(cap, i + 1, 2)) Then Exit Do
                    i = InStr(i + 1, cap, "F")
                Loop
                If i > 0 Then
                    fy = CLng(Mid$(cap, i + 1, 2))
                    fyEnd = (2000 + fy) * 100 + 6
                    If tt < fyEnd And latest >= fyEnd And months <= 12 Then
                        AddF "Medium", "Period / label", tName(t), ColL(c) & tCap(t), cap, "", "Stale YTD column", f & "-" & tt, f & "-" & fyEnd, _
                             "Fiscal-year column stops at " & tt & " although the sheet already runs to " & latest & ". Refresh or relabel it."
                    End If
                End If
            End If
            ' current month 'Act' column equal to the Exec projection rather than the actual
            If curP > 0 And cap = CStr(curP) & " Act" Then
                For r = tStart(t) To UBound(arr, 1)
                    lab = Txt(arr(r, tLabCol(t)))
                    If lab = "Contacts" Then
                        If IsNum(arr(r, c)) And IsNum(EV(1, 6, cG)) And IsNum(EV(1, 6, cF)) Then
                            If arr(r, c) = EV(1, 6, cG) And arr(r, c) <> EV(1, 6, cF) Then
                                AddF "Medium", "Period / label", tName(t), ColL(c) & r, "Total > Contacts", "", _
                                     "'" & cap & "' holds the run-rate PROJECTION, not the actual", arr(r, c), EV(1, 6, cF), _
                                     "The current-month 'Actual' in the trend sheet equals the Exec Projecting figure (" & gName(1) & " G6), not the booked actual (F6)."
                            End If
                        End If
                        Exit For
                    End If
                Next r
            End If
        Next c
        ' working days for the current month
        If tDays(t) > 0 And curP > 0 Then
            For c = 1 To nC
                If IsPeriod(arr(tFrom(t), c)) Then
                    If arr(tFrom(t), c) = curP Then
                        If IsNum(arr(tDays(t), c)) Then
                            daysTxt = daysTxt & IIf(daysTxt = "", "", "; ") & tName(t) & ": " & arr(tDays(t), c)
                            dv = CGet(daysVals, CStr(arr(tDays(t), c)), found)
                            If Not found Then
                                daysVals.Add 1, CStr(arr(tDays(t), c))
                                distinct = distinct + 1
                            End If
                        End If
                        Exit For
                    End If
                End If
            Next c
        End If
    Next t
    If distinct > 1 Then
        AddF "Medium", "Period / label", "multiple", "working-day header rows", curP & " working days", "", _
             "Same month must use the same working-day count", daysTxt, Empty, _
             "Sheets disagree on the working days for the current month, so per-day rates differ between sheets."
    End If
End Sub

'==========================================================================================
'  METADATA (+ error-value scan)
'==========================================================================================
Private Sub BuildMetadata()
    Dim ws As Worksheet, arr As Variant, nR As Long, nC As Long, r As Long, c As Long
    Dim sBuf() As Variant, nS As Long, cBuf() As Variant, nCB As Long, rBuf() As Variant, nRB As Long
    Dim nNon As Long, nNum As Long, nTxt As Long, nErr As Long, nForm As Long
    Dim isExec As Boolean, t As Long, i As Long, firstData As Long
    Dim mcol() As Boolean, nMon As Long, pMin As Long, pMax As Long
    Dim hdrTxt As String, grp As String, v As Variant
    Dim cNon As Long, cNum As Long, cTxt As Long, cErr As Long, cMin As Double, cMax As Double, cSum As Double, cZero As Long, cNeg As Long
    Dim rn As Long, rNum As Long, rTxt As Long, rErr As Long, rMin As Double, rMax As Double, rSum As Double
    Dim rFirst As Variant, rLast As Variant, rPrior As Variant, lab As String, tg As String
    Dim section As String, brand As String, blk As Long, blank As Long, flags As String, b As String, hasNum As Boolean
    Dim brandsFound As String, errTypes As Collection, errKey As String, ek As Variant, found As Boolean
    Set errTypes = New Collection
    Dim errNames As Variant, en As Variant, errCells As String, errCnt As Long, k As Long
    Dim hdr As Variant, data() As Variant

    ReDim sBuf(1 To 18, 1 To 40): ReDim cBuf(1 To 17, 1 To 2000): ReDim rBuf(1 To 21, 1 To 20000)
    For Each ws In ThisWorkbook.Worksheets
        If IsOutputName(ws.Name) Then GoTo NextWs
        Status "metadata: " & ws.Name
        nR = LastRowOf(ws): nC = LastColOf(ws)
        arr = ReadBlock(ws, nR, nC)
        nR = UBound(arr, 1): nC = UBound(arr, 2)
        isExec = (Left$(Txt(arr(2, 3)), 9) = "From Date")
        t = 0
        For i = 1 To tCount
            If tName(i) = ws.Name Then t = i
        Next i
        ReDim mcol(1 To nC)
        nMon = 0: pMin = 0: pMax = 0
        If t > 0 Then
            For c = 1 To nC
                If IsMonthlyCol(arr, t, c) Then
                    mcol(c) = True: nMon = nMon + 1
                    If pMin = 0 Or arr(tFrom(t), c) < pMin Then pMin = arr(tFrom(t), c)
                    If arr(tFrom(t), c) > pMax Then pMax = arr(tFrom(t), c)
                End If
            Next c
            firstData = tStart(t)
        ElseIf isExec Then
            firstData = 4
        Else
            firstData = 1
        End If
        nForm = 0
        On Error Resume Next
        nForm = ws.UsedRange.SpecialCells(-4123).Count     ' xlCellTypeFormulas
        If Err.Number <> 0 Then
            nForm = 0
            Err.Clear
        End If
        On Error GoTo 0

        ' ---------------- columns -----------------------------------------------------
        nNon = 0: nNum = 0: nTxt = 0: nErr = 0
        Set errTypes = New Collection
        For c = 1 To nC
            cNon = 0: cNum = 0: cTxt = 0: cErr = 0: cSum = 0: cZero = 0: cNeg = 0: cMin = 0: cMax = 0
            For r = 1 To nR
                v = arr(r, c)
                If IsError(v) Or IsErrText(v) Then
                    nErr = nErr + 1: nNon = nNon + 1
                    errKey = ErrName(v)
                    ek = CGet(errTypes, errKey, found)
                    If Not found Then
                        errTypes.Add Array(1, ColL(c) & r), errKey
                    Else
                        errTypes.Remove errKey
                        If ek(0) < 12 Then ek(1) = ek(1) & ", " & ColL(c) & r
                        ek(0) = ek(0) + 1
                        errTypes.Add ek, errKey
                    End If
                    If r >= firstData Then
                        cNon = cNon + 1
                        cErr = cErr + 1
                    End If
                ElseIf IsNum(v) Then
                    nNum = nNum + 1: nNon = nNon + 1
                    If r >= firstData Then
                        cNon = cNon + 1: cNum = cNum + 1: cSum = cSum + v
                        If cNum = 1 Then
                            cMin = v
                            cMax = v
                        End If
                        If v < cMin Then cMin = v
                        If v > cMax Then cMax = v
                        If v = 0 Then cZero = cZero + 1
                        If v < 0 Then cNeg = cNeg + 1
                    End If
                ElseIf Txt(v) <> "" Then
                    nTxt = nTxt + 1: nNon = nNon + 1
                    If r >= firstData Then
                        cNon = cNon + 1
                        cTxt = cTxt + 1
                    End If
                End If
            Next r
            ' header text
            hdrTxt = ""
            If isExec Then
                If c = 1 Then grp = ""
                If Txt(arr(2, c)) <> "" And Left$(Txt(arr(2, c)), 4) <> "From" Then grp = Txt(arr(2, c))
                If Txt(arr(3, c)) <> "" And Left$(Txt(arr(3, c)), 2) <> "To" Then hdrTxt = IIf(grp <> "", grp & " | ", "") & Txt(arr(3, c))
            ElseIf t > 0 Then
                hdrTxt = Txt(arr(tCap(t), c))
                If hdrTxt = "" And IsPeriod(arr(tFrom(t), c)) Then
                    If arr(tFrom(t), c) = arr(tTo(t), c) Then hdrTxt = CStr(arr(tFrom(t), c)) Else hdrTxt = arr(tFrom(t), c) & "-" & arr(tTo(t), c)
                End If
            End If
            If cNon > 0 Or hdrTxt <> "" Then
                nCB = nCB + 1
                If nCB > UBound(cBuf, 2) Then
                    ReDim Preserve cBuf(1 To 17, 1 To nCB * 2)
                End If
                cBuf(1, nCB) = ws.Name: cBuf(2, nCB) = ColL(c): cBuf(3, nCB) = c: cBuf(4, nCB) = hdrTxt
                If t > 0 Then
                    If IsPeriod(arr(tFrom(t), c)) Then cBuf(5, nCB) = arr(tFrom(t), c)
                    If IsPeriod(arr(tTo(t), c)) Then cBuf(6, nCB) = arr(tTo(t), c)
                    If tDays(t) > 0 Then If IsNum(arr(tDays(t), c)) Then cBuf(7, nCB) = arr(tDays(t), c)
                End If
                cBuf(8, nCB) = cNon: cBuf(9, nCB) = (nR - firstData + 1) - cNon: cBuf(10, nCB) = cNum: cBuf(11, nCB) = cTxt
                cBuf(12, nCB) = cErr
                If cNum > 0 Then
                    cBuf(13, nCB) = cMin
                    cBuf(14, nCB) = cMax
                    cBuf(15, nCB) = cSum / cNum
                End If
                cBuf(16, nCB) = cZero: cBuf(17, nCB) = cNeg
            End If
        Next c

        ' ---------------- rows --------------------------------------------------------
        section = "": brand = "": blk = 0: blank = 0: brandsFound = ""
        For r = 1 To nR
            rn = 0: rNum = 0: rTxt = 0: rErr = 0: rSum = 0: rMin = 0: rMax = 0
            rFirst = Empty: rLast = Empty: rPrior = Empty: hasNum = False
            For c = 1 To nC
                v = arr(r, c)
                If IsError(v) Or IsErrText(v) Then
                    rn = rn + 1: rErr = rErr + 1
                ElseIf IsNum(v) Then
                    rn = rn + 1
                    If c > 4 Then hasNum = True
                    If (t > 0 And mcol(c)) Or (t = 0 And c > 4) Then
                        rNum = rNum + 1: rSum = rSum + v
                        If rNum = 1 Then
                            rMin = v
                            rMax = v
                            rFirst = v
                        End If
                        If v < rMin Then rMin = v
                        If v > rMax Then rMax = v
                        rPrior = rLast: rLast = v
                    End If
                ElseIf Txt(v) <> "" Then
                    rn = rn + 1: rTxt = rTxt + 1
                End If
            Next c
            If rn = 0 Then
                blank = blank + 1
                GoTo NextRow
            End If
            lab = ""
            For c = 4 To 1 Step -1
                If Txt(arr(r, c)) <> "" Then
                    lab = Txt(arr(r, c))
                    Exit For
                End If
            Next c
            tg = Txt(arr(r, 2))
            If isExec And Left$(lab, 9) = "From Date" Then
                blk = blk + 1
                brand = ""
            End If
            b = BrandOf(lab)
            If b = "" Then b = BrandOf(tg)
            If Right$(lab, 11) = "Brand Total" Then brand = BrandOf(Replace(lab, " - Brand Total", ""))
            If lab <> "" And Not hasNum And r >= firstData Then
                section = lab
                If b <> "" Then
                    brand = b
                ElseIf blank >= 2 And Not isExec Then
                    brand = ""
                End If
            End If
            blank = 0
            If b = "" Then b = brand
            If b <> "" And InStr(brandsFound, b) = 0 Then brandsFound = brandsFound & IIf(brandsFound = "", "", ", ") & b
            flags = ""
            If rErr > 0 Then flags = rErr & " error value(s)"
            If rNum + rErr > 0 And lab = "" And r >= firstData Then flags = flags & IIf(flags = "", "", "; ") & "numbers without a row label"
            If t > 0 And rNum > 0 Then
                If (InStr(LCase$(lab), "conversion") > 0 Or InStr(LCase$(lab), "effectiveness") > 0) And rMax > 1 Then flags = flags & IIf(flags = "", "", "; ") & "rate > 100%"
                If rMin < 0 And (lab = "Contacts" Or lab = "Leads" Or lab = "Gross Sales" Or lab = "Net Sales") Then flags = flags & IIf(flags = "", "", "; ") & "negative volume"
            End If
            nRB = nRB + 1
            If nRB > UBound(rBuf, 2) Then
                ReDim Preserve rBuf(1 To 21, 1 To nRB * 2)
            End If
            rBuf(1, nRB) = ws.Name: rBuf(2, nRB) = r: rBuf(3, nRB) = IIf(isExec, blk, Empty)
            rBuf(4, nRB) = section: rBuf(5, nRB) = b: rBuf(6, nRB) = IIf(isExec, tg, Empty): rBuf(7, nRB) = lab
            rBuf(8, nRB) = IIf(r < firstData, "Yes", "")
            rBuf(9, nRB) = rn: rBuf(10, nRB) = rNum: rBuf(11, nRB) = rTxt: rBuf(12, nRB) = rErr
            rBuf(13, nRB) = IIf(t > 0, "monthly columns", "all numeric columns (E onward)")
            If rNum > 0 Then
                rBuf(14, nRB) = rMin: rBuf(15, nRB) = rMax: rBuf(16, nRB) = rSum: rBuf(17, nRB) = rFirst
                rBuf(18, nRB) = rPrior: rBuf(19, nRB) = rLast
                If IsNum(rPrior) Then
                    If rPrior <> 0 Then rBuf(20, nRB) = rLast / rPrior - 1
                End If
            End If
            rBuf(21, nRB) = flags
NextRow:
        Next r

        ' ---------------- error-value findings ----------------------------------------
        errNames = Array("#DIV/0!", "#VALUE!", "#REF!", "#N/A", "#NAME?", "#NUM!", "#NULL!", "#ERROR")
        For Each en In errNames
            ek = CGet(errTypes, CStr(en), found)
            If found Then
                gErrCount = gErrCount + ek(0)
                AddF IIf(en = "#DIV/0!" Or en = "#N/A", "Medium", "High"), "Error values", ws.Name, ek(1) & IIf(ek(0) > 12, " ...", ""), _
                     ek(0) & " cells", "", en & " stored in the sheet", ek(0), 0, _
                     ek(0) & " cells contain " & en & ". " & IIf(en = "#REF!", "The source formula pointed at a deleted row/column/sheet - the number is lost. ", _
                     IIf(en = "#DIV/0!", "A ratio was divided by a zero / blank volume. ", "")) & _
                     "Any SUM / AVERAGE touching them breaks; wrap source formulas in IFERROR or blank them."
            End If
        Next en

        ' ---------------- sheet row ---------------------------------------------------
        nS = nS + 1
        If nS > UBound(sBuf, 2) Then
            ReDim Preserve sBuf(1 To 18, 1 To nS * 2)
        End If
        sBuf(1, nS) = ws.Name
        If isExec Then
            sBuf(2, nS) = "Exec snapshot (metric groups x Actual / Projecting / Target / Var)"
        ElseIf t > 0 Then
            sBuf(2, nS) = "Time series (rows = metrics, columns = periods)"
        Else
            sBuf(2, nS) = "Other"
        End If
        sBuf(3, nS) = "A1:" & ColL(nC) & nR: sBuf(4, nS) = nR: sBuf(5, nS) = nC
        sBuf(6, nS) = nNon: sBuf(7, nS) = nNum: sBuf(8, nS) = nTxt: sBuf(9, nS) = nErr: sBuf(10, nS) = nForm
        If t > 0 Then
            sBuf(11, nS) = "from row " & tFrom(t) & ", to row " & tTo(t) & ", caption row " & tCap(t)
            sBuf(12, nS) = IIf(tDays(t) > 0, tDays(t), Empty)
            sBuf(13, nS) = firstData
            If nMon > 0 Then sBuf(14, nS) = pMin & " - " & pMax & " (" & nMon & " monthly columns)"
        ElseIf isExec Then
            sBuf(11, nS) = "rows 2-3 (metric group / measure)"
            sBuf(13, nS) = 4
            sBuf(14, nS) = Txt(arr(2, 3)) & " / " & Txt(arr(3, 3))
        End If
        sBuf(15, nS) = brandsFound
        For i = 1 To 3
            If gName(i) = ws.Name Then sBuf(16, nS) = Choose(i, "Current month", "Prior month", "Year to date") & " (" & gPeriod(i) & ")"
        Next i
        If nForm = 0 Then sBuf(18, nS) = "No live formulas - all values hard-coded"
        If nErr > 0 Then sBuf(18, nS) = sBuf(18, nS) & IIf(sBuf(18, nS) = "", "", "; ") & nErr & " Excel error values"
NextWs:
    Next ws

    ' ---------------- write -----------------------------------------------------------
    Dim out As Worksheet, j As Long
    Set out = NewSheet(OUT_MS)
    hdr = Array("Sheet", "Layout", "Used range", "Rows", "Columns", "Non-empty cells", "Numeric cells", "Text cells", _
                "Error-value cells", "Formula cells", "Header / period rows", "Working-day row", "First data row", _
                "Periods covered", "Brands found", "Role in audit", "(reserved)", "Notes")
    ReDim data(1 To nS, 1 To 18)
    For i = 1 To nS
        For j = 1 To 18
            data(i, j) = sBuf(j, i)
        Next j
    Next i
    DumpTable out, 1, hdr, data, nS
    out.Columns("A:R").ColumnWidth = 18
    out.Columns(17).Hidden = True
    out.Range(out.Cells(2, 1), out.Cells(nS + 1, 18)).WrapText = True
    SetTab out, RGB(127, 127, 127)

    Set out = NewSheet(OUT_MC)
    hdr = Array("Sheet", "Column", "Col #", "Header (inferred)", "Period from", "Period to", "Working days", "Non-blank", _
                "Blank", "Numeric", "Text", "Error values", "Min", "Max", "Mean", "Zeros", "Negatives")
    ReDim data(1 To nCB, 1 To 17)
    For i = 1 To nCB
        For j = 1 To 17
            data(i, j) = cBuf(j, i)
        Next j
    Next i
    DumpTable out, 1, hdr, data, nCB
    out.Columns("A:Q").ColumnWidth = 12
    out.Columns(1).ColumnWidth = 20
    out.Columns(4).ColumnWidth = 36
    SetTab out, RGB(127, 127, 127)

    Set out = NewSheet(OUT_MR)
    hdr = Array("Sheet", "Row", "Block #", "Section", "Brand (inferred)", "Tag (col B)", "Item / Metric label", "Is header row", _
                "Non-blank cells", "Numeric cells (stats basis)", "Text cells", "Error cells", "Stats basis", "Min", "Max", "Sum", _
                "First value", "Prior value", "Last value", "Last vs prior %", "Flags")
    ReDim data(1 To nRB, 1 To 21)
    For i = 1 To nRB
        For j = 1 To 21
            data(i, j) = rBuf(j, i)
        Next j
    Next i
    DumpTable out, 1, hdr, data, nRB
    out.Columns("A:U").ColumnWidth = 12
    out.Columns(1).ColumnWidth = 20
    out.Columns(4).ColumnWidth = 30
    out.Columns(7).ColumnWidth = 34
    out.Columns(21).ColumnWidth = 30
    out.Range(out.Cells(2, 20), out.Cells(nRB + 1, 20)).NumberFormat = "0.0%"
    SetTab out, RGB(127, 127, 127)
End Sub

'==========================================================================================
'  BRAND x ITEM  (current month vs prior month vs YTD vs target)
'==========================================================================================
Private Function BrandKey(ByVal s As Long, ByVal r As Long) As String
    BrandKey = ""
    If gLab(s, r) = "" Or gSB(s, r) <> "" Then Exit Function
    If gBB(s, r) <> "" Then
        If Right$(gLab(s, r), 11) = "Brand Total" Then BrandKey = gBB(s, r) & "|Brand Total" Else BrandKey = gBB(s, r) & "|" & gLab(s, r)
    Else
        BrandKey = "All brands|" & gLab(s, r)
    End If
End Function

Private Sub BuildBrandItem()
    Dim idx1 As Collection, idx2 As Collection, idx3 As Collection, s As Long, r As Long, key As String, found As Boolean
    Dim mN As Variant, mA As Variant, mP As Variant, mT As Variant, mK As Variant, mD As Variant
    Dim buf() As Variant, nb As Long, m As Long, ru As Variant, ry As Variant, haveU As Boolean, haveY As Boolean
    Dim ja As Variant, jp As Variant, jt As Variant, ua As Variant, ut As Variant, ya As Variant, yt As Variant
    Dim baseVal As Variant, jvar As Variant, uvar As Variant, yvar As Variant, mom As Variant, vsAvg As Variant
    Dim size As Double, material As Boolean, flags As String, score As Double, isVol As Boolean
    Dim parts As Variant, i As Long, j As Long, hdr As Variant, data() As Variant, out As Worksheet
    mN = Array("Contacts", "Conversion %", "Leads", "Gross Closing %", "Gross Sales", "Gross Effectiveness %", "NB Net CXD %", "NB Net Sales")
    mA = Array(cF, cL, cS, cY, cAF, cAL, cAP, cAT)
    mP = Array(cG, 0, cT, 0, cAG, 0, 0, cAU)
    mT = Array(cH, cM, cU, cZ, cAH, cAM, cAQ, cAV)
    mK = Array(True, False, True, False, True, False, False, True)
    mD = Array(cF, cF, cS, cS, cAF, cF, cAF, cAT)
    Set idx1 = New Collection
    Set idx2 = New Collection
    Set idx3 = New Collection
    For s = 1 To 3
        If gHave(s) Then
            For r = 4 To gRows(s)
                key = BrandKey(s, r)
                If key <> "" Then
                    If s = 1 Then CPut idx1, key, r
                    If s = 2 Then CPut idx2, key, r
                    If s = 3 Then CPut idx3, key, r
                End If
            Next r
        End If
    Next s
    ReDim buf(1 To 22, 1 To 4000)
    For r = 4 To gRows(1)
        key = BrandKey(1, r)
        If key = "" Then GoTo NextR
        found = False
        ru = CGet(idx1, key, found)
        If ru <> r Then GoTo NextR                  ' only the first occurrence of each brand|item
        haveU = False: haveY = False
        If gHave(2) Then ru = CGet(idx2, key, haveU)
        If gHave(3) Then ry = CGet(idx3, key, haveY)
        parts = Split(key, "|")
        For m = 0 To 7
            isVol = mK(m)
            ja = EV(1, r, mA(m))
            If mP(m) > 0 Then jp = EV(1, r, mP(m)) Else jp = Empty
            jt = EV(1, r, mT(m))
            ua = Empty: ut = Empty: ya = Empty: yt = Empty
            If haveU Then
                ua = EV(2, CLng(ru), mA(m))
                ut = EV(2, CLng(ru), mT(m))
            End If
            If haveY Then
                ya = EV(3, CLng(ry), mA(m))
                yt = EV(3, CLng(ry), mT(m))
            End If
            If Not (IsNum(ja) Or IsNum(ua) Or IsNum(ya)) Then GoTo NextM
            If isVol Then baseVal = jp Else baseVal = ja
            jvar = Empty: uvar = Empty: yvar = Empty: mom = Empty: vsAvg = Empty
            If IsNum(baseVal) And IsNum(jt) Then If jt <> 0 Then jvar = baseVal / jt - 1
            If IsNum(ua) And IsNum(ut) Then If ut <> 0 Then uvar = ua / ut - 1
            If IsNum(ya) And IsNum(yt) Then If yt <> 0 Then yvar = ya / yt - 1
            If isVol Then
                If IsNum(jp) And IsNum(ua) Then If ua <> 0 Then mom = jp / ua - 1
                If IsNum(jp) And IsNum(ya) Then If ya <> 0 Then vsAvg = jp / (ya / 12) - 1
                size = 0
                If IsNum(ja) Then size = ja
                If IsNum(ua) Then If ua > size Then size = ua
            Else
                If IsNum(ja) And IsNum(ua) Then mom = ja - ua
                If IsNum(ja) And IsNum(ya) Then vsAvg = ja - ya
                size = 0
                If IsNum(EV(1, r, mD(m))) Then size = EV(1, r, mD(m))
            End If
            material = (size >= 500)
            flags = "": score = 0
            If material Then
                If isVol Then
                    If Not IsEmpty(mom) Then
                        If Abs(mom) >= 0.2 Then
                            flags = flags & "; Cur (proj) vs prior " & Format(mom, "+0%;-0%")
                            score = score + Abs(mom) / 0.2
                        End If
                    End If
                    If Not IsEmpty(vsAvg) Then
                        If Abs(vsAvg) >= 0.25 Then
                            flags = flags & "; Cur vs YTD monthly avg " & Format(vsAvg, "+0%;-0%")
                            score = score + Abs(vsAvg) / 0.25
                        End If
                    End If
                Else
                    If Not IsEmpty(mom) Then
                        If Abs(mom) >= 0.05 Then
                            flags = flags & "; Cur vs prior " & Format(mom * 100, "+0.0;-0.0") & "pp"
                            score = score + Abs(mom) / 0.05
                        End If
                    End If
                    If Not IsEmpty(vsAvg) Then
                        If Abs(vsAvg) >= 0.05 Then
                            flags = flags & "; Cur vs YTD " & Format(vsAvg * 100, "+0.0;-0.0") & "pp"
                            score = score + Abs(vsAvg) / 0.05
                        End If
                    End If
                End If
                If Not IsEmpty(jvar) And Not IsEmpty(yvar) Then
                    If Abs(jvar - yvar) >= 0.15 Then
                        flags = flags & "; Target var swing " & Format(yvar, "+0%;-0%") & " (YTD) -> " & Format(jvar, "+0%;-0%") & " (cur)"
                        score = score + Abs(jvar - yvar) / 0.15
                    End If
                End If
                If Not IsEmpty(jvar) And Not IsEmpty(uvar) Then
                    If ((jvar > 0) <> (uvar > 0)) And Abs(jvar - uvar) >= 0.1 Then
                        flags = flags & "; crossed the target line vs prior month"
                        score = score + 1
                    End If
                End If
            End If
            If Not isVol And IsNum(ja) And mN(m) <> "NB Net CXD %" Then If ja > 1 Then flags = flags & "; rate > 100% (data issue)"
            If flags <> "" Then flags = Mid$(flags, 3)
            nb = nb + 1
            If nb > UBound(buf, 2) Then
                ReDim Preserve buf(1 To 22, 1 To nb * 2)
            End If
            buf(1, nb) = parts(0): buf(2, nb) = parts(1): buf(3, nb) = mN(m): buf(4, nb) = IIf(isVol, "vol", "rate")
            buf(5, nb) = ja: buf(6, nb) = jp: buf(7, nb) = jt: buf(8, nb) = jvar
            buf(9, nb) = ua: buf(10, nb) = ut: buf(11, nb) = uvar
            buf(12, nb) = ya: buf(13, nb) = yt: buf(14, nb) = yvar
            buf(15, nb) = mom: buf(16, nb) = vsAvg: buf(17, nb) = size
            buf(18, nb) = IIf(flags <> "" And score > 0, "YES", "")
            If score > 0 Then buf(19, nb) = Round(score, 2) Else buf(19, nb) = Empty
            buf(20, nb) = flags: buf(21, nb) = r
NextM:
        Next m
NextR:
    Next r
    Set out = NewSheet(OUT_BRAND)
    hdr = Array("Brand", "Item / Channel", "Metric", "Type", _
                "Cur Actual", "Cur Projecting", "Cur Target", "Cur Var vs Tar", _
                "Prior Actual", "Prior Target", "Prior Var vs Tar", "YTD Actual", "YTD Target", "YTD Var vs Tar", _
                "Cur vs Prior (% vol / pp rate)", "Cur vs YTD avg (% vol / pp rate)", "Base size", "Clear shift?", "Shift score", _
                "Shift detail", "Exec row (current)")
    ReDim data(1 To Max2(nb, 1), 1 To 21)
    For i = 1 To nb
        For j = 1 To 21
            data(i, j) = buf(j, i)
        Next j
    Next i
    DumpTable out, 1, hdr, data, nb
    out.Cells(1, 5).Value = "Cur Actual (" & gPeriod(1) & ")"
    If gHave(2) Then out.Cells(1, 9).Value = "Prior Actual (" & gPeriod(2) & ")"
    If gHave(3) Then out.Cells(1, 12).Value = "YTD Actual (" & gPeriod(3) & ")"
    out.Columns("A:U").ColumnWidth = 12
    out.Columns(2).ColumnWidth = 30
    out.Columns(3).ColumnWidth = 20
    out.Columns(20).ColumnWidth = 60
    out.Range(out.Cells(2, 5), out.Cells(nb + 1, 7)).NumberFormat = "#,##0"
    out.Range(out.Cells(2, 9), out.Cells(nb + 1, 10)).NumberFormat = "#,##0"
    out.Range(out.Cells(2, 12), out.Cells(nb + 1, 13)).NumberFormat = "#,##0"
    out.Range(out.Cells(2, 17), out.Cells(nb + 1, 17)).NumberFormat = "#,##0"
    out.Range(out.Cells(2, 8), out.Cells(nb + 1, 8)).NumberFormat = "0.0%"
    out.Range(out.Cells(2, 11), out.Cells(nb + 1, 11)).NumberFormat = "0.0%"
    out.Range(out.Cells(2, 14), out.Cells(nb + 1, 16)).NumberFormat = "0.0%"
    For i = 1 To nb
        If buf(4, i) = "rate" Then
            out.Range(out.Cells(i + 1, 5), out.Cells(i + 1, 7)).NumberFormat = "0.0%"
            out.Range(out.Cells(i + 1, 9), out.Cells(i + 1, 10)).NumberFormat = "0.0%"
            out.Range(out.Cells(i + 1, 12), out.Cells(i + 1, 13)).NumberFormat = "0.0%"
        End If
        If buf(18, i) = "YES" Then out.Range(out.Cells(i + 1, 1), out.Cells(i + 1, 21)).Interior.Color = RGB(255, 230, 153)
    Next i
    SetTab out, RGB(255, 192, 0)
End Sub

'==========================================================================================
'  WEIGHTED CONVERSION & CLOSING  (live formulas on the current Exec sheet)
'==========================================================================================
Private Sub BuildWeighted()
    Dim ws As Worksheet, r As Long, s As Long, detail As Long, top As Long, b As Long, k As Long, i As Long
    Dim brands As Variant, defs As Variant, bb As Long, dTot As Long, dRep As Long, tmp1 As Long, tmp2 As Long
    Set ws = NewSheet(OUT_WEIGHTED)
    NoGrid
    s = 1
    ws.Columns(1).ColumnWidth = 34: ws.Columns(2).ColumnWidth = 9
    For i = 3 To 10: ws.Columns(i).ColumnWidth = 14: Next i
    ws.Columns(11).ColumnWidth = 50
    ws.Cells(1, 1).Value = "How Conversion %, Gross Closing %, their Targets and Weighted (W) Targets are built"
    ws.Cells(1, 1).Font.Bold = True: ws.Cells(1, 1).Font.Size = 14: ws.Cells(1, 1).Font.Color = RGB(31, 78, 120)
    ws.Cells(2, 1).Value = "Source: '" & gName(1) & "' (period " & gPeriod(1) & "). Green = live link, black = formula. " & _
                           "Paste new data into the Exec sheet and these formulas update; re-run the macro to refresh row positions."
    ws.Cells(2, 1).Font.Italic = True

    r = 4
    defs = Array( _
        Array("Measure", "Formula (per line i)", "Notes"), _
        Array("Conversion %", "Eligible Leads / Eligible Contacts", "Only lines where a contact precedes the lead (Traditional, Competition, OQ-PNV, Direct Marketing, Affiliates, Hippo-Branded, DM Outbound). OQ-PV / End-to-End / Hippo-Unbranded / Broker have contacts = leads and are excluded."), _
        Array("Conversion Target %", "Target eligible Leads / Target Contacts", "Plan rate at the PLAN channel mix."), _
        Array("W Target Conversion %", "SUM(Actual eligible Contacts_i x Target Conv_i) / SUM(Actual eligible Contacts_i)", "Plan rate at the ACTUAL channel mix. Built bottom-up; a line with no target enters at 0% (see example D)."), _
        Array("Gross Closing %", "Gross Sales / Leads", "All leads."), _
        Array("Closing Target %", "Target Gross Sales / Target Leads", "Plan closing at the plan lead mix."), _
        Array("W Target Closing %", "SUM(Actual Leads_i x Target Closing_i) / SUM(Actual Leads_i)", "Plan closing at the actual lead mix."), _
        Array("Tar Var (rates)", "Actual / Target - 1", "Relative variance, not percentage points."), _
        Array("W Tar Var", "Actual / W Target - 1", "Performance vs plan rates once the mix is neutralised."), _
        Array("Mix effect", "W Target - Target", "Change in expected rate caused purely by the channel mix moving vs plan."), _
        Array("Rate effect", "Actual - W Target", "Genuine conversion / closing performance. Mix + Rate = Actual - Target."), _
        Array("Gross Effectiveness %", "Eligible Gross Sales / Eligible Contacts (= Conversion x Closing on eligible lines)", ""), _
        Array("NB Net CXD %", "1 - Net Sales / Gross Sales", ""), _
        Array("Projecting", "Actual x Total working days / Elapsed working days", "Factor used on this sheet: " & Format(ProjFactor(1), "0.00000")), _
        Array("Tar Dif / Tar Var (volumes)", "Projecting - Target ;  (Projecting - Target) / Target", ""))
    For i = 0 To UBound(defs)
        ws.Cells(r, 1).Value = defs(i)(0)
        ws.Cells(r, 3).Value = defs(i)(1)
        ws.Cells(r, 8).Value = defs(i)(2)
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Merge
        ws.Range(ws.Cells(r, 3), ws.Cells(r, 7)).Merge
        ws.Range(ws.Cells(r, 8), ws.Cells(r, 11)).Merge
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 11)).WrapText = True
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 11)).VerticalAlignment = -4160   ' xlTop
        If i = 0 Then
            StyleHeader ws, r, 11
        Else
            ws.Rows(r).RowHeight = IIf(Len(defs(i)(2)) > 110, 42, 28)
        End If
        r = r + 1
    Next i
    r = r + 1

    ' locate blocks on the current Exec sheet
    detail = -1: top = -1
    For i = 4 To gRows(s)
        If detail < 0 And gLab(s, i) = "Generic - Traditional" And gBB(s, i) = "" Then detail = gBlk(s, i)
        If top < 0 And gLab(s, i) = "Non-Direct" And gBB(s, i) = "" And gSB(s, i) = "" Then top = gBlk(s, i)
    Next i
    If detail < 0 Then detail = top

    Example ws, r, s, detail, "A. Direct - Conversion W Target (contacts-weighted)", _
        "Conversion-eligible children of Direct. OQ uses only OQ-PNV and Hippo only Hippo-Branded (the only lines with a conversion step). " & _
        "Competition has no target, so it enters at 0% and dilutes the parent W.", _
        cF, cL, cM, cP, Array("Traditional Combined", "Competitions Combined", "OQ - PNV", "Direct Marketing", "Affiliates", "Hippo - Branded", "DM Multibrand Outbound"), _
        "Direct", "Contacts", "Conv %", tmp1, tmp2
    Example ws, r, s, detail, "B. Direct - Gross Closing W Target (leads-weighted)", _
        "Every lead counts for closing. DM Multibrand Outbound already contains Referral.", _
        cS, cY, cZ, cAC, Array("Traditional Combined", "Competitions Combined", "OQ", "Direct Marketing", "Affiliates", "Hippo", "Better Compare", "Cover Compare", "Direct Other Clean Up", "DM Multibrand Outbound"), _
        "Direct", "Leads", "Closing %", tmp1, tmp2
    Example ws, r, s, detail, "C. OQ - Gross Closing W Target, one level down", _
        "Shows the recursion: OQ's W is the lead-weighted average of its three sub-lines. A small residual comes from a lower grain (product) not in the file.", _
        cS, cY, cZ, cAC, Array("OQ - PNV", "OQ - PV", "OQ - End to End"), "OQ", "Leads", "Closing %", tmp1, tmp2
    Example ws, r, s, top, "D. Where it breaks - Non-Direct Conversion W Target", _
        "Children with W = 0 although a plan target exists (e.g. Dealer Channel), and Affinities' own W near 0 (OEM / 1Life), make the parent W collapse " & _
        "and W Tar Var explode. That is a calculation error, not performance.", _
        cF, cL, cM, cP, Array("Affinities", "Bank", "Dealer Channel", "Broker", "Property Partners", "Other Cleanup"), _
        "Non-Direct", "Contacts", "Conv %", dTot, dRep
    ws.Cells(r - 1, 1).Formula = "=""Corrected view: using each child's plan target instead of its broken W gives ""&TEXT(F" & dTot & ",""0.0%"")" & _
        "&"" vs the reported W of ""&TEXT(G" & dRep & ",""0.00%"")&"", i.e. W Tar Var of about ""&TEXT(IF(F" & dTot & "=0,0,D" & dRep & "/F" & dTot & "-1),""+0%;-0%"")" & _
        "&"" instead of ""&TEXT(IF(G" & dRep & "=0,0,D" & dRep & "/G" & dRep & "-1),""+0%;-0%"")&""."""
    ws.Cells(r - 1, 1).Font.Bold = True: ws.Cells(r - 1, 1).Font.Color = RGB(192, 0, 0)
    r = r + 1

    brands = BrandList()
    For k = 0 To 4
        bb = -1
        For i = 4 To gRows(s)
            If gBB(s, i) = brands(k) And gLab(s, i) = "Direct (Excl Outbound)" Then
                bb = gBlk(s, i)
                Exit For
            End If
        Next i
        If bb >= 0 Then
            Example ws, r, s, bb, Chr$(69 + k) & ". " & brands(k) & " - Direct (Excl Outbound) Conversion W Target", _
                "Same mechanics inside the " & brands(k) & " brand page.", _
                cF, cL, cM, cP, Array("Traditional Combined", "Competitions Combined", "OQ - PNV", "Direct Marketing", "Affiliates", "Hippo - Branded"), _
                "Direct (Excl Outbound)", "Contacts", "Conv %", tmp1, tmp2
        End If
    Next k

    Calculator ws, r, s
    SetTab ws, RGB(0, 176, 80)
    Freeze ws, 4
End Sub

Private Sub Example(ByVal ws As Worksheet, ByRef r As Long, ByVal s As Long, ByVal blk As Long, ByVal title As String, ByVal note As String, _
                    ByVal volC As Long, ByVal rateC As Long, ByVal tgtC As Long, ByVal wC As Long, kids As Variant, ByVal parent As String, _
                    ByVal volName As String, ByVal kind As String, ByRef totRow As Long, ByRef repRow As Long)
    Dim src As String, i As Long, kr As Long, prow As Long, first As Long, last As Long, tot As Long, rep As Long, j As Long
    Dim hdr As Variant, labels As Variant, fm As Variant
    src = QS(gName(s))
    Banner ws, r, title, 11
    r = r + 1
    ws.Cells(r, 1).Value = note
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 11)).Merge
    ws.Cells(r, 1).WrapText = True: ws.Cells(r, 1).Font.Italic = True
    ws.Rows(r).RowHeight = 30
    r = r + 1
    prow = FindInBlock(s, blk, parent, 0)
    hdr = Array("Line (child)", "Src row", "Actual " & volName, "Actual " & kind, "Implied numerator", "Target " & kind, _
                "W Target " & kind, volName & " x W Target", volName & " x Target", "Mix share", "Comment")
    For j = 0 To 10
        ws.Cells(r, j + 1).Value = hdr(j)
    Next j
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 11))
        .Font.Bold = True: .Interior.Color = RGB(221, 235, 247): .WrapText = True
    End With
    first = r + 1
    last = r + UBound(kids) + 1
    For i = 0 To UBound(kids)
        r = r + 1
        If prow > 0 Then kr = FindInBlock(s, blk, CStr(kids(i)), 0) Else kr = 0
        ws.Cells(r, 1).Value = kids(i)
        If kr > 0 Then
            ws.Cells(r, 2).Value = kr
            ws.Cells(r, 3).Formula = "=N(" & src & ColL(volC) & kr & ")"
            ws.Cells(r, 4).Formula = "=N(" & src & ColL(rateC) & kr & ")"
            ws.Cells(r, 6).Formula = "=N(" & src & ColL(tgtC) & kr & ")"
            ws.Cells(r, 7).Formula = "=N(" & src & ColL(wC) & kr & ")"
        Else
            ws.Cells(r, 2).Value = "n/a"
            ws.Cells(r, 3).Value = 0: ws.Cells(r, 4).Value = 0: ws.Cells(r, 6).Value = 0: ws.Cells(r, 7).Value = 0
            ws.Cells(r, 11).Value = "Line not found in this block"
        End If
        ws.Cells(r, 5).Formula = "=C" & r & "*D" & r
        ws.Cells(r, 8).Formula = "=C" & r & "*G" & r
        ws.Cells(r, 9).Formula = "=C" & r & "*F" & r
        ws.Cells(r, 10).Formula = "=IF(SUM($C$" & first & ":$C$" & last & ")=0,0,C" & r & "/SUM($C$" & first & ":$C$" & last & "))"
        ws.Range(ws.Cells(r, 3), ws.Cells(r, 4)).Font.Color = RGB(0, 128, 0)
        ws.Range(ws.Cells(r, 6), ws.Cells(r, 7)).Font.Color = RGB(0, 128, 0)
    Next i
    r = r + 1: tot = r
    ws.Cells(r, 1).Value = "Total / re-derived"
    ws.Cells(r, 3).Formula = "=SUM(C" & first & ":C" & last & ")"
    ws.Cells(r, 4).Formula = "=IF(C" & r & "=0,0,E" & r & "/C" & r & ")"
    ws.Cells(r, 5).Formula = "=SUM(E" & first & ":E" & last & ")"
    ws.Cells(r, 6).Formula = "=IF(C" & r & "=0,0,I" & r & "/C" & r & ")"
    ws.Cells(r, 7).Formula = "=IF(C" & r & "=0,0,H" & r & "/C" & r & ")"
    ws.Cells(r, 8).Formula = "=SUM(H" & first & ":H" & last & ")"
    ws.Cells(r, 9).Formula = "=SUM(I" & first & ":I" & last & ")"
    ws.Cells(r, 10).Formula = "=SUM(J" & first & ":J" & last & ")"
    ws.Cells(r, 11).Value = "D = SUM(E)/SUM(C) re-derived actual; F = children's plan targets at actual mix; G = SUMPRODUCT(C,W)/SUM(C) = re-derived W Target"
    ws.Cells(r, 11).WrapText = True
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 11)).Font.Bold = True
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 11)).Interior.Color = RGB(242, 242, 242)
    ws.Rows(r).RowHeight = 40
    r = r + 1: rep = r
    ws.Cells(r, 1).Value = "Reported on Exec sheet (row " & IIf(prow > 0, prow, "n/a") & ")"
    If prow > 0 Then
        ws.Cells(r, 2).Value = prow
        ws.Cells(r, 3).Formula = "=N(" & src & ColL(volC) & prow & ")"
        ws.Cells(r, 4).Formula = "=N(" & src & ColL(rateC) & prow & ")"
        ws.Cells(r, 6).Formula = "=N(" & src & ColL(tgtC) & prow & ")"
        ws.Cells(r, 7).Formula = "=N(" & src & ColL(wC) & prow & ")"
    End If
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Font.Bold = True
    ws.Range(ws.Cells(r, 3), ws.Cells(r, 7)).Font.Color = RGB(0, 128, 0)
    r = r + 1
    ws.Cells(r, 1).Value = "Difference (re-derived - reported)"
    ws.Cells(r, 4).Formula = "=D" & tot & "-D" & rep
    ws.Cells(r, 7).Formula = "=G" & tot & "-G" & rep
    ws.Cells(r, 8).Formula = "=IF(ABS(G" & tot & "-G" & rep & ")<0.0005,""W Target reconciles"",""W Target does NOT reconcile"")"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 8)).Font.Bold = True
    ws.Range(ws.Cells(r, 4), ws.Cells(r, 7)).NumberFormat = "+0.00%;-0.00%;0.00%"
    r = r + 1
    ws.Cells(r, 1).Value = "Variance split (reported figures)": ws.Cells(r, 1).Font.Bold = True
    labels = Array("Actual - Target (total gap, pp)", "Mix effect = W Target - Target (pp)", "Rate effect = Actual - W Target (pp)", _
                   "Check: Mix + Rate - Total (should be 0)", "Tar Var = Actual / Target - 1", "W Tar Var = Actual / W Target - 1")
    fm = Array("=D" & rep & "-F" & rep, "=G" & rep & "-F" & rep, "=D" & rep & "-G" & rep, _
               "=(G" & rep & "-F" & rep & ")+(D" & rep & "-G" & rep & ")-(D" & rep & "-F" & rep & ")", _
               "=IF(F" & rep & "=0,0,D" & rep & "/F" & rep & "-1)", "=IF(G" & rep & "=0,0,D" & rep & "/G" & rep & "-1)")
    For i = 0 To 5
        r = r + 1
        ws.Cells(r, 1).Value = labels(i)
        ws.Cells(r, 4).Formula = fm(i)
        ws.Cells(r, 4).NumberFormat = "+0.00%;-0.00%;0.00%"
        ws.Cells(r, 4).Font.Bold = True
    Next i
    ' formats
    ws.Range(ws.Cells(first, 3), ws.Cells(rep, 3)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(first, 5), ws.Cells(rep, 5)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(first, 8), ws.Cells(rep, 9)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(first, 4), ws.Cells(rep, 4)).NumberFormat = "0.00%"
    ws.Range(ws.Cells(first, 6), ws.Cells(rep, 7)).NumberFormat = "0.00%"
    ws.Range(ws.Cells(first, 10), ws.Cells(rep, 10)).NumberFormat = "0.0%"
    ws.Range(ws.Cells(first - 1, 1), ws.Cells(rep, 11)).Borders.LineStyle = 1
    totRow = tot: repRow = rep
    r = r + 2
End Sub

Private Sub Calculator(ByVal ws As Worksheet, ByRef r As Long, ByVal s As Long)
    Dim i As Long, first As Long, last As Long, baseVal As Long, n As Long, cRows(1 To 6) As Long, prow As Long, lbl As Variant
    Dim hdr As Variant, j As Long, outs As Variant, bk As Long
    Banner ws, r, "Calculator - type your own lines (yellow cells) to see Target vs W Target vs Actual", 11
    r = r + 1
    ' prefill with the first brand split under Direct (Excl Outbound)
    n = 0: prow = 0
    For i = 4 To gRows(s)
        If gSB(s, i) <> "" And gSP(s, i) = "Direct (Excl Outbound)" And gSB(s, i) <> "TIH Multibrand" Then
            If prow = 0 Then prow = i
            If gBlk(s, i) = gBlk(s, prow) And n < 5 Then
                n = n + 1
                cRows(n) = i
            End If
        End If
    Next i
    ws.Cells(r, 1).Value = "Pre-filled with Direct (Excl Outbound) by brand from '" & gName(1) & "' (values copied at run time). " & _
                           "Overwrite volumes, rates or targets to test scenarios. Rows with blank volume are ignored."
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 11)).Merge
    ws.Cells(r, 1).Font.Italic = True
    r = r + 1
    hdr = Array("Line", "", "Actual volume", "Actual rate", "Implied numerator", "Target rate", "Target volume", _
                "Volume x Target rate", "Target numerator", "Mix share actual", "Mix share plan")
    For j = 0 To 10: ws.Cells(r, j + 1).Value = hdr(j): Next j
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 11))
        .Font.Bold = True: .Interior.Color = RGB(221, 235, 247): .WrapText = True
    End With
    first = r + 1: last = r + 6
    For i = 1 To 6
        r = r + 1
        If i <= n Then
            ws.Cells(r, 1).Value = gLab(s, cRows(i))
            ws.Cells(r, 3).Value = EV(s, cRows(i), cF)
            ws.Cells(r, 4).Value = EV(s, cRows(i), cL)
            ws.Cells(r, 6).Value = EV(s, cRows(i), cM)
            ws.Cells(r, 7).Value = EV(s, cRows(i), cH)
        Else
            ws.Cells(r, 1).Value = "(spare line)"
        End If
        With ws.Range(ws.Cells(r, 3), ws.Cells(r, 4))
            .Interior.Color = RGB(255, 255, 0): .Font.Color = RGB(0, 0, 255)
        End With
        With ws.Range(ws.Cells(r, 6), ws.Cells(r, 7))
            .Interior.Color = RGB(255, 255, 0): .Font.Color = RGB(0, 0, 255)
        End With
        ws.Cells(r, 5).Formula = "=N(C" & r & ")*N(D" & r & ")"
        ws.Cells(r, 8).Formula = "=N(C" & r & ")*N(F" & r & ")"
        ws.Cells(r, 9).Formula = "=N(G" & r & ")*N(F" & r & ")"
        ws.Cells(r, 10).Formula = "=IF(SUM($C$" & first & ":$C$" & last & ")=0,0,N(C" & r & ")/SUM($C$" & first & ":$C$" & last & "))"
        ws.Cells(r, 11).Formula = "=IF(SUM($G$" & first & ":$G$" & last & ")=0,0,N(G" & r & ")/SUM($G$" & first & ":$G$" & last & "))"
    Next i
    ws.Range(ws.Cells(first, 3), ws.Cells(last, 3)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(first, 5), ws.Cells(last, 5)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(first, 7), ws.Cells(last, 9)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(first, 4), ws.Cells(last, 4)).NumberFormat = "0.00%"
    ws.Range(ws.Cells(first, 6), ws.Cells(last, 6)).NumberFormat = "0.00%"
    ws.Range(ws.Cells(first, 10), ws.Cells(last, 11)).NumberFormat = "0.0%"
    ws.Range(ws.Cells(first - 1, 1), ws.Cells(last, 11)).Borders.LineStyle = 1
    r = r + 1: baseVal = r
    outs = Array( _
        Array("Actual rate (volume-weighted)", "=IF(SUM(C" & first & ":C" & last & ")=0,0,SUM(E" & first & ":E" & last & ")/SUM(C" & first & ":C" & last & "))"), _
        Array("Target rate (plan mix) = SUM(target numerators) / SUM(target volumes)", "=IF(SUM(G" & first & ":G" & last & ")=0,0,SUM(I" & first & ":I" & last & ")/SUM(G" & first & ":G" & last & "))"), _
        Array("W Target rate (actual mix) = SUMPRODUCT(volume, target rate) / SUM(volume)", "=IF(SUM(C" & first & ":C" & last & ")=0,0,SUM(H" & first & ":H" & last & ")/SUM(C" & first & ":C" & last & "))"), _
        Array("Mix effect = W Target - Target", "=D" & baseVal + 2 & "-D" & baseVal + 1), _
        Array("Rate effect = Actual - W Target", "=D" & baseVal & "-D" & baseVal + 2), _
        Array("Tar Var = Actual / Target - 1", "=IF(D" & baseVal + 1 & "=0,0,D" & baseVal & "/D" & baseVal + 1 & "-1)"), _
        Array("W Tar Var = Actual / W Target - 1", "=IF(D" & baseVal + 2 & "=0,0,D" & baseVal & "/D" & baseVal + 2 & "-1)"))
    For i = 0 To UBound(outs)
        ws.Cells(r, 1).Value = outs(i)(0)
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Merge
        ws.Cells(r, 1).Font.Bold = True
        ws.Cells(r, 4).Formula = outs(i)(1)
        ws.Cells(r, 4).Font.Bold = True
        ws.Cells(r, 4).NumberFormat = IIf(i < 3, "0.00%", "+0.00%;-0.00%;0.00%")
        r = r + 1
    Next i
    ' compare with the reported W of the parent line
    prow = 0
    If n > 0 Then
        bk = gBlk(s, cRows(1))
        For i = gBS(s, bk) To cRows(1) - 1
            If gLab(s, i) = gSP(s, cRows(1)) And gSB(s, i) = "" Then prow = i
        Next i
    End If
    If prow > 0 And n > 0 Then
        ws.Cells(r, 1).Value = "Reported W Target on Exec sheet (P" & prow & ")"
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Merge
        ws.Cells(r, 4).Formula = "=N(" & QS(gName(s)) & "P" & prow & ")"
        ws.Cells(r, 4).NumberFormat = "0.00%": ws.Cells(r, 4).Font.Color = RGB(0, 128, 0): ws.Cells(r, 4).Font.Bold = True
        ws.Cells(r, 5).Formula = "=""Weighting at BRAND level gives ""&TEXT(D" & baseVal + 2 & ",""0.0%"")&"" vs the reported ""&TEXT(D" & r & ",""0.0%"")&" & _
                                 """. The Exec sheet weights at sub-channel level inside each brand - always build W from the lowest grain upward."""
        ws.Range(ws.Cells(r, 5), ws.Cells(r, 11)).Merge
        ws.Cells(r, 5).WrapText = True: ws.Cells(r, 5).Font.Italic = True
        ws.Rows(r).RowHeight = 30
    End If
    r = r + 2
End Sub

'==========================================================================================
'  SUMMARY
'==========================================================================================
Private Sub BuildSummary()
    Dim ws As Worksheet, r As Long, i As Long, j As Long, cats As Collection, catList() As String, nCat As Long
    Set cats = New Collection
    Dim found As Boolean, v As Variant, sev As Variant, cnt As Long, src As Worksheet, nRows As Long
    Dim brands As Variant, k As Long, best As Long, bestScore As Double, used() As Boolean, picked As Long
    Dim t As Long, seen As String, key As String, ok As Boolean, x As Variant
    Set ws = NewSheet(OUT_SUMMARY)
    NoGrid
    ws.Columns(1).ColumnWidth = 22: ws.Columns(2).ColumnWidth = 30: ws.Columns(3).ColumnWidth = 22
    ws.Columns(4).ColumnWidth = 14: ws.Columns(5).ColumnWidth = 14: ws.Columns(6).ColumnWidth = 10: ws.Columns(7).ColumnWidth = 80
    ws.Cells(1, 1).Value = "UPdate audit - summary"
    ws.Cells(1, 1).Font.Bold = True: ws.Cells(1, 1).Font.Size = 14: ws.Cells(1, 1).Font.Color = RGB(31, 78, 120)
    ws.Cells(2, 1).Value = "Run " & Format(Now, "dd mmm yyyy") & " " & Format(Now, "hh:mm") & ".  Current month: '" & gName(1) & "' (" & gPeriod(1) & ")" & _
                           IIf(gHave(2), ";  prior month: '" & gName(2) & "' (" & gPeriod(2) & ")", ";  prior month: not found") & _
                           IIf(gHave(3), ";  YTD: '" & gName(3) & "' (" & gPeriod(3) & ")", ";  YTD: not found") & "."
    ws.Cells(2, 1).Font.Italic = True
    r = 3
    Dim tl As String
    For t = 1 To tCount: tl = tl & IIf(tl = "", "", ", ") & tName(t): Next t
    ws.Cells(r, 1).Value = "Trend sheets found: " & IIf(tl = "", "none", tl) & ".   Excel error values in the data: " & Format(gErrCount, "#,##0") & "."
    ws.Cells(r, 1).Font.Italic = True
    r = 5

    ' counts by category x severity
    Banner ws, r, "Findings by category (details in '" & OUT_ERRORS & "')", 7
    r = r + 1
    ReDim catList(1 To 50)
    For i = 1 To gN
        v = CGet(cats, CStr(gF(i)(1)), found)
        If Not found Then
            nCat = nCat + 1
            catList(nCat) = gF(i)(1)
            cats.Add nCat, CStr(gF(i)(1))
        End If
    Next i
    ws.Cells(r, 1).Value = "Category": ws.Cells(r, 2).Value = "High": ws.Cells(r, 3).Value = "Medium"
    ws.Cells(r, 4).Value = "Low": ws.Cells(r, 5).Value = "Info": ws.Cells(r, 6).Value = "Total"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 6)).Font.Bold = True
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 6)).Interior.Color = RGB(221, 235, 247)
    For j = 1 To nCat
        r = r + 1
        ws.Cells(r, 1).Value = catList(j)
        k = 0
        For Each sev In Array("High", "Medium", "Low", "Info")
            cnt = 0
            For i = 1 To gN
                If gF(i)(1) = catList(j) And gF(i)(0) = sev Then cnt = cnt + 1
            Next i
            k = k + 1
            ws.Cells(r, 1 + k).Value = cnt
        Next sev
        ws.Cells(r, 6).Formula = "=SUM(B" & r & ":E" & r & ")"
    Next j
    r = r + 1
    ws.Cells(r, 1).Value = "Total": ws.Cells(r, 1).Font.Bold = True
    For k = 2 To 6
        ws.Cells(r, k).Formula = "=SUM(" & ColL(k) & (r - nCat) & ":" & ColL(k) & (r - 1) & ")"
        ws.Cells(r, k).Font.Bold = True
    Next k
    r = r + 2

    ' top High findings (one or two per category)
    Banner ws, r, "Headline calculation errors (first High findings per category)", 7
    r = r + 1
    ws.Cells(r, 1).Value = "Category": ws.Cells(r, 2).Value = "Where": ws.Cells(r, 3).Value = "Item"
    ws.Cells(r, 4).Value = "Reported": ws.Cells(r, 5).Value = "Re-derived": ws.Cells(r, 6).Value = "Brand": ws.Cells(r, 7).Value = "Check / impact"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Font.Bold = True
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Interior.Color = RGB(221, 235, 247)
    Dim wantHigh As Boolean
    For j = 1 To nCat
        cnt = 0
        wantHigh = HasHigh(catList(j))
        For i = 1 To gN
            If gF(i)(1) = catList(j) And (gF(i)(0) = "High" Or Not wantHigh) Then
                r = r + 1: cnt = cnt + 1
                ws.Cells(r, 1).Value = gF(i)(1)
                ws.Cells(r, 2).Value = gF(i)(2) & " " & Left$(gF(i)(3), 40)
                ws.Cells(r, 3).Value = gF(i)(4)
                ws.Cells(r, 4).Value = gF(i)(7)
                ws.Cells(r, 5).Value = gF(i)(8)
                ws.Cells(r, 6).Value = gF(i)(5)
                ws.Cells(r, 7).Value = gF(i)(6) & " - " & gF(i)(10)
                ws.Cells(r, 7).WrapText = True
                For k = 4 To 5
                    If IsNum(ws.Cells(r, k).Value) Then
                        If Abs(ws.Cells(r, k).Value) < 5 And ws.Cells(r, k).Value <> Int(ws.Cells(r, k).Value) Then
                            ws.Cells(r, k).NumberFormat = "0.00%"
                        Else
                            ws.Cells(r, k).NumberFormat = "#,##0"
                        End If
                    End If
                Next k
                If cnt >= 2 Then Exit For
            End If
        Next i
    Next j
    r = r + 2

    ' clearest brand shifts
    Banner ws, r, "Clearest brand shifts - top 6 per brand by shift score (base >= 500; rates > 100% and tiny targets excluded)", 7
    r = r + 1
    ws.Cells(r, 1).Value = "Volumes: current month PROJECTED vs prior month actual. Rates: current actual vs prior actual (pp). Full list in '" & OUT_BRAND & "'."
    ws.Cells(r, 1).Font.Italic = True
    r = r + 1
    ws.Cells(r, 1).Value = "Brand": ws.Cells(r, 2).Value = "Item / Channel": ws.Cells(r, 3).Value = "Metric"
    ws.Cells(r, 4).Value = "Current": ws.Cells(r, 5).Value = "Prior": ws.Cells(r, 6).Value = "Score": ws.Cells(r, 7).Value = "Shift detail"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Font.Bold = True
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Interior.Color = RGB(221, 235, 247)
    Set src = ThisWorkbook.Worksheets(OUT_BRAND)
    nRows = LastRowOf(src)
    If nRows >= 2 Then
        Dim bd As Variant
        bd = src.Range(src.Cells(1, 1), src.Cells(nRows, 21)).Value
        ReDim used(1 To nRows)
        brands = Array("All brands", "Budget", "First For Women", "Dialdirect", "Auto & General", "Virseker")
        For k = 0 To 5
            seen = "|"
            For picked = 1 To 6
                best = 0: bestScore = 0
                For i = 2 To nRows
                    If Not used(i) And bd(i, 1) = brands(k) And bd(i, 18) = "YES" And IsNum(bd(i, 19)) Then
                        ok = (InStr(bd(i, 20), "data issue") = 0)
                        If ok And bd(i, 4) = "rate" Then
                            For Each x In Array(bd(i, 5), bd(i, 9), bd(i, 12))
                                If IsNum(x) Then If x > 1 Then ok = False
                            Next x
                        End If
                        If ok Then
                            For Each x In Array(bd(i, 8), bd(i, 11), bd(i, 14))
                                If IsNum(x) Then If Abs(x) > 3 Then ok = False
                            Next x
                        End If
                        key = bd(i, 3) & "~" & CStr(bd(i, 5)) & "~" & CStr(bd(i, 9))
                        If ok And InStr(seen, "|" & key & "|") > 0 Then ok = False
                        If ok And bd(i, 19) > bestScore Then
                            best = i
                            bestScore = bd(i, 19)
                        End If
                    End If
                Next i
                If best = 0 Then Exit For
                used(best) = True
                seen = seen & bd(best, 3) & "~" & CStr(bd(best, 5)) & "~" & CStr(bd(best, 9)) & "|"
                r = r + 1
                ws.Cells(r, 1).Value = bd(best, 1): ws.Cells(r, 2).Value = bd(best, 2): ws.Cells(r, 3).Value = bd(best, 3)
                If bd(best, 4) = "vol" Then
                    ws.Cells(r, 4).Value = bd(best, 6): ws.Cells(r, 4).NumberFormat = "#,##0": ws.Cells(r, 5).NumberFormat = "#,##0"
                Else
                    ws.Cells(r, 4).Value = bd(best, 5): ws.Cells(r, 4).NumberFormat = "0.0%": ws.Cells(r, 5).NumberFormat = "0.0%"
                End If
                ws.Cells(r, 5).Value = bd(best, 9)
                ws.Cells(r, 6).Value = bd(best, 19)
                ws.Cells(r, 7).Value = bd(best, 20)
                ws.Cells(r, 7).WrapText = True
            Next picked
        Next k
    End If
    r = r + 2
    Banner ws, r, "Re-running", 7
    r = r + 1
    ws.Cells(r, 1).Value = "Paste the refreshed Exec / trend sheets, then Alt+F8 > RunFullAudit. All audit sheets are rebuilt; your data sheets are never changed."
    r = r + 1
    ws.Cells(r, 1).Value = "Severity: High = wrong number / broken logic that changes a KPI; Medium = label, period or presentation issue; Low = rounding-sized; Info = expected structural gap."
    SetTab ws, RGB(31, 78, 120)
End Sub

Private Function HasHigh(ByVal cat As String) As Boolean
    Dim i As Long
    HasHigh = False
    For i = 1 To gN
        If gF(i)(1) = cat And gF(i)(0) = "High" Then
            HasHigh = True
            Exit Function
        End If
    Next i
End Function
