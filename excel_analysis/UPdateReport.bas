Attribute VB_Name = "UPdateReport"
Option Explicit
'==========================================================================================
' UPdateReport - replica views, recalculation checks, brand x source and Monday analysis
'
' HOW TO USE
'   1. Alt+F11 > File > Import File... > UPdateReport.bas   (remove the old UPdateAudit module)
'   2. Paste your refreshed sheets as usual.
'   3. Alt+F8 > RunReport > Run.   Save as .xlsm to keep the macro.
'
' DESIGN RULES (why the numbers stay right on any version of the document)
'   * Nothing is found by a fixed column letter or row number.
'       - Exec sheets: the row holding "From Date" is the header; metric groups ("Contacts",
'         "Leads", "Gross Sales" ...) and measures ("Actual", "Projecting", "Target", "Tar Dif",
'         "Tar Var", "W Target", "W Tar Var") are located by their header TEXT.
'       - Rows are located by their LABEL inside their own block (brand page / section).
'       - Summary: months, weeks and days are located from the "Period ->" and weekday rows;
'         every source x brand block is located by its heading ("Google - Budget", ...).
'   * Volumes are LINKS to your cells. Every %, difference and variance is a LIVE FORMULA.
'   * Rates are always ratio-of-totals (e.g. SUM(leads) / SUM(processed)), never an average
'     of daily percentages.
'   * Each recalculated value is compared with the value your sheet reports; differences are
'     highlighted and counted, so a wrong number can never pass silently.
'
' OUTPUT (rebuilt on every run; your sheets are never changed)
'   Report Checks              - the control sheet: what reconciles and what does not
'   Replica - <period> (x3)    - replica of each Exec sheet with recalculated values
'   Brand x Source             - Summary tab: every source by brand, current vs prior month
'   Monday View                - Summary tab: Mon vs Tue-Thu vs Fri vs Sat by source and brand
'   Weighted Conv & Closing    - how Conversion / Closing targets and W Targets are built
'==========================================================================================

Private Const OUT_CHECKS As String = "Report Checks"
Private Const OUT_BXS As String = "Brand x Source"
Private Const OUT_MON As String = "Monday View"
Private Const OUT_W As String = "Weighted Conv & Closing"
Private Const REP_PREFIX As String = "Replica - "

' colours
Private Const C_NAVY As Long = 7884319       ' RGB(31,78,120)
Private Const C_BLUE_LT As Long = 16247773   ' RGB(221,235,247)
Private Const C_GREY_LT As Long = 15921906   ' RGB(242,242,242)
Private Const C_GOLD_LT As Long = 10086143   ' RGB(255,230,153)
Private Const C_MISMATCH As Long = 11389944  ' RGB(248,203,173)
Private Const C_GREEN As Long = 32768        ' RGB(0,128,0)
Private Const C_RED As Long = 192            ' RGB(192,0,0)

' number formats: positive green, negative red, zero / other black
Private Const NF_INT As String = "#,##0;-#,##0;0"
Private Const NF_PCT As String = "0.0%;-0.0%;0.0%"
Private Const NF_DIF As String = "[Color10]+#,##0;[Red]-#,##0;0"
Private Const NF_VAR As String = "[Color10]+0.0%;[Red]-0.0%;0.0%"
Private Const NF_VAR_INV As String = "[Red]+0.0%;[Color10]-0.0%;0.0%"
Private Const NF_PP As String = "[Color10]+0.0"" pp"";[Red]-0.0"" pp"";0.0"" pp"""
Private Const NF_PP_INV As String = "[Red]+0.0"" pp"";[Color10]-0.0"" pp"";0.0"" pp"""
Private Const NF_CHK As String = "[Red]#,##0;[Red]-#,##0;[Color10]0"

' metric groups and measures in the Exec sheets
Private Const G_CONTACTS As Long = 1, G_CONV As Long = 2, G_LEADS As Long = 3, G_CLOSE As Long = 4
Private Const G_GSALES As Long = 5, G_GEFF As Long = 6, G_CXD As Long = 7, G_NSALES As Long = 8, G_NEFF As Long = 9
Private Const M_ACT As Long = 1, M_PROJ As Long = 2, M_TAR As Long = 3, M_DIF As Long = 4, M_VAR As Long = 5
Private Const M_WT As Long = 6, M_WVAR As Long = 7

Private gSilent As Boolean
Private gStep As String

' ---- Exec sheets: slot 1 = current month, 2 = prior month, 3 = YTD -------------------------
Private gHave(1 To 3) As Boolean
Private gName(1 To 3) As String
Private gPeriod(1 To 3) As String
Private gData(1 To 3) As Variant
Private gRows(1 To 3) As Long
Private gLastCol(1 To 3) As Long
Private gHdrRow(1 To 3) As Long
Private gLabCol(1 To 3) As Long
Private gCol(1 To 3, 1 To 9, 1 To 7) As Long
Private gGroupName(1 To 3, 1 To 9) As String
Private gLab() As String
Private gTag() As String
Private gBlk() As Long
Private gBB() As String
Private gSB() As String
Private gSP() As String
Private gNBlk(1 To 3) As Long
Private gBS() As Long
Private gBE() As Long
Private gRepName(1 To 3) As String
Private gRepCol(1 To 3, 1 To 9, 1 To 7) As Long
Private gRepLast(1 To 3) As Long
Private gRepMis(1 To 3) As Long
Private gRepHier(1 To 3) As Long

' ---- Summary sheet -------------------------------------------------------------------------
Private smHave As Boolean
Private smName As String
Private smData As Variant
Private smRows As Long
Private smCols As Long
Private smPerRow As Long
Private smDayRow As Long
Private smCapRow As Long
Private smCurCol As Long
Private smPrvCol As Long
Private smCurLab As String
Private smPrvLab As String
Private smD1 As Long
Private smD2 As Long
Private smNB As Long
Private smBH() As Long          ' heading row
Private smBSrc() As String
Private smBBrand() As String
Private smBKind() As String     ' dialler / inbound
Private smBInc() As Long
Private smBProc() As Long
Private smBLeads() As Long
Private smBSales() As Long
Private smBAtt() As Long
Private smBConv() As Long

'==========================================================================================
'  ENTRY POINTS
'==========================================================================================
Public Sub RunReport()
    Dim oldCalc As Long, t0 As Double, msg As String
    If Not gSilent Then
        If MsgBox("Build the report sheets?" & vbCrLf & vbCrLf & OUT_CHECKS & ", " & REP_PREFIX & "<period>, " & _
                  OUT_BXS & ", " & OUT_MON & ", " & OUT_W & vbCrLf & vbCrLf & _
                  "Only these sheets are replaced. Your data sheets are not changed.", _
                  vbOKCancel + vbQuestion, "UPdate report") <> vbOK Then Exit Sub
    End If
    t0 = Timer
    On Error GoTo Fail
    AppState True, oldCalc
    Status "removing previous report sheets"
    DeleteOutput
    Status "finding Exec sheets"
    DetectExec
    If Not gHave(1) Then
        AppState False, oldCalc
        If Not gSilent Then MsgBox "No Exec sheet found (a cell near the top-left must start with 'From Date').", vbExclamation
        Exit Sub
    End If
    Dim s As Long
    For s = 1 To 3
        If gHave(s) Then
            Status "building replica of " & gName(s)
            BuildReplica s
        End If
    Next s
    Status "finding the Summary sheet"
    DetectSummary
    If smHave Then
        Status "building Brand x Source"
        BuildBrandSource
        Status "building Monday View"
        BuildMondayView
    End If
    Status "building Weighted Conv & Closing"
    BuildWeighted
    Status "calculating"
    Application.Calculate
    Status "highlighting differences"
    For s = 1 To 3
        If gHave(s) Then HighlightReplica s
    Next s
    Status "building Report Checks"
    BuildChecks
    OrderSheets
    AppState False, oldCalc
    Status ""
    If Not gSilent Then MsgBox "Report built in " & Format(Timer - t0, "0") & " s. Start with '" & OUT_CHECKS & "'.", vbInformation, "UPdate report"
    Exit Sub
Fail:
    msg = "Report stopped while " & gStep & ": " & Err.Description & " (error " & Err.Number & ")"
    On Error Resume Next
    AppState False, oldCalc
    Status ""
    If gSilent Then
        Err.Raise 1002, "UPdateReport", msg
    Else
        MsgBox msg, vbCritical, "UPdate report"
    End If
End Sub

Public Sub RunReportSilent()
    gSilent = True
    RunReport
    gSilent = False
End Sub

Public Function ReportLastStep() As String
    ReportLastStep = gStep
End Function

'==========================================================================================
'  HELPERS
'==========================================================================================
Private Sub AppState(ByVal starting As Boolean, ByRef oldCalc As Long)
    On Error Resume Next
    If starting Then
        oldCalc = Application.Calculation
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        Application.EnableEvents = False
        Application.Calculation = -4135
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
        Application.StatusBar = "UPdate report: " & s & " ..."
    End If
End Sub

Private Function IsOutput(ByVal nm As String) As Boolean
    IsOutput = (nm = OUT_CHECKS Or nm = OUT_BXS Or nm = OUT_MON Or nm = OUT_W Or Left$(nm, Len(REP_PREFIX)) = REP_PREFIX)
    If Not IsOutput Then
        Select Case nm
            Case "Audit Summary", "Weighted Conv & Closing", "Calc Errors", "Brand x Item", "Trend Shifts", _
                 "Meta - Sheets", "Meta - Columns", "Meta - Rows", "READ ME", "Key Findings"
                IsOutput = True
        End Select
    End If
End Function

Private Sub DeleteOutput()
    Dim i As Long, nm As String
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        nm = ThisWorkbook.Worksheets(i).Name
        If nm = OUT_CHECKS Or nm = OUT_BXS Or nm = OUT_MON Or nm = OUT_W Or Left$(nm, Len(REP_PREFIX)) = REP_PREFIX Then
            ThisWorkbook.Worksheets(i).Delete
        End If
    Next i
End Sub

Private Function NewSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = nm
    ws.Cells.Font.Name = "Arial"
    ws.Cells.Font.Size = 9
    Set NewSheet = ws
End Function

Private Sub OrderSheets()
    On Error Resume Next
    Dim nm As Variant, i As Long, names() As String, n As Long
    ReDim names(1 To 10)
    n = 0
    n = n + 1
    names(n) = OUT_CHECKS
    For i = 1 To 3
        If gHave(i) Then
            n = n + 1
            names(n) = gRepName(i)
        End If
    Next i
    n = n + 1
    names(n) = OUT_BXS
    n = n + 1
    names(n) = OUT_MON
    n = n + 1
    names(n) = OUT_W
    For i = n To 1 Step -1
        If SheetExists(names(i)) Then ThisWorkbook.Worksheets(names(i)).Move Before:=ThisWorkbook.Worksheets(1)
    Next i
    ThisWorkbook.Worksheets(OUT_CHECKS).Activate
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

Private Function LastRowOf(ByVal ws As Worksheet) As Long
    LastRowOf = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
End Function

Private Function LastColOf(ByVal ws As Worksheet) As Long
    LastColOf = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
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

Private Function ColL(ByVal c As Long) As String
    Dim s As String, m As Long
    Do While c > 0
        m = (c - 1) Mod 26
        s = Chr$(65 + m) & s
        c = (c - m - 1) \ 26
    Loop
    ColL = s
End Function

Private Function QS(ByVal nm As String) As String
    QS = "'" & Replace(nm, "'", "''") & "'!"
End Function

Private Function BrandOf(ByVal s As String) As String
    Select Case LCase$(Trim$(s))
        Case "budget"
            BrandOf = "Budget"
        Case "ffw", "first for women"
            BrandOf = "First For Women"
        Case "dial", "dialdirect"
            BrandOf = "Dialdirect"
        Case "a&g", "auto & general"
            BrandOf = "Auto & General"
        Case "virseker"
            BrandOf = "Virseker"
        Case "tih multibrand"
            BrandOf = "TIH Multibrand"
        Case Else
            BrandOf = ""
    End Select
End Function

Private Function SafeText(ByVal s As String) As String
    If s = "" Then
        SafeText = ""
    ElseIf InStr("=+-@", Left$(s, 1)) > 0 Then
        SafeText = "'" & s
    Else
        SafeText = s
    End If
End Function

Private Sub HeaderStyle(ByVal rng As Range, Optional ByVal fill As Long = C_NAVY, Optional ByVal white As Boolean = True)
    With rng
        .Font.Bold = True
        If white Then .Font.Color = RGB(255, 255, 255) Else .Font.Color = RGB(0, 0, 0)
        .Interior.Color = fill
        .HorizontalAlignment = -4108
        .VerticalAlignment = -4108
        .WrapText = True
    End With
End Sub

Private Sub MergeTitle(ByVal ws As Worksheet, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long, ByVal text As String, _
                       Optional ByVal fill As Long = C_NAVY, Optional ByVal white As Boolean = True)
    ws.Cells(r, c1).Value = text
    If c2 > c1 Then ws.Range(ws.Cells(r, c1), ws.Cells(r, c2)).Merge
    HeaderStyle ws.Range(ws.Cells(r, c1), ws.Cells(r, c2)), fill, white
    If r = 1 Then ws.Cells(r, c1).HorizontalAlignment = -4131      ' sheet titles left-aligned
End Sub

Private Sub Freeze(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long)
    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(r, c).Select
    ActiveWindow.FreezePanes = True
End Sub

Private Sub NoGrid()
    If gSilent Then Exit Sub
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
End Sub

Private Sub SetTab(ByVal ws As Worksheet, ByVal clr As Long)
    On Error Resume Next
    ws.Tab.Color = clr
End Sub

Private Sub Box(ByVal rng As Range)
    On Error Resume Next
    With rng.Borders
        .LineStyle = 1
        .Weight = 2
        .Color = RGB(191, 191, 191)
    End With
End Sub

Private Function CGet(ByVal col As Collection, ByVal key As String, ByRef found As Boolean) As Variant
    On Error GoTo NotFound
    CGet = col.Item(key)
    found = True
    Exit Function
NotFound:
    found = False
    CGet = Empty
End Function

'==========================================================================================
'  EXEC SHEETS: detection, header mapping, row parsing
'==========================================================================================
Private Function FindFromDate(ByVal ws As Worksheet, ByRef hr As Long, ByRef lc As Long) As Boolean
    Dim r As Long, c As Long
    FindFromDate = False
    For r = 1 To 10
        For c = 1 To 8
            If Left$(Txt(ws.Cells(r, c).Value), 9) = "From Date" Then
                hr = r
                lc = c
                FindFromDate = True
                Exit Function
            End If
        Next c
    Next r
End Function

Private Sub DetectExec()
    Dim ws As Worksheet, hr As Long, lc As Long, p1 As Variant, p2 As Variant
    Dim curP As Double, prvP As Double, ytdP As Double, s As Long
    Dim curWs As Worksheet, prvWs As Worksheet, ytdWs As Worksheet
    For s = 1 To 3
        gHave(s) = False
    Next s
    For Each ws In ThisWorkbook.Worksheets
        If Not IsOutput(ws.Name) Then
            If FindFromDate(ws, hr, lc) Then
                p1 = ws.Cells(hr, lc - 1).Value
                p2 = ws.Cells(hr + 1, lc - 1).Value
                If IsNum(p1) And IsNum(p2) Then
                    If p1 = p2 Then
                        If p1 > curP Then
                            curP = p1
                            Set curWs = ws
                        End If
                    ElseIf p2 > ytdP Then
                        ytdP = p2
                        Set ytdWs = ws
                    End If
                End If
            End If
        End If
    Next ws
    If curWs Is Nothing Then Exit Sub
    For Each ws In ThisWorkbook.Worksheets
        If Not IsOutput(ws.Name) And ws.Name <> curWs.Name Then
            If FindFromDate(ws, hr, lc) Then
                p1 = ws.Cells(hr, lc - 1).Value
                p2 = ws.Cells(hr + 1, lc - 1).Value
                If IsNum(p1) And IsNum(p2) Then
                    If p1 = p2 And p1 < curP And p1 > prvP Then
                        prvP = p1
                        Set prvWs = ws
                    End If
                End If
            End If
        End If
    Next ws
    Dim maxR As Long
    maxR = LastRowOf(curWs)
    If Not prvWs Is Nothing Then
        If LastRowOf(prvWs) > maxR Then maxR = LastRowOf(prvWs)
    End If
    If Not ytdWs Is Nothing Then
        If LastRowOf(ytdWs) > maxR Then maxR = LastRowOf(ytdWs)
    End If
    ReDim gLab(1 To 3, 1 To maxR)
    ReDim gTag(1 To 3, 1 To maxR)
    ReDim gBlk(1 To 3, 1 To maxR)
    ReDim gBB(1 To 3, 1 To maxR)
    ReDim gSB(1 To 3, 1 To maxR)
    ReDim gSP(1 To 3, 1 To maxR)
    ReDim gBS(1 To 3, 0 To 500)
    ReDim gBE(1 To 3, 0 To 500)
    LoadExec 1, curWs
    If Not prvWs Is Nothing Then LoadExec 2, prvWs
    If Not ytdWs Is Nothing Then LoadExec 3, ytdWs
End Sub

Private Function GroupIndex(ByVal t As String) As Long
    Dim l As String
    l = LCase$(Trim$(t))
    GroupIndex = 0
    If l = "" Then Exit Function
    If InStr(l, "cxd") > 0 Then
        GroupIndex = G_CXD
    ElseIf InStr(l, "closing") > 0 Then
        GroupIndex = G_CLOSE
    ElseIf InStr(l, "gross effectiveness") > 0 Then
        GroupIndex = G_GEFF
    ElseIf InStr(l, "net effectiveness") > 0 Then
        GroupIndex = G_NEFF
    ElseIf InStr(l, "conversion") > 0 Then
        GroupIndex = G_CONV
    ElseIf InStr(l, "contacts") > 0 Then
        GroupIndex = G_CONTACTS
    ElseIf InStr(l, "leads") > 0 Then
        GroupIndex = G_LEADS
    ElseIf InStr(l, "net sales") > 0 Then
        GroupIndex = G_NSALES
    ElseIf InStr(l, "gross sales") > 0 Then
        GroupIndex = G_GSALES
    End If
End Function

Private Function MeasureIndex(ByVal t As String) As Long
    Dim l As String
    l = LCase$(Trim$(t))
    MeasureIndex = 0
    If l = "w tar var" Then
        MeasureIndex = M_WVAR
    ElseIf l = "w target" Then
        MeasureIndex = M_WT
    ElseIf l = "tar var" Then
        MeasureIndex = M_VAR
    ElseIf l = "tar dif" Or l = "tar diff" Then
        MeasureIndex = M_DIF
    ElseIf l = "target" Then
        MeasureIndex = M_TAR
    ElseIf l = "projecting" Or l = "projected" Then
        MeasureIndex = M_PROJ
    ElseIf l = "actual" Then
        MeasureIndex = M_ACT
    End If
End Function

Private Sub LoadExec(ByVal s As Long, ByVal ws As Worksheet)
    Dim hr As Long, lc As Long, nC As Long, c As Long, g As Long, curG As Long, m As Long, gi As Long, mi As Long
    gName(s) = ws.Name
    gHave(s) = True
    Call FindFromDate(ws, hr, lc)
    gHdrRow(s) = hr
    gLabCol(s) = lc
    gRows(s) = LastRowOf(ws)
    nC = LastColOf(ws)
    If nC < lc + 5 Then nC = lc + 5
    gLastCol(s) = nC
    gData(s) = ws.Range(ws.Cells(1, 1), ws.Cells(gRows(s), nC)).Value
    If ws.Cells(hr, lc - 1).Value = ws.Cells(hr + 1, lc - 1).Value Then
        gPeriod(s) = CStr(ws.Cells(hr, lc - 1).Value)
    Else
        gPeriod(s) = CStr(ws.Cells(hr, lc - 1).Value) & "-" & CStr(ws.Cells(hr + 1, lc - 1).Value)
    End If
    ' ---- header mapping by text ----
    For gi = 1 To 9
        gGroupName(s, gi) = ""
        For mi = 1 To 7
            gCol(s, gi, mi) = 0
        Next mi
    Next gi
    curG = 0
    For c = lc + 1 To nC
        g = GroupIndex(Txt(gData(s)(hr, c)))
        If g > 0 Then
            curG = g
            If gGroupName(s, g) = "" Then gGroupName(s, g) = Txt(gData(s)(hr, c))
        End If
        If curG > 0 Then
            m = MeasureIndex(Txt(gData(s)(hr + 1, c)))
            If m > 0 Then
                If gCol(s, curG, m) = 0 Then gCol(s, curG, m) = c
            End If
        End If
    Next c
    ParseExec s
End Sub

Private Sub ParseExec(ByVal s As Long)
    Dim r As Long, lab As String, tg As String, blk As Long, brand As String, parent As String, rb As String
    Dim isBrandRow As Boolean, lc As Long
    lc = gLabCol(s)
    blk = 0
    gBS(s, 0) = 1
    For r = 1 To gRows(s)
        lab = Txt(gData(s)(r, lc))
        If lc > 1 Then tg = Txt(gData(s)(r, lc - 1)) Else tg = ""
        gLab(s, r) = ""
        gTag(s, r) = tg
        gBB(s, r) = ""
        gSB(s, r) = ""
        gSP(s, r) = ""
        If Left$(lab, 9) = "From Date" Then
            gBE(s, blk) = r - 1
            blk = blk + 1
            gBS(s, blk) = r
            brand = ""
            parent = ""
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

Private Function FindInBlock(ByVal s As Long, ByVal b As Long, ByVal lab As String, ByVal after As Long) As Long
    Dim r As Long, r1 As Long, r2 As Long
    FindInBlock = 0
    If b < 0 Or b > gNBlk(s) Then Exit Function
    r1 = gBS(s, b)
    r2 = gBE(s, b)
    For r = r1 To r2
        If r > after Then
            If gLab(s, r) = lab And gSB(s, r) = "" Then
                FindInBlock = r
                Exit Function
            End If
        End If
    Next r
End Function

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

' Children rows of a parent row (comma separated list), "" if the row is not a parent here.
Private Function ChildrenOf(ByVal s As Long, ByVal r As Long) As String
    Dim lab As String, b As Long, rules As Variant, k As Long, parts As Variant, kids As Variant, i As Long
    Dim kr As Long, lst As String, n As Long, anywhere As Boolean, rr As Long
    ChildrenOf = ""
    lab = gLab(s, r)
    If lab = "" Then Exit Function
    b = gBlk(s, r)
    ' brand split parent: the brand rows that follow it
    If gSB(s, r) = "" Then
        lst = ""
        n = 0
        For rr = r + 1 To gBE(s, b)
            If gSB(s, rr) <> "" And gSP(s, rr) = lab Then
                lst = lst & IIf(lst = "", "", ",") & rr
                n = n + 1
            ElseIf gLab(s, rr) <> "" And gSB(s, rr) = "" Then
                Exit For
            End If
        Next rr
        If n >= 2 Then
            ChildrenOf = lst
            Exit Function
        End If
    End If
    If gSB(s, r) <> "" Then Exit Function
    ' brand total = Direct + Non-Direct of the same brand page
    If Right$(lab, 11) = "Brand Total" Then
        kr = FindInBlock(s, b, "Direct", r)
        rr = FindInBlock(s, b, "Non-Direct", r)
        If kr > 0 And rr > 0 Then ChildrenOf = kr & "," & rr
        Exit Function
    End If
    rules = HierarchyRules()
    For k = LBound(rules) To UBound(rules)
        parts = Split(rules(k), "|")
        If parts(0) = lab Then
            If FindInBlock(s, b, lab, 0) <> r Then Exit Function
            kids = Split(parts(1), ";")
            anywhere = (lab = "Total Personal Lines Incl MBO" Or lab = "Direct" Or lab = "Direct (Incl Outbound)" Or lab = "Non-Direct")
            lst = ""
            n = 0
            For i = 0 To UBound(kids)
                If anywhere Then
                    kr = FindInBlock(s, b, CStr(kids(i)), 0)
                Else
                    kr = FindInBlock(s, b, CStr(kids(i)), r)
                End If
                If kr > 0 And kr <> r Then
                    lst = lst & IIf(lst = "", "", ",") & kr
                    n = n + 1
                End If
            Next i
            If n >= 2 Then ChildrenOf = lst
            Exit Function
        End If
    Next k
End Function

'==========================================================================================
'  REPLICA
'==========================================================================================
Private Function RepSheetName(ByVal s As Long) As String
    Dim nm As String
    nm = REP_PREFIX & gPeriod(s)
    If s = 3 Then nm = REP_PREFIX & "YTD " & gPeriod(s)
    If Len(nm) > 31 Then nm = Left$(nm, 31)
    RepSheetName = nm
End Function

Private Function GroupOrder(ByVal s As Long) As Variant
    ' groups sorted by their position in the source sheet
    Dim g(1 To 9) As Long, c(1 To 9) As Long, i As Long, j As Long, n As Long, tg As Long, tc As Long
    Dim m As Long, firstC As Long, out() As Long
    n = 0
    For i = 1 To 9
        firstC = 0
        For m = 1 To 7
            If gCol(s, i, m) > 0 Then
                If firstC = 0 Or gCol(s, i, m) < firstC Then firstC = gCol(s, i, m)
            End If
        Next m
        If firstC > 0 Then
            n = n + 1
            g(n) = i
            c(n) = firstC
        End If
    Next i
    For i = 2 To n
        tg = g(i)
        tc = c(i)
        j = i - 1
        Do While j >= 1
            If c(j) <= tc Then Exit Do
            g(j + 1) = g(j)
            c(j + 1) = c(j)
            j = j - 1
        Loop
        g(j + 1) = tg
        c(j + 1) = tc
    Next i
    If n = 0 Then
        GroupOrder = Array()
        Exit Function
    End If
    ReDim out(0 To n - 1)
    For i = 1 To n
        out(i - 1) = g(i)
    Next i
    GroupOrder = out
End Function

Private Function MeasureOrder(ByVal s As Long, ByVal g As Long) As Variant
    Dim m(1 To 7) As Long, c(1 To 7) As Long, i As Long, j As Long, n As Long, tm As Long, tc As Long, out() As Long
    n = 0
    For i = 1 To 7
        If gCol(s, g, i) > 0 Then
            n = n + 1
            m(n) = i
            c(n) = gCol(s, g, i)
        End If
    Next i
    For i = 2 To n
        tm = m(i)
        tc = c(i)
        j = i - 1
        Do While j >= 1
            If c(j) <= tc Then Exit Do
            m(j + 1) = m(j)
            c(j + 1) = c(j)
            j = j - 1
        Loop
        m(j + 1) = tm
        c(j + 1) = tc
    Next i
    ReDim out(0 To n - 1)
    For i = 1 To n
        out(i - 1) = m(i)
    Next i
    MeasureOrder = out
End Function

Private Function MeasureLabel(ByVal m As Long) As String
    MeasureLabel = Choose(m, "Actual", "Projecting", "Target", "Tar Dif", "Tar Var", "W Target", "W Tar Var")
End Function

Private Function IsVolumeGroup(ByVal g As Long) As Boolean
    IsVolumeGroup = (g = G_CONTACTS Or g = G_LEADS Or g = G_GSALES Or g = G_NSALES)
End Function

' Is this replica measure recalculated (True) or linked (False)?
Private Function IsRecalc(ByVal g As Long, ByVal m As Long) As Boolean
    If IsVolumeGroup(g) Then
        IsRecalc = (m = M_DIF Or m = M_VAR)
    ElseIf g = G_CLOSE Or g = G_CXD Then
        IsRecalc = (m = M_ACT Or m = M_TAR Or m = M_VAR Or m = M_WVAR)
    Else
        IsRecalc = (m = M_VAR Or m = M_WVAR)
    End If
    If (g = G_CLOSE) And (m = M_TAR) Then
        ' closing target needs leads target and sales target
        IsRecalc = True
    End If
End Function

Private Function RepRef(ByVal s As Long, ByVal g As Long, ByVal m As Long, ByVal rr As Long) As String
    If gRepCol(s, g, m) > 0 Then RepRef = ColL(gRepCol(s, g, m)) & rr Else RepRef = ""
End Function

Private Function LinkF(ByVal src As String) As String
    LinkF = "=IF(ISNUMBER(" & src & ")," & src & ","""")"
End Function

Private Function RatioVarF(ByVal a As String, ByVal t As String) As String
    RatioVarF = "=IF(AND(ISNUMBER(" & a & "),ISNUMBER(" & t & ")),IF(" & t & "=0,""""," & a & "/" & t & "-1),"""")"
End Function

Private Function ReplicaFormula(ByVal s As Long, ByVal g As Long, ByVal m As Long, ByVal r As Long, ByVal rr As Long) As String
    Dim src As String, a As String, t As String, p As String, d As String, num As String, den As String
    src = QS(gName(s)) & ColL(gCol(s, g, m)) & r
    ReplicaFormula = LinkF(src)
    If IsVolumeGroup(g) Then
        If m = M_DIF Then
            p = RepRef(s, g, M_PROJ, rr)
            t = RepRef(s, g, M_TAR, rr)
            If p <> "" And t <> "" Then ReplicaFormula = "=IF(AND(ISNUMBER(" & p & "),ISNUMBER(" & t & "))," & p & "-" & t & ","""")"
        ElseIf m = M_VAR Then
            d = RepRef(s, g, M_DIF, rr)
            t = RepRef(s, g, M_TAR, rr)
            p = RepRef(s, g, M_PROJ, rr)
            If d <> "" And t <> "" Then
                ReplicaFormula = "=IF(AND(ISNUMBER(" & d & "),N(" & t & ")<>0)," & d & "/" & t & ","""")"
            ElseIf p <> "" And t <> "" Then
                ReplicaFormula = "=IF(AND(ISNUMBER(" & p & "),N(" & t & ")<>0)," & p & "/" & t & "-1,"""")"
            End If
        End If
        Exit Function
    End If
    If g = G_CLOSE And (m = M_ACT Or m = M_TAR) Then
        num = RepRef(s, G_GSALES, m, rr)
        den = RepRef(s, G_LEADS, m, rr)
        If num <> "" And den <> "" Then ReplicaFormula = "=IF(N(" & den & ")=0,"""",N(" & num & ")/N(" & den & "))"
        Exit Function
    End If
    If g = G_CXD And (m = M_ACT Or m = M_TAR) Then
        num = RepRef(s, G_NSALES, m, rr)
        den = RepRef(s, G_GSALES, m, rr)
        If num <> "" And den <> "" Then ReplicaFormula = "=IF(N(" & den & ")=0,"""",1-N(" & num & ")/N(" & den & "))"
        Exit Function
    End If
    If m = M_VAR Then
        a = RepRef(s, g, M_ACT, rr)
        t = RepRef(s, g, M_TAR, rr)
        If a <> "" And t <> "" Then ReplicaFormula = RatioVarF(a, t)
        Exit Function
    End If
    If m = M_WVAR Then
        a = RepRef(s, g, M_ACT, rr)
        t = RepRef(s, g, M_WT, rr)
        If a <> "" And t <> "" Then ReplicaFormula = RatioVarF(a, t)
        Exit Function
    End If
End Function

Private Function MeasureFormat(ByVal g As Long, ByVal m As Long) As String
    If IsVolumeGroup(g) Then
        If m = M_DIF Then
            MeasureFormat = NF_DIF
        ElseIf m = M_VAR Then
            MeasureFormat = NF_VAR
        Else
            MeasureFormat = NF_INT
        End If
    Else
        If m = M_VAR Or m = M_WVAR Then
            If g = G_CXD Then MeasureFormat = NF_VAR_INV Else MeasureFormat = NF_VAR
        Else
            MeasureFormat = NF_PCT
        End If
    End If
End Function

Private Sub BuildReplica(ByVal s As Long)
    Dim ws As Worksheet, gro As Variant, mo As Variant, gi As Long, mi As Long, g As Long, m As Long
    Dim col As Long, c0 As Long, r As Long, rr As Long, nRows As Long, lastC As Long, i As Long
    Dim fm() As Variant, lab As String, kids As String, kk As Variant, k As Long
    Dim cMis As Long, cHC As Long, cHL As Long, cHS As Long, cDC As Long, cDL As Long, cDS As Long
    Dim term As String, terms As String, tol As String, srcRef As String
    Dim hdrTop As Long, hdrSub As Long, firstRR As Long

    gRepName(s) = RepSheetName(s)
    Set ws = NewSheet(gRepName(s))
    NoGrid
    ' ---- column plan --------------------------------------------------------------------
    gro = GroupOrder(s)
    For gi = 1 To 9
        For mi = 1 To 7
            gRepCol(s, gi, mi) = 0
        Next mi
    Next gi
    col = 4
    For gi = 0 To UBound(gro)
        g = gro(gi)
        mo = MeasureOrder(s, g)
        For mi = 0 To UBound(mo)
            gRepCol(s, g, mo(mi)) = col
            col = col + 1
        Next mi
    Next gi
    cMis = col + 1
    cHC = cMis + 2
    cDC = cHC + 1
    cHL = cDC + 1
    cDL = cHL + 1
    cHS = cDL + 1
    cDS = cHS + 1
    lastC = cDS
    gRepLast(s) = lastC
    hdrTop = 4
    hdrSub = 5
    ' ---- titles -------------------------------------------------------------------------
    MergeTitle ws, 1, 1, lastC, "Replica of '" & gName(s) & "'  (period " & gPeriod(s) & ")", C_NAVY, True
    ws.Rows(1).RowHeight = 22
    ws.Cells(1, 1).Font.Size = 12
    ws.Cells(2, 1).Value = "Volumes (Actual / Projecting / Target, rates the source system calculates on eligible lines) are LINKED to the source cell. " & _
                           "Every Tar Dif, Tar Var, Closing, CXD and W Tar Var is RE-CALCULATED by formula. " & _
                           "Orange cell = the recalculated value differs from what the source sheet shows. Green = positive, red = negative, black = other."
    ws.Range(ws.Cells(2, 1), ws.Cells(2, lastC)).Merge
    ws.Cells(2, 1).WrapText = True
    ws.Cells(2, 1).Font.Italic = True
    ws.Rows(2).RowHeight = 26
    ' header rows
    ws.Cells(hdrTop, 1).Value = "Src row"
    ws.Cells(hdrTop, 2).Value = "Tag"
    ws.Cells(hdrTop, 3).Value = "Line"
    ws.Range(ws.Cells(hdrTop, 1), ws.Cells(hdrSub, 1)).Merge
    ws.Range(ws.Cells(hdrTop, 2), ws.Cells(hdrSub, 2)).Merge
    ws.Range(ws.Cells(hdrTop, 3), ws.Cells(hdrSub, 3)).Merge
    For gi = 0 To UBound(gro)
        g = gro(gi)
        mo = MeasureOrder(s, g)
        c0 = gRepCol(s, g, mo(0))
        MergeTitle ws, hdrTop, c0, gRepCol(s, g, mo(UBound(mo))), gGroupName(s, g), IIf(gi Mod 2 = 0, C_NAVY, RGB(47, 117, 181)), True
        For mi = 0 To UBound(mo)
            ws.Cells(hdrSub, gRepCol(s, g, mo(mi))).Value = MeasureLabel(mo(mi)) & IIf(IsRecalc(g, mo(mi)), " *", "")
        Next mi
    Next gi
    MergeTitle ws, hdrTop, cMis, cMis, "Check", RGB(192, 0, 0), True
    ws.Cells(hdrSub, cMis).Value = "Recalc differences"
    MergeTitle ws, hdrTop, cHC, cDS, "Parent line vs sum of its child lines", RGB(192, 0, 0), True
    ws.Cells(hdrSub, cHC).Value = "Contacts: children"
    ws.Cells(hdrSub, cDC).Value = "Contacts: diff"
    ws.Cells(hdrSub, cHL).Value = "Leads: children"
    ws.Cells(hdrSub, cDL).Value = "Leads: diff"
    ws.Cells(hdrSub, cHS).Value = "Gross Sales: children"
    ws.Cells(hdrSub, cDS).Value = "Gross Sales: diff"
    HeaderStyle ws.Range(ws.Cells(hdrTop, 1), ws.Cells(hdrSub, 3))
    HeaderStyle ws.Range(ws.Cells(hdrSub, 4), ws.Cells(hdrSub, lastC)), C_BLUE_LT, False
    ws.Rows(hdrSub).RowHeight = 26
    ' ---- formulas, built in one array ---------------------------------------------------
    firstRR = gHdrRow(s) + 2 + 2           ' source data starts two rows under "From Date"
    nRows = gRows(s) - (gHdrRow(s) + 2) + 1
    If nRows < 1 Then Exit Sub
    ReDim fm(1 To nRows, 1 To lastC)
    For r = gHdrRow(s) + 2 To gRows(s)
        rr = r + 2
        i = r - (gHdrRow(s) + 2) + 1
        lab = gLab(s, r)
        fm(i, 1) = r
        If gTag(s, r) <> "" Then fm(i, 2) = SafeText(gTag(s, r))
        If Left$(Txt(gData(s)(r, gLabCol(s))), 9) = "From Date" Or Left$(Txt(gData(s)(r, gLabCol(s))), 7) = "To Date" Then
            fm(i, 3) = SafeText(Txt(gData(s)(r, gLabCol(s))))
        ElseIf lab <> "" Or RowHasNumbers(s, r) Then
            If lab = "" Then
                fm(i, 3) = "(no label - stray value)"
            Else
                fm(i, 3) = SafeText(lab)
            End If
            terms = ""
            For gi = 0 To UBound(gro)
                g = gro(gi)
                mo = MeasureOrder(s, g)
                For mi = 0 To UBound(mo)
                    m = mo(mi)
                    If lab = "" Then
                        fm(i, gRepCol(s, g, m)) = LinkF(QS(gName(s)) & ColL(gCol(s, g, m)) & r)
                    Else
                        fm(i, gRepCol(s, g, m)) = ReplicaFormula(s, g, m, r, rr)
                    End If
                    If IsRecalc(g, m) And lab <> "" Then
                        srcRef = QS(gName(s)) & ColL(gCol(s, g, m)) & r
                        If IsVolumeGroup(g) And m = M_DIF Then tol = "0.5" Else tol = "0.0006"
                        term = "(ABS(N(" & ColL(gRepCol(s, g, m)) & rr & ")-N(" & srcRef & "))>" & tol & ")"
                        terms = terms & IIf(terms = "", "", "+") & term
                    End If
                Next mi
            Next gi
            If terms <> "" Then fm(i, cMis) = "=" & terms
            kids = ChildrenOf(s, r)
            If kids <> "" Then
                fm(i, cHC) = "=" & SumList(kids, gRepCol(s, G_CONTACTS, M_ACT))
                fm(i, cHL) = "=" & SumList(kids, gRepCol(s, G_LEADS, M_ACT))
                fm(i, cHS) = "=" & SumList(kids, gRepCol(s, G_GSALES, M_ACT))
                If gRepCol(s, G_CONTACTS, M_ACT) > 0 Then fm(i, cDC) = "=N(" & ColL(gRepCol(s, G_CONTACTS, M_ACT)) & rr & ")-N(" & ColL(cHC) & rr & ")"
                If gRepCol(s, G_LEADS, M_ACT) > 0 Then fm(i, cDL) = "=N(" & ColL(gRepCol(s, G_LEADS, M_ACT)) & rr & ")-N(" & ColL(cHL) & rr & ")"
                If gRepCol(s, G_GSALES, M_ACT) > 0 Then fm(i, cDS) = "=N(" & ColL(gRepCol(s, G_GSALES, M_ACT)) & rr & ")-N(" & ColL(cHS) & rr & ")"
            End If
        End If
    Next r
    WriteFormulas ws, firstRR, fm, nRows, lastC
    ' ---- formats ------------------------------------------------------------------------
    Dim lastRR As Long
    lastRR = firstRR + nRows - 1
    For gi = 0 To UBound(gro)
        g = gro(gi)
        mo = MeasureOrder(s, g)
        For mi = 0 To UBound(mo)
            ws.Range(ws.Cells(firstRR, gRepCol(s, g, mo(mi))), ws.Cells(lastRR, gRepCol(s, g, mo(mi)))).NumberFormat = MeasureFormat(g, mo(mi))
        Next mi
        ' group separator
        On Error Resume Next
        With ws.Range(ws.Cells(hdrTop, gRepCol(s, g, mo(0))), ws.Cells(lastRR, gRepCol(s, g, mo(0)))).Borders(7)
            .LineStyle = 1
            .Color = RGB(31, 78, 120)
        End With
        On Error GoTo 0
    Next gi
    ws.Range(ws.Cells(firstRR, cMis), ws.Cells(lastRR, cMis)).NumberFormat = NF_CHK
    ws.Range(ws.Cells(firstRR, cHC), ws.Cells(lastRR, cHC)).NumberFormat = NF_INT
    ws.Range(ws.Cells(firstRR, cHL), ws.Cells(lastRR, cHL)).NumberFormat = NF_INT
    ws.Range(ws.Cells(firstRR, cHS), ws.Cells(lastRR, cHS)).NumberFormat = NF_INT
    ws.Range(ws.Cells(firstRR, cDC), ws.Cells(lastRR, cDC)).NumberFormat = NF_CHK
    ws.Range(ws.Cells(firstRR, cDL), ws.Cells(lastRR, cDL)).NumberFormat = NF_CHK
    ws.Range(ws.Cells(firstRR, cDS), ws.Cells(lastRR, cDS)).NumberFormat = NF_CHK
    ws.Range(ws.Cells(firstRR, 1), ws.Cells(lastRR, 1)).Font.Color = RGB(128, 128, 128)
    ws.Range(ws.Cells(firstRR, 2), ws.Cells(lastRR, 2)).Font.Color = RGB(128, 128, 128)
    ' ---- row styling --------------------------------------------------------------------
    For r = gHdrRow(s) + 2 To gRows(s)
        rr = r + 2
        lab = Txt(gData(s)(r, gLabCol(s)))
        If Left$(lab, 9) = "From Date" Then
            ws.Cells(rr, 3).Value = lab & "   |   " & Txt(gData(s)(r + 1, gLabCol(s)))
            ws.Range(ws.Cells(rr, 3), ws.Cells(rr, lastC)).Merge
            HeaderStyle ws.Range(ws.Cells(rr, 1), ws.Cells(rr, lastC))
            ws.Cells(rr, 3).HorizontalAlignment = -4131
            ' repeat measure headers on the To Date row
            For gi = 0 To UBound(gro)
                g = gro(gi)
                mo = MeasureOrder(s, g)
                For mi = 0 To UBound(mo)
                    ws.Cells(rr + 1, gRepCol(s, g, mo(mi))).Value = Left$(gGroupName(s, g), 12) & " " & MeasureLabel(mo(mi))
                Next mi
            Next gi
            ws.Cells(rr + 1, 3).Value = ""
            With ws.Range(ws.Cells(rr + 1, 4), ws.Cells(rr + 1, lastC))
                .Font.Italic = True
                .Font.Size = 8
                .Font.Color = RGB(89, 89, 89)
                .Interior.Color = C_BLUE_LT
                .WrapText = True
            End With
        ElseIf lab = "" And gTag(s, r) <> "" And Not RowHasNumbers(s, r) Then
            If BrandOf(gTag(s, r)) <> "" Or UCase$(gTag(s, r)) = gTag(s, r) Then
                ws.Cells(rr, 3).Value = UCase$(gTag(s, r))
                ws.Range(ws.Cells(rr, 3), ws.Cells(rr, lastC)).Merge
                HeaderStyle ws.Range(ws.Cells(rr, 1), ws.Cells(rr, lastC)), C_GOLD_LT, False
                ws.Cells(rr, 3).HorizontalAlignment = -4131
            End If
        ElseIf gLab(s, r) <> "" Then
            If Right$(gLab(s, r), 11) = "Brand Total" Or gLab(s, r) = "Total Personal Lines Incl MBO" Then
                ws.Range(ws.Cells(rr, 1), ws.Cells(rr, lastC)).Interior.Color = C_GOLD_LT
                ws.Range(ws.Cells(rr, 3), ws.Cells(rr, lastC)).Font.Bold = True
            ElseIf ChildrenOf(s, r) <> "" Then
                ws.Range(ws.Cells(rr, 3), ws.Cells(rr, cMis - 1)).Font.Bold = True
                ws.Range(ws.Cells(rr, 1), ws.Cells(rr, cMis - 1)).Interior.Color = C_GREY_LT
            End If
            If gSB(s, r) <> "" Then ws.Cells(rr, 3).IndentLevel = 2
        ElseIf RowHasNumbers(s, r) Then
            ws.Cells(rr, 3).Font.Color = RGB(192, 0, 0)
            ws.Cells(rr, 3).Font.Italic = True
        End If
    Next r
    ' ---- layout -------------------------------------------------------------------------
    ws.Columns(1).ColumnWidth = 6
    ws.Columns(2).ColumnWidth = 10
    ws.Columns(3).ColumnWidth = 38
    ws.Range(ws.Cells(1, 4), ws.Cells(1, lastC)).ColumnWidth = 10
    ws.Columns(cMis - 1).ColumnWidth = 2
    ws.Columns(cMis + 1).ColumnWidth = 2
    ws.Columns(cMis).ColumnWidth = 11
    Box ws.Range(ws.Cells(hdrTop, 1), ws.Cells(lastRR, cMis))
    Box ws.Range(ws.Cells(hdrTop, cHC), ws.Cells(lastRR, cDS))
    Freeze ws, firstRR, 4
    SetTab ws, RGB(0, 112, 192)
End Sub

Private Function RowHasNumbers(ByVal s As Long, ByVal r As Long) As Boolean
    Dim c As Long
    RowHasNumbers = False
    For c = gLabCol(s) + 1 To gLastCol(s)
        If IsNum(gData(s)(r, c)) Then
            RowHasNumbers = True
            Exit Function
        End If
    Next c
End Function

Private Function NoVolumes(ByVal s As Long, ByVal r As Long) As Boolean
    Dim cC As Long, cL As Long
    cC = gCol(s, G_CONTACTS, M_ACT)
    cL = gCol(s, G_LEADS, M_ACT)
    NoVolumes = True
    If cC > 0 Then
        If IsNum(gData(s)(r, cC)) Then NoVolumes = False
    End If
    If cL > 0 Then
        If IsNum(gData(s)(r, cL)) Then NoVolumes = False
    End If
End Function

Private Function SumList(ByVal rowsCsv As String, ByVal col As Long) As String
    Dim p As Variant, i As Long, s As String
    If col = 0 Then
        SumList = "0"
        Exit Function
    End If
    p = Split(rowsCsv, ",")
    s = ""
    For i = 0 To UBound(p)
        s = s & IIf(s = "", "", ",") & ColL(col) & (CLng(p(i)) + 2)
    Next i
    SumList = "SUM(" & s & ")"
End Function

Private Sub WriteFormulas(ByVal ws As Worksheet, ByVal topRow As Long, fm As Variant, ByVal nRows As Long, ByVal nCols As Long)
    Dim r As Long, c As Long
    On Error GoTo CellByCell
    ws.Range(ws.Cells(topRow, 1), ws.Cells(topRow + nRows - 1, nCols)).Formula = fm
    Exit Sub
CellByCell:
    Resume CellLoop
CellLoop:
    On Error GoTo 0
    For r = 1 To nRows
        For c = 1 To nCols
            If Not IsEmpty(fm(r, c)) Then ws.Cells(topRow + r - 1, c).Formula = fm(r, c)
        Next c
    Next r
End Sub

Private Sub HighlightReplica(ByVal s As Long)
    ' after calculation: colour every recalculated cell that disagrees with the source
    Dim ws As Worksheet, v As Variant, r As Long, rr As Long, gi As Long, mi As Long, g As Long, m As Long
    Dim gro As Variant, mo As Variant, rv As Variant, sv As Variant, tol As Double, n As Long, top As Long
    Dim cMis As Long, hier As Long, d As Variant
    Set ws = ThisWorkbook.Worksheets(gRepName(s))
    top = gHdrRow(s) + 4
    v = ws.Range(ws.Cells(1, 1), ws.Cells(gRows(s) + 2, gRepLast(s))).Value
    gro = GroupOrder(s)
    n = 0
    For r = gHdrRow(s) + 2 To gRows(s)
        rr = r + 2
        If gLab(s, r) <> "" Then
            For gi = 0 To UBound(gro)
                g = gro(gi)
                mo = MeasureOrder(s, g)
                For mi = 0 To UBound(mo)
                    m = mo(mi)
                    If IsRecalc(g, m) Then
                        rv = v(rr, gRepCol(s, g, m))
                        sv = gData(s)(r, gCol(s, g, m))
                        If IsVolumeGroup(g) And m = M_DIF Then tol = 0.5 Else tol = 0.0006
                        If DiffersN(rv, sv, tol) Then
                            ws.Cells(rr, gRepCol(s, g, m)).Interior.Color = C_MISMATCH
                            ws.Cells(rr, gRepCol(s, g, m)).Font.Bold = True
                            n = n + 1
                        End If
                    End If
                Next mi
            Next gi
        End If
    Next r
    gRepMis(s) = n
    ' count parent / child differences
    hier = 0
    For rr = top To gRows(s) + 2
        For mi = gRepLast(s) - 4 To gRepLast(s) Step 2
            d = v(rr, mi)
            If IsNum(d) Then
                If Abs(d) > 0.5 Then
                    hier = hier + 1
                    ws.Cells(rr, mi).Interior.Color = C_MISMATCH
                End If
            End If
        Next mi
    Next rr
    gRepHier(s) = hier
End Sub

Private Function NVal(ByVal v As Variant) As Double
    If IsNum(v) Then NVal = v Else NVal = 0
End Function

Private Function DiffersN(ByVal a As Variant, ByVal b As Variant, ByVal tol As Double) As Boolean
    DiffersN = (Abs(NVal(a) - NVal(b)) > tol)
End Function

'==========================================================================================
'  SUMMARY SHEET: detection and block map
'==========================================================================================
Private Sub DetectSummary()
    Dim ws As Worksheet, r As Long, c As Long, found As Boolean, nR As Long, nC As Long
    Dim capDaily As Long, capYest As Long, v As Variant, lastMon As Long, prevMon As Long
    smHave = False
    For Each ws In ThisWorkbook.Worksheets
        If Not IsOutput(ws.Name) Then
            For r = 1 To 12
                For c = 1 To 6
                    If Left$(Txt(ws.Cells(r, c).Value), 6) = "Period" Then
                        If InStr(1, Txt(ws.Cells(r + 1, c).Value), "Working Days", vbTextCompare) > 0 Then
                            smName = ws.Name
                            smPerRow = r
                            smDayRow = r + 1
                            smCapRow = r - 1
                            found = True
                            Exit For
                        End If
                    End If
                Next c
                If found Then Exit For
            Next r
        End If
        If found Then Exit For
    Next ws
    If Not found Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(smName)
    nR = LastRowOf(ws)
    nC = LastColOf(ws)
    smRows = nR
    smCols = nC
    smData = ws.Range(ws.Cells(1, 1), ws.Cells(nR, nC)).Value
    ' months: YYYYMM numbers in the period row
    lastMon = 0
    prevMon = 0
    For c = 1 To nC
        v = smData(smPerRow, c)
        If IsNum(v) Then
            If v >= 200001 And v <= 209912 And v = Int(v) Then
                If v > lastMon Then
                    prevMon = lastMon
                    smPrvCol = smCurCol
                    lastMon = v
                    smCurCol = c
                ElseIf v > prevMon And v < lastMon Then
                    prevMon = v
                    smPrvCol = c
                End If
            End If
        End If
    Next c
    If smCurCol > 0 Then smCurLab = CStr(smData(smPerRow, smCurCol))
    If smPrvCol > 0 Then smPrvLab = CStr(smData(smPerRow, smPrvCol))
    ' daily columns: from the "Daily" caption up to the "Yesterday" caption, weekday names in the day row
    capDaily = 0
    capYest = 0
    For c = 1 To nC
        If InStr(1, Txt(smData(smCapRow, c)), "Daily", vbTextCompare) > 0 And capDaily = 0 Then capDaily = c
        If InStr(1, Txt(smData(smCapRow, c)), "Yesterday", vbTextCompare) > 0 And capYest = 0 Then capYest = c
    Next c
    smD1 = 0
    smD2 = 0
    If capDaily > 0 Then
        If capYest = 0 Then capYest = nC + 1
        For c = capDaily To capYest - 1
            If IsWeekday(Txt(smData(smDayRow, c))) Then
                If smD1 = 0 Then smD1 = c
                smD2 = c
            End If
        Next c
    End If
    MapSummaryBlocks
    smHave = (smNB > 0)
End Sub

Private Function IsWeekday(ByVal t As String) As Boolean
    Select Case t
        Case "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"
            IsWeekday = True
        Case Else
            IsWeekday = False
    End Select
End Function

Private Function SLab(ByVal r As Long) As String
    If r < 1 Or r > smRows Then
        SLab = ""
    Else
        SLab = Txt(smData(r, 3))
    End If
End Function

Private Sub MapSummaryBlocks()
    Dim r As Long, h As String, nxt As String, k As Long, l As String, kind As String
    Dim src As String, brand As String, p As Long, b As String, cap As Long
    cap = 400
    ReDim smBH(1 To cap)
    ReDim smBSrc(1 To cap)
    ReDim smBBrand(1 To cap)
    ReDim smBKind(1 To cap)
    ReDim smBInc(1 To cap)
    ReDim smBProc(1 To cap)
    ReDim smBLeads(1 To cap)
    ReDim smBSales(1 To cap)
    ReDim smBAtt(1 To cap)
    ReDim smBConv(1 To cap)
    smNB = 0
    For r = smDayRow + 1 To smRows - 1
        h = SLab(r)
        nxt = SLab(r + 1)
        kind = ""
        If h <> "" And nxt = "Incoming Contacts" Then kind = "dialler"
        If h <> "" And nxt = "Answered Calls" Then kind = "inbound"
        If kind <> "" And smNB < cap Then
            smNB = smNB + 1
            smBH(smNB) = r
            smBKind(smNB) = kind
            ' source / brand from the heading
            src = h
            brand = "All brands"
            p = InStrRev(h, " - ")
            If p > 0 Then
                b = BrandOf(Mid$(h, p + 3))
                If b <> "" Then
                    src = Left$(h, p - 1)
                    brand = b
                End If
            End If
            If kind = "inbound" Then
                If BrandOf(h) <> "" Then
                    brand = BrandOf(h)
                    src = "JHB Inbound (Connex)"
                ElseIf h = "Total Inbound" Then
                    src = "JHB Inbound (Connex)"
                End If
            End If
            smBSrc(smNB) = src
            smBBrand(smNB) = brand
            ' metric rows inside the block (until the next heading)
            For k = r + 1 To r + 26
                If k > smRows Then Exit For
                l = SLab(k)
                If k > r + 1 And (SLab(k + 1) = "Incoming Contacts" Or SLab(k + 1) = "Answered Calls") Then Exit For
                If kind = "dialler" Then
                    If l = "Incoming Contacts" And smBInc(smNB) = 0 Then smBInc(smNB) = k
                    If l = "Processed Contacts" And smBProc(smNB) = 0 Then smBProc(smNB) = k
                    If l = "Leads" And smBLeads(smNB) = 0 Then smBLeads(smNB) = k
                    If l = "Gross Sales" And smBSales(smNB) = 0 Then smBSales(smNB) = k
                    If l = "Attempts" And smBAtt(smNB) = 0 Then smBAtt(smNB) = k
                    If l = "Conversion" And smBConv(smNB) = 0 Then smBConv(smNB) = k
                Else
                    If l = "Answered Calls" And smBInc(smNB) = 0 Then smBInc(smNB) = k
                    If l = "N ACD Calls" And smBProc(smNB) = 0 Then smBProc(smNB) = k
                    If l = "Leads" And smBLeads(smNB) = 0 Then smBLeads(smNB) = k
                    If (l = "Sales" Or l = "Gross Sales") And smBSales(smNB) = 0 Then smBSales(smNB) = k
                    If l = "Conversion" And smBConv(smNB) = 0 Then smBConv(smNB) = k
                End If
            Next k
            If smBProc(smNB) = 0 Or smBLeads(smNB) = 0 Or smBSales(smNB) = 0 Then smNB = smNB - 1
        End If
    Next r
End Sub

Private Function SRef(ByVal r As Long, ByVal c As Long) As String
    SRef = QS(smName) & ColL(c) & r
End Function

Private Function SRng(ByVal r As Long) As String
    SRng = QS(smName) & "$" & ColL(smD1) & "$" & r & ":$" & ColL(smD2) & "$" & r
End Function

Private Function DayRow() As String
    DayRow = QS(smName) & "$" & ColL(smD1) & "$" & smDayRow & ":$" & ColL(smD2) & "$" & smDayRow
End Function

Private Function DayCond(ByVal which As String) As String
    Dim w As String
    w = DayRow()
    Select Case which
        Case "TueThu"
            DayCond = "((" & w & "=""Tue"")+(" & w & "=""Wed"")+(" & w & "=""Thu""))"
        Case Else
            DayCond = "(" & w & "=""" & which & """)"
    End Select
End Function

Private Function DaySum(ByVal r As Long, ByVal which As String) As String
    DaySum = "SUMPRODUCT(" & DayCond(which) & "*1," & SRng(r) & ")"
End Function

Private Function DayCount(ByVal which As String) As String
    DayCount = "SUMPRODUCT(" & DayCond(which) & "*1)"
End Function

Private Function BrandList6() As Variant
    BrandList6 = Array("All brands", "Budget", "First For Women", "Dialdirect", "Auto & General", "Virseker")
End Function

'==========================================================================================
'  BRAND x SOURCE (current month vs prior month)
'==========================================================================================
Private Sub BuildBrandSource()
    Dim ws As Worksheet, r As Long, bi As Long, brands As Variant, i As Long, c As Long, first As Long, n As Long
    Dim hdr As Variant, j As Long, lastC As Long, cur As Long, prv As Long, rr As Long
    Set ws = NewSheet(OUT_BXS)
    NoGrid
    lastC = 21
    MergeTitle ws, 1, 1, lastC, "Brand x Source - '" & smName & "' tab  (current month " & smCurLab & " vs prior month " & smPrvLab & ")", C_NAVY, True
    ws.Rows(1).RowHeight = 22
    ws.Cells(1, 1).Font.Size = 12
    ws.Cells(2, 1).Value = "Volumes link to the Summary tab. Conversion = Leads / Processed, Closing = Gross Sales / Leads, Effectiveness = Gross Sales / Processed " & _
                           "(inbound: Processed = N ACD calls, so conversion can exceed 100%). Changes in percentage points (pp). Green = up, red = down. " & _
                           "Check column: recalculated conversion minus the conversion the Summary tab reports (should be 0)."
    ws.Range(ws.Cells(2, 1), ws.Cells(2, lastC)).Merge
    ws.Cells(2, 1).WrapText = True
    ws.Cells(2, 1).Font.Italic = True
    ws.Rows(2).RowHeight = 30
    cur = smCurCol
    prv = smPrvCol
    r = 4
    brands = BrandList6()
    For bi = 0 To 5
        n = 0
        For i = 1 To smNB
            If smBBrand(i) = brands(bi) Then n = n + 1
        Next i
        If n = 0 Then GoTo NextBrand
        MergeTitle ws, r, 1, lastC, UCase$(brands(bi)), C_GOLD_LT, False
        ws.Cells(r, 1).HorizontalAlignment = -4131
        r = r + 1
        ' group header row
        MergeTitle ws, r, 1, 2, "Source", C_NAVY, True
        MergeTitle ws, r, 3, 9, "Current month " & smCurLab, C_NAVY, True
        MergeTitle ws, r, 10, 15, "Prior month " & smPrvLab, RGB(47, 117, 181), True
        MergeTitle ws, r, 16, 19, "Change (current vs prior)", C_NAVY, True
        MergeTitle ws, r, 20, 21, "Check vs Summary tab", RGB(192, 0, 0), True
        r = r + 1
        hdr = Array("Source", "Summary row", "Incoming", "Processed", "Leads", "Gross Sales", "Conversion", "Closing", "Effectiveness", _
                    "Processed", "Leads", "Gross Sales", "Conversion", "Closing", "Effectiveness", _
                    "Processed %", "Conversion", "Closing", "Effectiveness", "Reported conversion", "Recalc - reported")
        For j = 0 To UBound(hdr)
            ws.Cells(r, j + 1).Value = hdr(j)
        Next j
        HeaderStyle ws.Range(ws.Cells(r, 1), ws.Cells(r, lastC)), C_BLUE_LT, False
        ws.Rows(r).RowHeight = 26
        first = r + 1
        For i = 1 To smNB
            If smBBrand(i) = brands(bi) Then
                r = r + 1
                ws.Cells(r, 1).Value = SafeText(smBSrc(i))
                ws.Cells(r, 2).Value = smBH(i)
                If smBInc(i) > 0 Then ws.Cells(r, 3).Formula = LinkF(SRef(smBInc(i), cur))
                ws.Cells(r, 4).Formula = LinkF(SRef(smBProc(i), cur))
                ws.Cells(r, 5).Formula = LinkF(SRef(smBLeads(i), cur))
                ws.Cells(r, 6).Formula = LinkF(SRef(smBSales(i), cur))
                ws.Cells(r, 7).Formula = "=IF(N(D" & r & ")=0,"""",E" & r & "/D" & r & ")"
                ws.Cells(r, 8).Formula = "=IF(N(E" & r & ")=0,"""",F" & r & "/E" & r & ")"
                ws.Cells(r, 9).Formula = "=IF(N(D" & r & ")=0,"""",F" & r & "/D" & r & ")"
                If prv > 0 Then
                    ws.Cells(r, 10).Formula = LinkF(SRef(smBProc(i), prv))
                    ws.Cells(r, 11).Formula = LinkF(SRef(smBLeads(i), prv))
                    ws.Cells(r, 12).Formula = LinkF(SRef(smBSales(i), prv))
                    ws.Cells(r, 13).Formula = "=IF(N(J" & r & ")=0,"""",K" & r & "/J" & r & ")"
                    ws.Cells(r, 14).Formula = "=IF(N(K" & r & ")=0,"""",L" & r & "/K" & r & ")"
                    ws.Cells(r, 15).Formula = "=IF(N(J" & r & ")=0,"""",L" & r & "/J" & r & ")"
                    ws.Cells(r, 16).Formula = "=IF(N(J" & r & ")=0,"""",D" & r & "/J" & r & "-1)"
                    ws.Cells(r, 17).Formula = "=IF(AND(ISNUMBER(G" & r & "),ISNUMBER(M" & r & ")),(G" & r & "-M" & r & ")*100,"""")"
                    ws.Cells(r, 18).Formula = "=IF(AND(ISNUMBER(H" & r & "),ISNUMBER(N" & r & ")),(H" & r & "-N" & r & ")*100,"""")"
                    ws.Cells(r, 19).Formula = "=IF(AND(ISNUMBER(I" & r & "),ISNUMBER(O" & r & ")),(I" & r & "-O" & r & ")*100,"""")"
                End If
                If smBConv(i) > 0 Then
                    ws.Cells(r, 20).Formula = LinkF(SRef(smBConv(i), cur))
                    ws.Cells(r, 21).Formula = "=IF(AND(ISNUMBER(G" & r & "),ISNUMBER(T" & r & ")),(G" & r & "-T" & r & ")*100,"""")"
                End If
                If smBSrc(i) Like "Dialler - Total*" Or smBSrc(i) Like "JHB*" Or smBSrc(i) Like "KZN*" Or brands(bi) <> "All brands" And i = FirstBlockOfBrand(brands(bi)) Then
                    ws.Range(ws.Cells(r, 1), ws.Cells(r, lastC)).Font.Bold = True
                End If
            End If
        Next i
        ' formats for this table
        ws.Range(ws.Cells(first, 3), ws.Cells(r, 6)).NumberFormat = NF_INT
        ws.Range(ws.Cells(first, 10), ws.Cells(r, 12)).NumberFormat = NF_INT
        ws.Range(ws.Cells(first, 7), ws.Cells(r, 9)).NumberFormat = NF_PCT
        ws.Range(ws.Cells(first, 13), ws.Cells(r, 15)).NumberFormat = NF_PCT
        ws.Range(ws.Cells(first, 16), ws.Cells(r, 16)).NumberFormat = NF_VAR
        ws.Range(ws.Cells(first, 17), ws.Cells(r, 19)).NumberFormat = NF_PP
        ws.Range(ws.Cells(first, 20), ws.Cells(r, 20)).NumberFormat = NF_PCT
        ws.Range(ws.Cells(first, 21), ws.Cells(r, 21)).NumberFormat = "[Red]+0.00"" pp"";[Red]-0.00"" pp"";[Color10]0.00"" pp"""
        ws.Range(ws.Cells(first, 2), ws.Cells(r, 2)).Font.Color = RGB(128, 128, 128)
        Box ws.Range(ws.Cells(first - 2, 1), ws.Cells(r, lastC))
        r = r + 2
NextBrand:
    Next bi
    ws.Columns(1).ColumnWidth = 34
    ws.Columns(2).ColumnWidth = 8
    ws.Range(ws.Cells(1, 3), ws.Cells(1, lastC)).ColumnWidth = 11
    Freeze ws, 4, 3
    SetTab ws, RGB(0, 176, 80)
End Sub

Private Function FirstBlockOfBrand(ByVal brand As String) As Long
    Dim i As Long
    FirstBlockOfBrand = 0
    For i = 1 To smNB
        If smBBrand(i) = brand Then
            FirstBlockOfBrand = i
            Exit Function
        End If
    Next i
End Function

'==========================================================================================
'  MONDAY VIEW (weekday profile, ratio of totals)
'==========================================================================================
Private Sub BuildMondayView()
    Dim ws As Worksheet, r As Long, bi As Long, brands As Variant, i As Long, j As Long, n As Long, first As Long
    Dim days As Variant, d As Long, c As Long, lastC As Long, hdr As Variant, pr As Long, ld As Long, sl As Long
    Dim fP As String, fL As String, fS As String, cnt As String
    If smD1 = 0 Then Exit Sub
    Set ws = NewSheet(OUT_MON)
    NoGrid
    days = Array("Mon", "TueThu", "Fri", "Sat")
    lastC = 2 + 4 * 4 + 4 + 2
    MergeTitle ws, 1, 1, lastC, "Monday View - weekday profile by source and brand  ('" & smName & "' daily columns " & _
               Format(smData(smPerRow, smD1), "dd mmm") & " to " & Format(smData(smPerRow, smD2), "dd mmm yyyy") & ")", C_NAVY, True
    ws.Rows(1).RowHeight = 22
    ws.Cells(1, 1).Font.Size = 12
    ws.Cells(2, 1).Value = "Every rate is a ratio of TOTALS over all days of that weekday in the daily columns (e.g. Monday conversion = all Monday leads / all Monday " & _
                           "processed contacts) - not an average of daily percentages. Processed / day = total processed / number of those weekdays. " & _
                           "Gap = Monday minus Tue-Thu (pp). Sales lost / Monday = Monday processed per day x (Tue-Thu effectiveness - Monday effectiveness)."
    ws.Range(ws.Cells(2, 1), ws.Cells(2, lastC)).Merge
    ws.Cells(2, 1).WrapText = True
    ws.Cells(2, 1).Font.Italic = True
    ws.Rows(2).RowHeight = 30
    r = 4
    brands = BrandList6()
    For bi = 0 To 5
        n = 0
        For i = 1 To smNB
            If smBBrand(i) = brands(bi) Then n = n + 1
        Next i
        If n = 0 Then GoTo NextBrand
        MergeTitle ws, r, 1, lastC, UCase$(brands(bi)), C_GOLD_LT, False
        ws.Cells(r, 1).HorizontalAlignment = -4131
        r = r + 1
        MergeTitle ws, r, 1, 2, "Source", C_NAVY, True
        For d = 0 To 3
            c = 3 + d * 4
            ws.Cells(r, c).Formula = "=""" & IIf(days(d) = "TueThu", "Tue-Thu", days(d)) & " (""&" & DayCount(CStr(days(d))) & "&"" days)"""
            ws.Range(ws.Cells(r, c), ws.Cells(r, c + 3)).Merge
            HeaderStyle ws.Range(ws.Cells(r, c), ws.Cells(r, c + 3)), IIf(d = 0, RGB(192, 0, 0), IIf(d Mod 2 = 1, C_NAVY, RGB(47, 117, 181))), True
        Next d
        MergeTitle ws, r, 19, 22, "Monday vs Tue-Thu", RGB(192, 0, 0), True
        MergeTitle ws, r, 23, 24, "Dialler attempts / contact", C_NAVY, True
        r = r + 1
        hdr = Array("Source", "Summary row")
        ws.Cells(r, 1).Value = "Source"
        ws.Cells(r, 2).Value = "Summary row"
        For d = 0 To 3
            c = 3 + d * 4
            ws.Cells(r, c).Value = "Processed / day"
            ws.Cells(r, c + 1).Value = "Conversion"
            ws.Cells(r, c + 2).Value = "Closing"
            ws.Cells(r, c + 3).Value = "Effectiveness"
        Next d
        ws.Cells(r, 19).Value = "Conversion gap"
        ws.Cells(r, 20).Value = "Closing gap"
        ws.Cells(r, 21).Value = "Effectiveness gap"
        ws.Cells(r, 22).Value = "Sales lost / Monday"
        ws.Cells(r, 23).Value = "Monday"
        ws.Cells(r, 24).Value = "Tue-Thu"
        HeaderStyle ws.Range(ws.Cells(r, 1), ws.Cells(r, lastC)), C_BLUE_LT, False
        ws.Rows(r).RowHeight = 26
        first = r + 1
        For i = 1 To smNB
            If smBBrand(i) = brands(bi) Then
                r = r + 1
                pr = smBProc(i)
                ld = smBLeads(i)
                sl = smBSales(i)
                ws.Cells(r, 1).Value = SafeText(smBSrc(i))
                ws.Cells(r, 2).Value = smBH(i)
                For d = 0 To 3
                    c = 3 + d * 4
                    fP = DaySum(pr, CStr(days(d)))
                    fL = DaySum(ld, CStr(days(d)))
                    fS = DaySum(sl, CStr(days(d)))
                    cnt = DayCount(CStr(days(d)))
                    ws.Cells(r, c).Formula = "=IF(" & cnt & "=0,""""," & fP & "/" & cnt & ")"
                    ws.Cells(r, c + 1).Formula = "=IF(" & fP & "=0,""""," & fL & "/" & fP & ")"
                    ws.Cells(r, c + 2).Formula = "=IF(" & fL & "=0,""""," & fS & "/" & fL & ")"
                    ws.Cells(r, c + 3).Formula = "=IF(" & fP & "=0,""""," & fS & "/" & fP & ")"
                Next d
                ws.Cells(r, 19).Formula = "=IF(AND(ISNUMBER(D" & r & "),ISNUMBER(H" & r & ")),(D" & r & "-H" & r & ")*100,"""")"
                ws.Cells(r, 20).Formula = "=IF(AND(ISNUMBER(E" & r & "),ISNUMBER(I" & r & ")),(E" & r & "-I" & r & ")*100,"""")"
                ws.Cells(r, 21).Formula = "=IF(AND(ISNUMBER(F" & r & "),ISNUMBER(J" & r & ")),(F" & r & "-J" & r & ")*100,"""")"
                ws.Cells(r, 22).Formula = "=IF(AND(ISNUMBER(C" & r & "),ISNUMBER(F" & r & "),ISNUMBER(J" & r & ")),C" & r & "*(J" & r & "-F" & r & "),"""")"
                If smBAtt(i) > 0 Then
                    ws.Cells(r, 23).Formula = "=IF(" & DaySum(pr, "Mon") & "=0,""""," & DaySum(smBAtt(i), "Mon") & "/" & DaySum(pr, "Mon") & ")"
                    ws.Cells(r, 24).Formula = "=IF(" & DaySum(pr, "TueThu") & "=0,""""," & DaySum(smBAtt(i), "TueThu") & "/" & DaySum(pr, "TueThu") & ")"
                End If
                If smBSrc(i) Like "Dialler - Total*" Or smBSrc(i) Like "JHB*" Or smBSrc(i) Like "KZN*" Then
                    ws.Range(ws.Cells(r, 1), ws.Cells(r, lastC)).Font.Bold = True
                End If
            End If
        Next i
        For d = 0 To 3
            c = 3 + d * 4
            ws.Range(ws.Cells(first, c), ws.Cells(r, c)).NumberFormat = "#,##0.0"
            ws.Range(ws.Cells(first, c + 1), ws.Cells(r, c + 3)).NumberFormat = NF_PCT
        Next d
        ws.Range(ws.Cells(first, 19), ws.Cells(r, 21)).NumberFormat = NF_PP
        ws.Range(ws.Cells(first, 22), ws.Cells(r, 22)).NumberFormat = "[Red]#,##0.0;[Color10]-#,##0.0;0.0"
        ws.Range(ws.Cells(first, 23), ws.Cells(r, 24)).NumberFormat = "0.0"
        ws.Range(ws.Cells(first, 3), ws.Cells(r, 6)).Interior.Color = RGB(252, 228, 214)
        ws.Range(ws.Cells(first, 2), ws.Cells(r, 2)).Font.Color = RGB(128, 128, 128)
        Box ws.Range(ws.Cells(first - 2, 1), ws.Cells(r, lastC))
        r = r + 2
NextBrand:
    Next bi
    ws.Columns(1).ColumnWidth = 34
    ws.Columns(2).ColumnWidth = 8
    ws.Range(ws.Cells(1, 3), ws.Cells(1, lastC)).ColumnWidth = 10.5
    Freeze ws, 4, 3
    SetTab ws, RGB(192, 0, 0)
End Sub

'==========================================================================================
'  WEIGHTED CONVERSION & CLOSING (built from the header map - no fixed columns)
'==========================================================================================
Private Sub BuildWeighted()
    Dim ws As Worksheet, r As Long, s As Long, i As Long, detail As Long, top As Long, k As Long, bb As Long
    Dim brands As Variant, defs As Variant, dTot As Long, dRep As Long, t1 As Long, t2 As Long
    s = 1
    Set ws = NewSheet(OUT_W)
    NoGrid
    ws.Columns(1).ColumnWidth = 34
    ws.Columns(2).ColumnWidth = 8
    For i = 3 To 10
        ws.Columns(i).ColumnWidth = 14
    Next i
    ws.Columns(11).ColumnWidth = 52
    MergeTitle ws, 1, 1, 11, "How Conversion %, Gross Closing %, Targets and Weighted (W) Targets are built  -  source '" & gName(s) & "' (" & gPeriod(s) & ")", C_NAVY, True
    ws.Rows(1).RowHeight = 22
    ws.Cells(1, 1).Font.Size = 12
    r = 3
    MergeTitle ws, r, 1, 2, "Measure", C_NAVY, True
    MergeTitle ws, r, 3, 7, "Formula (per line i)", C_NAVY, True
    MergeTitle ws, r, 8, 11, "Notes", C_NAVY, True
    r = r + 1
    defs = Array( _
        Array("Conversion %", "Eligible leads / eligible contacts", "Measured only on lines where a contact precedes the lead (Traditional, Competition, OQ-PNV, DM, Affiliates, Hippo-Branded, DM Outbound)."), _
        Array("Conversion Target %", "Target leads / target contacts", "Plan rate at the PLAN channel mix."), _
        Array("W Target Conversion %", "SUM(actual contacts_i x target conv_i) / SUM(actual contacts_i)", "Plan rates at the ACTUAL channel mix; built from the lowest line upward. A line with no target counts as 0%."), _
        Array("Gross Closing %", "Gross sales / leads", "All leads."), _
        Array("Closing Target %", "Target gross sales / target leads", "Plan closing at the plan lead mix."), _
        Array("W Target Closing %", "SUM(actual leads_i x target closing_i) / SUM(actual leads_i)", "Plan closing at the actual lead mix."), _
        Array("Mix effect", "W Target - Target", "Change in the expected rate caused only by the channel mix moving."), _
        Array("Rate effect", "Actual - W Target", "Genuine performance. Mix + Rate = Actual - Target."), _
        Array("Tar Var / W Tar Var", "Actual / Target - 1  ;  Actual / W Target - 1", "Relative variance, not percentage points."))
    For i = 0 To UBound(defs)
        ws.Cells(r, 1).Value = defs(i)(0)
        ws.Cells(r, 3).Value = defs(i)(1)
        ws.Cells(r, 8).Value = defs(i)(2)
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Merge
        ws.Range(ws.Cells(r, 3), ws.Cells(r, 7)).Merge
        ws.Range(ws.Cells(r, 8), ws.Cells(r, 11)).Merge
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 11)).WrapText = True
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 11)).VerticalAlignment = -4160
        ws.Cells(r, 1).Font.Bold = True
        ws.Rows(r).RowHeight = 28
        r = r + 1
    Next i
    Box ws.Range(ws.Cells(3, 1), ws.Cells(r - 1, 11))
    r = r + 1
    detail = -1
    top = -1
    For i = 1 To gRows(s)
        If detail < 0 And gLab(s, i) = "Generic - Traditional" And gBB(s, i) = "" Then detail = gBlk(s, i)
        If top < 0 And gLab(s, i) = "Non-Direct" And gBB(s, i) = "" And gSB(s, i) = "" Then top = gBlk(s, i)
    Next i
    If detail < 0 Then detail = top
    WExample ws, r, s, detail, "A. Direct - Conversion W Target (contacts-weighted)", _
        "Conversion-eligible children of Direct: OQ uses only OQ-PNV and Hippo only Hippo-Branded.", G_CONTACTS, G_CONV, _
        Array("Traditional Combined", "Competitions Combined", "OQ - PNV", "Direct Marketing", "Affiliates", "Hippo - Branded", "DM Multibrand Outbound"), "Direct", t1, t2
    WExample ws, r, s, detail, "B. Direct - Gross Closing W Target (leads-weighted)", _
        "Every lead counts for closing. DM Multibrand Outbound already contains Referral.", G_LEADS, G_CLOSE, _
        Array("Traditional Combined", "Competitions Combined", "OQ", "Direct Marketing", "Affiliates", "Hippo", "Better Compare", "Cover Compare", "Direct Other Clean Up", "DM Multibrand Outbound"), "Direct", t1, t2
    WExample ws, r, s, detail, "C. OQ - Gross Closing W Target, one level down", _
        "OQ's W is the lead-weighted average of its three sub-lines.", G_LEADS, G_CLOSE, _
        Array("OQ - PNV", "OQ - PV", "OQ - End to End"), "OQ", t1, t2
    WExample ws, r, s, top, "D. Non-Direct - Conversion W Target", _
        "If children carry W = 0 although a plan target exists, the parent W collapses and W Tar Var explodes - a calculation error, not performance.", G_CONTACTS, G_CONV, _
        Array("Affinities", "Bank", "Dealer Channel", "Broker", "Property Partners", "Other Cleanup"), "Non-Direct", dTot, dRep
    brands = Array("Budget", "First For Women", "Dialdirect", "Auto & General", "Virseker")
    For k = 0 To 4
        bb = -1
        For i = 1 To gRows(s)
            If gBB(s, i) = brands(k) And gLab(s, i) = "Direct (Excl Outbound)" Then
                bb = gBlk(s, i)
                Exit For
            End If
        Next i
        If bb >= 0 Then
            WExample ws, r, s, bb, Chr$(69 + k) & ". " & brands(k) & " - Direct (Excl Outbound) Conversion W Target", _
                "Same mechanics inside the " & brands(k) & " brand page.", G_CONTACTS, G_CONV, _
                Array("Traditional Combined", "Competitions Combined", "OQ - PNV", "Direct Marketing", "Affiliates", "Hippo - Branded"), "Direct (Excl Outbound)", t1, t2
        End If
    Next k
    Freeze ws, 3, 1
    SetTab ws, RGB(112, 48, 160)
End Sub

Private Sub WExample(ByVal ws As Worksheet, ByRef r As Long, ByVal s As Long, ByVal blk As Long, ByVal title As String, ByVal note As String, _
                     ByVal gVol As Long, ByVal gRate As Long, kids As Variant, ByVal parent As String, ByRef totRow As Long, ByRef repRow As Long)
    Dim src As String, i As Long, kr As Long, prow As Long, first As Long, last As Long, tot As Long, rep As Long, j As Long
    Dim hdr As Variant, labels As Variant, fm As Variant, cV As Long, cR As Long, cT As Long, cW As Long
    Dim volName As String, kind As String
    src = QS(gName(s))
    cV = gCol(s, gVol, M_ACT)
    cR = gCol(s, gRate, M_ACT)
    cT = gCol(s, gRate, M_TAR)
    cW = gCol(s, gRate, M_WT)
    volName = gGroupName(s, gVol)
    kind = gGroupName(s, gRate)
    MergeTitle ws, r, 1, 11, title, C_NAVY, True
    ws.Cells(r, 1).HorizontalAlignment = -4131
    r = r + 1
    ws.Cells(r, 1).Value = note
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 11)).Merge
    ws.Cells(r, 1).WrapText = True
    ws.Cells(r, 1).Font.Italic = True
    ws.Rows(r).RowHeight = 26
    r = r + 1
    If cV = 0 Or cR = 0 Or cT = 0 Or cW = 0 Then
        ws.Cells(r, 1).Value = "Columns not found in '" & gName(s) & "' (needs " & volName & " Actual, " & kind & " Actual / Target / W Target)."
        ws.Cells(r, 1).Font.Color = RGB(192, 0, 0)
        r = r + 2
        Exit Sub
    End If
    prow = FindInBlock(s, blk, parent, 0)
    hdr = Array("Line (child)", "Src row", "Actual " & volName, "Actual " & kind, "Implied numerator", "Target " & kind, _
                "W Target " & kind, "Volume x W Target", "Volume x Target", "Mix share", "Comment")
    For j = 0 To 10
        ws.Cells(r, j + 1).Value = hdr(j)
    Next j
    HeaderStyle ws.Range(ws.Cells(r, 1), ws.Cells(r, 11)), C_BLUE_LT, False
    ws.Rows(r).RowHeight = 26
    first = r + 1
    last = r + UBound(kids) + 1
    For i = 0 To UBound(kids)
        r = r + 1
        If prow > 0 Then kr = FindInBlock(s, blk, CStr(kids(i)), 0) Else kr = 0
        ws.Cells(r, 1).Value = kids(i)
        If kr > 0 Then
            ws.Cells(r, 2).Value = kr
            ws.Cells(r, 3).Formula = "=N(" & src & ColL(cV) & kr & ")"
            ws.Cells(r, 4).Formula = "=N(" & src & ColL(cR) & kr & ")"
            ws.Cells(r, 6).Formula = "=N(" & src & ColL(cT) & kr & ")"
            ws.Cells(r, 7).Formula = "=N(" & src & ColL(cW) & kr & ")"
        Else
            ws.Cells(r, 2).Value = "n/a"
            ws.Cells(r, 3).Value = 0
            ws.Cells(r, 4).Value = 0
            ws.Cells(r, 6).Value = 0
            ws.Cells(r, 7).Value = 0
            ws.Cells(r, 11).Value = "Line not on this page"
        End If
        ws.Cells(r, 5).Formula = "=C" & r & "*D" & r
        ws.Cells(r, 8).Formula = "=C" & r & "*G" & r
        ws.Cells(r, 9).Formula = "=C" & r & "*F" & r
        ws.Cells(r, 10).Formula = "=IF(SUM($C$" & first & ":$C$" & last & ")=0,0,C" & r & "/SUM($C$" & first & ":$C$" & last & "))"
    Next i
    r = r + 1
    tot = r
    ws.Cells(r, 1).Value = "Total / re-derived"
    ws.Cells(r, 3).Formula = "=SUM(C" & first & ":C" & last & ")"
    ws.Cells(r, 4).Formula = "=IF(C" & r & "=0,0,E" & r & "/C" & r & ")"
    ws.Cells(r, 5).Formula = "=SUM(E" & first & ":E" & last & ")"
    ws.Cells(r, 6).Formula = "=IF(C" & r & "=0,0,I" & r & "/C" & r & ")"
    ws.Cells(r, 7).Formula = "=IF(C" & r & "=0,0,H" & r & "/C" & r & ")"
    ws.Cells(r, 8).Formula = "=SUM(H" & first & ":H" & last & ")"
    ws.Cells(r, 9).Formula = "=SUM(I" & first & ":I" & last & ")"
    ws.Cells(r, 10).Formula = "=SUM(J" & first & ":J" & last & ")"
    ws.Cells(r, 11).Value = "G = SUMPRODUCT(volume, W) / SUM(volume) = re-derived W Target"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 11)).Font.Bold = True
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 11)).Interior.Color = C_GREY_LT
    r = r + 1
    rep = r
    ws.Cells(r, 1).Value = "Reported on the Exec sheet"
    If prow > 0 Then
        ws.Cells(r, 2).Value = prow
        ws.Cells(r, 3).Formula = "=N(" & src & ColL(cV) & prow & ")"
        ws.Cells(r, 4).Formula = "=N(" & src & ColL(cR) & prow & ")"
        ws.Cells(r, 6).Formula = "=N(" & src & ColL(cT) & prow & ")"
        ws.Cells(r, 7).Formula = "=N(" & src & ColL(cW) & prow & ")"
    End If
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 11)).Font.Bold = True
    r = r + 1
    ws.Cells(r, 1).Value = "Difference (re-derived - reported)"
    ws.Cells(r, 4).Formula = "=D" & tot & "-D" & rep
    ws.Cells(r, 7).Formula = "=G" & tot & "-G" & rep
    ws.Cells(r, 8).Formula = "=IF(ABS(G" & tot & "-G" & rep & ")<0.0005,""Reconciles"",""Does NOT reconcile"")"
    ws.Range(ws.Cells(r, 4), ws.Cells(r, 7)).NumberFormat = NF_VAR
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 8)).Font.Bold = True
    r = r + 1
    labels = Array("Actual - Target (pp)", "Mix effect = W Target - Target (pp)", "Rate effect = Actual - W Target (pp)", "Tar Var = Actual / Target - 1", "W Tar Var = Actual / W Target - 1")
    fm = Array("=(D" & rep & "-F" & rep & ")*100", "=(G" & rep & "-F" & rep & ")*100", "=(D" & rep & "-G" & rep & ")*100", _
               "=IF(F" & rep & "=0,"""",D" & rep & "/F" & rep & "-1)", "=IF(G" & rep & "=0,"""",D" & rep & "/G" & rep & "-1)")
    For i = 0 To 4
        ws.Cells(r, 1).Value = labels(i)
        ws.Cells(r, 4).Formula = fm(i)
        If i < 3 Then ws.Cells(r, 4).NumberFormat = NF_PP Else ws.Cells(r, 4).NumberFormat = NF_VAR
        ws.Cells(r, 4).Font.Bold = True
        r = r + 1
    Next i
    ws.Range(ws.Cells(first, 3), ws.Cells(rep, 3)).NumberFormat = NF_INT
    ws.Range(ws.Cells(first, 5), ws.Cells(rep, 5)).NumberFormat = NF_INT
    ws.Range(ws.Cells(first, 8), ws.Cells(rep, 9)).NumberFormat = NF_INT
    ws.Range(ws.Cells(first, 4), ws.Cells(rep, 4)).NumberFormat = "0.00%"
    ws.Range(ws.Cells(first, 6), ws.Cells(rep, 7)).NumberFormat = "0.00%"
    ws.Range(ws.Cells(first, 10), ws.Cells(rep, 10)).NumberFormat = NF_PCT
    Box ws.Range(ws.Cells(first - 1, 1), ws.Cells(rep, 11))
    totRow = tot
    repRow = rep
    r = r + 1
End Sub

'==========================================================================================
'  REPORT CHECKS (control sheet)
'==========================================================================================
Private Sub BuildChecks()
    Dim ws As Worksheet, r As Long, s As Long, rep As Worksheet, b As Variant, k As Long, i As Long, rr As Long
    Dim brandRow As Long, dRow As Long, nRow As Long, lastC As Long, cC As Long, cL As Long, cS As Long, v As Variant
    Dim cnt As Long, listed As Long
    Set ws = NewSheet(OUT_CHECKS)
    NoGrid
    lastC = 9
    MergeTitle ws, 1, 1, lastC, "Report Checks - does every number reconcile?", C_NAVY, True
    ws.Rows(1).RowHeight = 24
    ws.Cells(1, 1).Font.Size = 13
    ws.Cells(2, 1).Value = "Run " & Format(Now, "dd mmm yyyy") & " " & Format(Now, "hh:mm") & ". Columns, rows, months and weekdays are located by their header text and labels, " & _
                           "so the checks stay correct when the layout or the numbers change. Green 0 = reconciles; red = does not."
    ws.Range(ws.Cells(2, 1), ws.Cells(2, lastC)).Merge
    ws.Cells(2, 1).WrapText = True
    ws.Cells(2, 1).Font.Italic = True
    ws.Rows(2).RowHeight = 28
    r = 4
    gStep = "building Report Checks section 1"
    ' ---- 1. what was found ---------------------------------------------------------------
    MergeTitle ws, r, 1, lastC, "1. Sheets found and how they were read", C_NAVY, True
    ws.Cells(r, 1).HorizontalAlignment = -4131
    r = r + 1
    ws.Cells(r, 1).Value = "Role"
    ws.Cells(r, 2).Value = "Sheet"
    ws.Cells(r, 3).Value = "Period"
    ws.Cells(r, 4).Value = "Header row"
    ws.Cells(r, 5).Value = "Label column"
    ws.Cells(r, 6).Value = "Columns found by header"
    ws.Range(ws.Cells(r, 6), ws.Cells(r, lastC)).Merge
    HeaderStyle ws.Range(ws.Cells(r, 1), ws.Cells(r, lastC)), C_BLUE_LT, False
    For s = 1 To 3
        If gHave(s) Then
            r = r + 1
            ws.Cells(r, 1).Value = Choose(s, "Current month", "Prior month", "Year to date")
            ws.Cells(r, 2).Value = gName(s)
            ws.Cells(r, 3).Value = "'" & gPeriod(s)
            ws.Cells(r, 4).Value = gHdrRow(s)
            ws.Cells(r, 5).Value = ColL(gLabCol(s))
            ws.Cells(r, 6).Value = ColumnMapText(s)
            ws.Range(ws.Cells(r, 6), ws.Cells(r, lastC)).Merge
            ws.Cells(r, 6).WrapText = True
            ws.Rows(r).RowHeight = 40
        End If
    Next s
    If smHave Then
        r = r + 1
        ws.Cells(r, 1).Value = "Summary (sources x brands)"
        ws.Cells(r, 2).Value = smName
        ws.Cells(r, 3).Value = "'" & smCurLab & " vs " & smPrvLab
        ws.Cells(r, 4).Value = smPerRow
        ws.Cells(r, 5).Value = "C"
        ws.Cells(r, 6).Value = smNB & " source x brand blocks; months " & ColL(smPrvCol) & "/" & ColL(smCurCol) & "; daily columns " & ColL(smD1) & ":" & ColL(smD2)
        ws.Range(ws.Cells(r, 6), ws.Cells(r, lastC)).Merge
    End If
    Box ws.Range(ws.Cells(5, 1), ws.Cells(r, lastC))
    r = r + 2
    gStep = "building Report Checks section 2"
    ' ---- 2. recalculation ---------------------------------------------------------------
    MergeTitle ws, r, 1, lastC, "2. Recalculated values vs what the sheet reports (Tar Dif, Tar Var, Closing, CXD, W Tar Var)", C_NAVY, True
    ws.Cells(r, 1).HorizontalAlignment = -4131
    r = r + 1
    ws.Cells(r, 1).Value = "Replica"
    ws.Cells(r, 2).Value = "Cells that differ"
    ws.Cells(r, 3).Value = "Parent lines that do not equal their child lines"
    ws.Range(ws.Cells(r, 3), ws.Cells(r, 4)).Merge
    ws.Cells(r, 5).Value = "Go to"
    HeaderStyle ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)), C_BLUE_LT, False
    For s = 1 To 3
        If gHave(s) Then
            r = r + 1
            ws.Cells(r, 1).Value = gRepName(s)
            ws.Cells(r, 2).Formula = "=SUM(" & QS(gRepName(s)) & ColL(gRepLast(s) - 7) & ":" & ColL(gRepLast(s) - 7) & ")"
            ws.Cells(r, 3).Value = gRepHier(s)
            ws.Range(ws.Cells(r, 3), ws.Cells(r, 4)).Merge
            ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).NumberFormat = NF_CHK
            LinkCell ws.Cells(r, 5), QS(gRepName(s)) & "A1"
        End If
    Next s
    Box ws.Range(ws.Cells(r - 3, 1), ws.Cells(r, 5))
    r = r + 1
    ' lines with recalculation differences in the current-month replica
    Dim repWs As Worksheet, mv As Variant, cm As Long, nList As Long
    Set repWs = ThisWorkbook.Worksheets(gRepName(1))
    cm = gRepLast(1) - 7
    mv = repWs.Range(repWs.Cells(1, 1), repWs.Cells(gRows(1) + 2, cm)).Value
    nList = 0
    For i = gHdrRow(1) + 2 To gRows(1)
        If IsNum(mv(i + 2, cm)) Then
            If mv(i + 2, cm) > 0 And nList < 60 Then
                r = r + 1
                nList = nList + 1
                ws.Cells(r, 1).Value = "  " & SafeText(gLab(1, i))
                ws.Cells(r, 2).Value = mv(i + 2, cm) & " cell(s) differ - row " & i
                ws.Cells(r, 3).Value = IIf(NoVolumes(1, i), _
                                           "rates shown without volumes - cannot be recalculated", "see orange cells")
                ws.Range(ws.Cells(r, 3), ws.Cells(r, 4)).Merge
                LinkCell ws.Cells(r, 5), QS(gRepName(1)) & "C" & (i + 2)
            End If
        End If
    Next i
    r = r + 2
    gStep = "building Report Checks section 3"
    ' ---- 3. brand totals ----------------------------------------------------------------
    MergeTitle ws, r, 1, lastC, "3. Brand Total = Direct + Non-Direct  (current month, from the replica - live)", C_NAVY, True
    ws.Cells(r, 1).HorizontalAlignment = -4131
    r = r + 1
    v = Array("Brand", "Measure", "Brand Total", "Direct", "Non-Direct", "Direct + Non-Direct", "Gap")
    For k = 0 To 6
        ws.Cells(r, k + 1).Value = v(k)
    Next k
    HeaderStyle ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)), C_BLUE_LT, False
    s = 1
    For Each b In Array("Budget", "First For Women", "Dialdirect", "Auto & General", "Virseker")
        brandRow = 0
        For i = 1 To gRows(s)
            If gLab(s, i) = b & " - Brand Total" Then
                brandRow = i
                Exit For
            End If
        Next i
        If brandRow > 0 Then
            dRow = FindInBlock(s, gBlk(s, brandRow), "Direct", brandRow)
            nRow = FindInBlock(s, gBlk(s, brandRow), "Non-Direct", brandRow)
            For k = 1 To 3
                cC = gRepCol(s, Choose(k, G_CONTACTS, G_LEADS, G_GSALES), M_ACT)
                If cC > 0 And dRow > 0 And nRow > 0 Then
                    r = r + 1
                    ws.Cells(r, 1).Value = b
                    ws.Cells(r, 2).Value = Choose(k, "Contacts", "Leads", "Gross Sales")
                    ws.Cells(r, 3).Formula = "=N(" & QS(gRepName(s)) & ColL(cC) & (brandRow + 2) & ")"
                    ws.Cells(r, 4).Formula = "=N(" & QS(gRepName(s)) & ColL(cC) & (dRow + 2) & ")"
                    ws.Cells(r, 5).Formula = "=N(" & QS(gRepName(s)) & ColL(cC) & (nRow + 2) & ")"
                    ws.Cells(r, 6).Formula = "=D" & r & "+E" & r
                    ws.Cells(r, 7).Formula = "=C" & r & "-F" & r
                    ws.Range(ws.Cells(r, 3), ws.Cells(r, 6)).NumberFormat = NF_INT
                    ws.Cells(r, 7).NumberFormat = NF_CHK
                End If
            Next k
        End If
    Next b
    Box ws.Range(ws.Cells(r - 14, 1), ws.Cells(r, 7))
    r = r + 2
    gStep = "building Report Checks section 4"
    ' ---- 4. parent / child differences list ---------------------------------------------
    MergeTitle ws, r, 1, lastC, "4. Lines whose children do not add up (current month) - click to open the replica row", C_NAVY, True
    ws.Cells(r, 1).HorizontalAlignment = -4131
    r = r + 1
    v = Array("Line", "Block / brand", "Contacts diff", "Leads diff", "Gross Sales diff", "Go to")
    For k = 0 To 5
        ws.Cells(r, k + 1).Value = v(k)
    Next k
    HeaderStyle ws.Range(ws.Cells(r, 1), ws.Cells(r, 6)), C_BLUE_LT, False
    listed = 0
    Set rep = ThisWorkbook.Worksheets(gRepName(1))
    Dim vv As Variant
    vv = rep.Range(rep.Cells(1, 1), rep.Cells(gRows(1) + 2, gRepLast(1))).Value
    For i = gHdrRow(1) + 2 To gRows(1)
        rr = i + 2
        cnt = 0
        For k = gRepLast(1) - 4 To gRepLast(1) Step 2
            If IsNum(vv(rr, k)) Then
                If Abs(vv(rr, k)) > 0.5 Then cnt = cnt + 1
            End If
        Next k
        If cnt > 0 And listed < 150 Then
            r = r + 1
            listed = listed + 1
            ws.Cells(r, 1).Value = SafeText(gLab(1, i))
            ws.Cells(r, 2).Value = IIf(gBB(1, i) <> "", gBB(1, i), "All brands") & " (row " & i & ")"
            ws.Cells(r, 3).Formula = "=" & QS(gRepName(1)) & ColL(gRepLast(1) - 4) & rr
            ws.Cells(r, 4).Formula = "=" & QS(gRepName(1)) & ColL(gRepLast(1) - 2) & rr
            ws.Cells(r, 5).Formula = "=" & QS(gRepName(1)) & ColL(gRepLast(1)) & rr
            ws.Range(ws.Cells(r, 3), ws.Cells(r, 5)).NumberFormat = NF_CHK
            LinkCell ws.Cells(r, 6), QS(gRepName(1)) & "C" & rr
        End If
    Next i
    If listed = 0 Then
        r = r + 1
        ws.Cells(r, 1).Value = "All parent lines equal the sum of their child lines."
        ws.Cells(r, 1).Font.Color = C_GREEN
    End If
    Box ws.Range(ws.Cells(r - listed - IIf(listed = 0, 1, 0), 1), ws.Cells(r, 6))
    r = r + 2
    gStep = "building Report Checks section 5"
    ' ---- 5. summary checks ----------------------------------------------------------------
    If smHave Then
        MergeTitle ws, r, 1, lastC, "5. Summary tab: recalculated conversion (Leads / Processed) vs reported", C_NAVY, True
        ws.Cells(r, 1).HorizontalAlignment = -4131
        r = r + 1
        ws.Cells(r, 1).Value = "Blocks with a difference > 0.05 pp"
        ws.Cells(r, 2).Formula = "=COUNTIF(" & QS(OUT_BXS) & "U1:U3000,"">0.05"")+COUNTIF(" & QS(OUT_BXS) & "U1:U3000,""<-0.05"")"
        ws.Cells(r, 2).NumberFormat = NF_CHK
        LinkCell ws.Cells(r, 3), QS(OUT_BXS) & "A1"
        r = r + 2
    End If
    gStep = "building Report Checks section 6"
    ' ---- 6. legend ------------------------------------------------------------------------
    MergeTitle ws, r, 1, lastC, "6. How to read the report", C_NAVY, True
    ws.Cells(r, 1).HorizontalAlignment = -4131
    v = Array("Replica sheets: same rows and labels as your Exec sheet. Volumes are links; columns marked * are recalculated by formula.", _
              "Orange cell in a replica = recalculated value differs from your sheet. The 'Recalc differences' column counts them per line.", _
              "Parent vs children: the sum of the child lines (e.g. Traditional = Generic + E-Commerce) and the difference to the parent line.", _
              "Projecting is shown only in the Projecting column - it is never used in place of Actual.", _
              "Brand x Source / Monday View: rates are ratios of totals (sum of leads / sum of processed), never averages of daily percentages.", _
              "Colours: green = positive, red = negative, black = other. For CXD variances the colours are inverted (more cancellations = red).")
    For k = 0 To UBound(v)
        r = r + 1
        ws.Cells(r, 1).Value = v(k)
        ws.Range(ws.Cells(r, 1), ws.Cells(r, lastC)).Merge
        ws.Cells(r, 1).WrapText = True
        ws.Rows(r).RowHeight = 15
    Next k
    ws.Columns(1).ColumnWidth = 30
    ws.Columns(2).ColumnWidth = 22
    ws.Range("C1:I1").ColumnWidth = 14
    SetTab ws, RGB(192, 0, 0)
End Sub

Private Sub LinkCell(ByVal cell As Range, ByVal target As String)
    ' target like 'Sheet'!C12  ->  =HYPERLINK("#'Sheet'!C12","open")
    On Error Resume Next
    cell.Formula = "=HYPERLINK(""#" & Replace(target, """", """""") & """,""open"")"
    cell.Font.Color = RGB(0, 102, 204)
    cell.Font.Underline = 2
End Sub

Private Function ColumnMapText(ByVal s As Long) As String
    Dim g As Long, m As Long, t As String, part As String
    t = ""
    For g = 1 To 9
        If gGroupName(s, g) <> "" Then
            part = gGroupName(s, g) & ": "
            For m = 1 To 7
                If gCol(s, g, m) > 0 Then part = part & MeasureLabel(m) & "=" & ColL(gCol(s, g, m)) & " "
            Next m
            t = t & IIf(t = "", "", " | ") & Trim$(part)
        End If
    Next g
    ColumnMapText = t
End Function
