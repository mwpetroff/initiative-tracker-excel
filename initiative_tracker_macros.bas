Attribute VB_Name = "InitiativeTracker"
Option Explicit

' ==============================================================================
'  Initiative Tracker VBA — CRUD via REST API
'  Works on: Windows Excel 2016+ and Mac Excel 2016+
'
'  Windows : MSXML2.XMLHTTP60 (built-in COM)
'  Mac     : curl via MacScript/AppleScript (curl is pre-installed on macOS)
'  JSON    : pure-VBA parser — no COM / no VBScript.RegExp — works on both
'
'  SETUP
'    1. Enable macros when prompted (or: File > Options > Trust Center >
'       Macro Settings > Enable all macros)
'    2. Set your API URL in the Config sheet cell B4
'    3. Run TestConnection to verify, then run CreateButtons once to set up
'       the sidebar buttons and the two built-in form sheets
'
'  USAGE
'    Sidebar buttons on the Initiatives sheet open styled form sheets.
'    Fill in the form and click Save — the sheet updates immediately.
'    No multi-step dialog chains needed.
' ==============================================================================

' ─── Config ───────────────────────────────────────────────────────────────────
Private Function ApiBase() As String
    On Error Resume Next
    ApiBase = Trim(ThisWorkbook.Sheets("Config").Range("B4").Value)
    If ApiBase = "" Then ApiBase = "http://localhost/api"
    On Error GoTo 0
End Function

' ─── HTTP — platform switch ───────────────────────────────────────────────────
Private Function Http(method As String, path As String, _
                      Optional body As String = "") As String
#If Mac Then
    Http = HttpMac(method, ApiBase() & path, body)
#Else
    Http = HttpWin(method, ApiBase() & path, body)
#End If
End Function

' Windows: MSXML2.XMLHTTP60
#If Not Mac Then
Private Function HttpWin(method As String, url As String, body As String) As String
    Dim x As Object
    Set x = CreateObject("MSXML2.XMLHTTP60")
    x.Open method, url, False
    x.setRequestHeader "Content-Type", "application/json"
    x.setRequestHeader "Accept", "application/json"
    If body <> "" Then x.Send body Else x.Send
    If x.Status >= 200 And x.Status < 300 Then
        HttpWin = x.responseText
    Else
        Err.Raise vbObjectError + 500, , "HTTP " & x.Status & Chr(10) & x.responseText
    End If
End Function
#End If

' Mac: curl via MacScript (synchronous, returns stdout)
' Uses a unique sentinel string to separate the response body from the
' HTTP status code that curl appends via -w.
' Single-quoted curl args avoid JSON double-quote escaping entirely.
#If Mac Then
Private Function HttpMac(method As String, url As String, body As String) As String
    Const SEP As String = "|||HTTPSTATUS|||"
    Dim q As String: q = Chr(39)   ' single-quote char

    Dim cmd As String
    cmd = "curl -s -w " & q & SEP & "%{http_code}" & q
    cmd = cmd & " -X " & method
    cmd = cmd & " -H " & q & "Content-Type: application/json" & q
    cmd = cmd & " -H " & q & "Accept: application/json" & q

    If body <> "" Then
        ' Escape any single-quotes inside the body for shell single-quoting:
        ' replace each ' with '\''
        Dim sb As String
        sb = Replace(body, q, q & Chr(92) & q & q)
        cmd = cmd & " -d " & q & sb & q
    End If
    cmd = cmd & " " & q & url & q

    ' Wrap in AppleScript: only need to escape " as \" since we used ' for everything else
    Dim asCmd As String
    asCmd = Replace(cmd, Chr(34), Chr(92) & Chr(34))

    Dim raw As String
    On Error GoTo MacFail
    raw = MacScript("do shell script """ & asCmd & """")
    On Error GoTo 0

    Dim sepPos As Long: sepPos = InStr(raw, SEP)
    Dim status As String, resp As String
    If sepPos > 0 Then
        status = Trim(Mid(raw, sepPos + Len(SEP)))
        resp   = Left(raw, sepPos - 1)
    Else
        status = "0": resp = raw
    End If

    If Val(status) >= 200 And Val(status) < 300 Then
        HttpMac = resp
    Else
        Err.Raise vbObjectError + 500, , "HTTP " & status & Chr(10) & resp
    End If
    Exit Function
MacFail:
    Err.Raise vbObjectError + 500, , _
        "MacScript failed. If you are on macOS Ventura or later, " & _
        "go to System Settings > Privacy & Security > Automation and " & _
        "allow Excel to control the shell." & Chr(10) & Chr(10) & Err.Description
End Function
#End If

' ─── JSON — pure VBA, no COM objects (Windows + Mac) ─────────────────────────
' Extracts the value for a given key from a flat/shallow JSON object.
Private Function JGet(json As String, key As String) As String
    Dim srch As String: srch = Chr(34) & key & Chr(34) & ":"
    Dim p As Long: p = InStr(1, json, srch, vbBinaryCompare)
    If p = 0 Then Exit Function
    p = p + Len(srch)
    ' skip whitespace
    Do While p <= Len(json) And (Mid(json, p, 1) = " " Or Mid(json, p, 1) = Chr(9))
        p = p + 1
    Loop
    Dim ch As String: ch = Mid(json, p, 1)
    Select Case ch
        Case Chr(34)                    ' string value
            p = p + 1
            Dim e As Long: e = p
            Do While e <= Len(json)
                If Mid(json, e, 1) = Chr(92) Then   ' backslash — skip next char
                    e = e + 2
                ElseIf Mid(json, e, 1) = Chr(34) Then
                    Exit Do
                Else
                    e = e + 1
                End If
            Loop
            Dim v As String: v = Mid(json, p, e - p)
            ' unescape JSON sequences
            v = Replace(v, Chr(92) & Chr(92), Chr(1))  ' \\ -> placeholder
            v = Replace(v, Chr(92) & Chr(34), Chr(34)) ' \" -> "
            v = Replace(v, Chr(92) & "n",     Chr(10)) ' \n -> newline
            v = Replace(v, Chr(92) & "r",     "")      ' \r -> nothing
            v = Replace(v, Chr(92) & "t",     Chr(9))  ' \t -> tab
            v = Replace(v, Chr(1),            Chr(92)) ' placeholder -> \
            JGet = v
        Case "n": JGet = ""             ' null
        Case "t": JGet = "true"
        Case "f": JGet = "false"
        Case Else                       ' number
            Dim ne As Long: ne = p
            Do While ne <= Len(json)
                ch = Mid(json, ne, 1)
                If InStr(",}] " & Chr(9) & Chr(10), ch) > 0 Then Exit Do
                ne = ne + 1
            Loop
            JGet = Mid(json, p, ne - p)
    End Select
End Function

' Escape a string for embedding in a JSON double-quoted value
Private Function JEsc(s As String) As String
    s = Replace(s, Chr(92), Chr(92) & Chr(92))  ' \ -> \\
    s = Replace(s, Chr(34), Chr(92) & Chr(34))  ' " -> \"
    s = Replace(s, Chr(10), Chr(92) & "n")      ' LF -> \n
    s = Replace(s, Chr(13), "")                 ' CR -> (drop)
    JEsc = s
End Function

' Wrap a string in double quotes
Private Function Q(s As String) As String
    Q = Chr(34) & s & Chr(34)
End Function

' Split a top-level JSON array "[{...},{...}]" into individual object strings.
' Returns an empty Variant array when input is empty or not an array.
Private Function JSplitArray(json As String) As Variant
    json = Trim(json)
    If Len(json) < 2 Then JSplitArray = Array(): Exit Function
    If Left(json, 1) <> "[" Then JSplitArray = Array(): Exit Function

    Dim inner As String: inner = Trim(Mid(json, 2, Len(json) - 2))
    If inner = "" Then JSplitArray = Array(): Exit Function

    Dim items()  As String
    Dim count    As Long:    count    = 0
    Dim depth    As Long:    depth    = 0
    Dim inStr    As Boolean: inStr    = False
    Dim escaped  As Boolean: escaped  = False
    Dim startPos As Long:    startPos = 1
    Dim i        As Long
    Dim ch       As String

    ReDim items(0)
    For i = 1 To Len(inner)
        ch = Mid(inner, i, 1)
        If escaped Then
            escaped = False
        ElseIf inStr Then
            If ch = Chr(92) Then escaped = True
            If ch = Chr(34) Then inStr = False
        Else
            If ch = Chr(34) Then inStr = True
            If ch = "{" Or ch = "[" Then depth = depth + 1
            If ch = "}" Or ch = "]" Then depth = depth - 1
            If ch = "," And depth = 0 Then
                ReDim Preserve items(count)
                items(count) = Trim(Mid(inner, startPos, i - startPos))
                count = count + 1
                startPos = i + 1
            End If
        End If
    Next i

    Dim lastItem As String: lastItem = Trim(Mid(inner, startPos))
    If lastItem <> "" Then
        ReDim Preserve items(count)
        items(count) = lastItem
        count = count + 1
    End If

    If count = 0 Then
        JSplitArray = Array()
    Else
        ReDim Preserve items(count - 1)
        JSplitArray = items
    End If
End Function

' Build the JSON body for a POST/PUT initiative request
Private Function BuildBody(title As String, status As String, priority As String, _
    progress As Long, owner As String, dept As String, _
    desc As String, sd As String, ed As String) As String
    Dim p As String
    p = Q("title")    & ":" & Q(JEsc(title))
    p = p & "," & Q("status")    & ":" & Q(status)
    p = p & "," & Q("priority")  & ":" & Q(priority)
    p = p & "," & Q("progress")  & ":" & CStr(progress)
    If owner <> "" Then p = p & "," & Q("owner")       & ":" & Q(JEsc(owner))
    If dept  <> "" Then p = p & "," & Q("department")  & ":" & Q(JEsc(dept))
    If desc  <> "" Then p = p & "," & Q("description") & ":" & Q(JEsc(desc))
    If sd    <> "" Then p = p & "," & Q("startDate")   & ":" & Q(sd)
    If ed    <> "" Then p = p & "," & Q("endDate")     & ":" & Q(ed)
    BuildBody = "{" & p & "}"
End Function

' ─── Form sheet names & cell addresses ───────────────────────────────────────
' Two hidden sheets act as modal forms. VBA shows/hides them on demand.
Private Const INIT_FORM   As String = "Initiative Form"
Private Const UPDATE_FORM As String = "Update Form"

' Initiative Form — input cells in column B
Private Const IFC_TITLE  As String = "B3"
Private Const IFC_STATUS As String = "B4"
Private Const IFC_PRIO   As String = "B5"
Private Const IFC_PROG   As String = "B6"
Private Const IFC_OWNER  As String = "B7"
Private Const IFC_DEPT   As String = "B8"
Private Const IFC_DESC   As String = "B9"
Private Const IFC_SDATE  As String = "B10"
Private Const IFC_EDATE  As String = "B11"
' State stored off-screen in column G (hidden from users)
Private Const IFC_MODE   As String = "G1"
Private Const IFC_ID     As String = "G2"

' Update Form — input cells
Private Const UFC_NM     As String = "B3"
Private Const UFC_NOTE   As String = "B4"
Private Const UFC_AUTHOR As String = "B6"
Private Const UFC_ID     As String = "G1"

' ─── Row / sheet helpers ──────────────────────────────────────────────────────
Private Function RowId(ws As Worksheet, row As Long) As String
    On Error Resume Next
    RowId = ws.Cells(row, 1).Comment.Text
    On Error GoTo 0
End Function

Private Function InitWS() As Worksheet
    On Error Resume Next
    Set InitWS = ThisWorkbook.Sheets("Initiatives")
    On Error GoTo 0
    If InitWS Is Nothing Then
        MsgBox "Initiatives sheet not found.", vbExclamation
    End If
End Function

Private Function StatusLabel(s As String) As String
    Select Case LCase(Trim(s))
        Case "on_track":    StatusLabel = "On Track"
        Case "at_risk":     StatusLabel = "At Risk"
        Case "delayed":     StatusLabel = "Delayed"
        Case "completed":   StatusLabel = "Completed"
        Case "not_started": StatusLabel = "Not Started"
        Case Else:          StatusLabel = s
    End Select
End Function

Private Function StatusCode(label As String) As String
    Select Case LCase(Trim(label))
        Case "on track":    StatusCode = "on_track"
        Case "at risk":     StatusCode = "at_risk"
        Case "delayed":     StatusCode = "delayed"
        Case "completed":   StatusCode = "completed"
        Case "not started": StatusCode = "not_started"
        Case Else:          StatusCode = LCase(Replace(Trim(label), " ", "_"))
    End Select
End Function

Private Function CurrentUser() As String
#If Mac Then
    CurrentUser = Environ("USER")
#Else
    CurrentUser = Environ("USERNAME")
#End If
    If CurrentUser = "" Then CurrentUser = "User"
End Function

Private Sub WriteInitRow(ws As Worksheet, r As Long, _
    id As String, rowNum As Long, _
    title As String, status As String, priority As String, progress As Long, _
    owner As String, dept As String, sd As String, ed As String, updatedAt As String)
    ws.Cells(r, 1).Value  = rowNum
    ws.Cells(r, 2).Value  = title
    ws.Cells(r, 3).Value  = StatusLabel(status)
    ws.Cells(r, 4).Value  = StrConv(priority, vbProperCase)
    ws.Cells(r, 5).Value  = progress / 100
    ws.Cells(r, 5).NumberFormat = "0%"
    ws.Cells(r, 6).Value  = IIf(Trim(owner) = "", Chr(8212), owner)
    ws.Cells(r, 7).Value  = IIf(Trim(dept)  = "", Chr(8212), dept)
    ws.Cells(r, 8).Value  = Left(sd, 10)
    ws.Cells(r, 9).Value  = Left(ed, 10)
    ws.Cells(r, 10).Value = IIf(updatedAt = "", Format(Now(), "mmm dd, yyyy"), updatedAt)
    If id <> "" Then
        On Error Resume Next
        ws.Cells(r, 1).Comment.Delete
        ws.Cells(r, 1).AddComment id
        ws.Cells(r, 1).Comment.Visible = False
        On Error GoTo 0
    End If
End Sub

' ─── Initiative Form ──────────────────────────────────────────────────────────
' Clicking Create or Edit opens this hidden sheet. User fills the fields and
' clicks Save — the API is called and the Initiatives sheet updates in-place.

Public Sub ShowInitForm(mode As String, id As String)
    Dim fws As Worksheet
    On Error Resume Next
    Set fws = ThisWorkbook.Sheets(INIT_FORM)
    On Error GoTo 0
    If fws Is Nothing Then
        MsgBox "Initiative Form sheet not found." & Chr(10) & _
               "Run CreateButtons to rebuild it.", vbExclamation
        Exit Sub
    End If

    fws.Range(IFC_MODE).Value = mode
    fws.Range(IFC_ID).Value   = id
    fws.Cells(1, 1).Value = IIf(mode = "edit", _
        "  Edit Initiative", "  Create New Initiative")

    If mode = "edit" And id <> "" Then
        Dim json As String: json = Http("GET", "/initiatives/" & id)
        fws.Range(IFC_TITLE).Value  = JGet(json, "title")
        fws.Range(IFC_STATUS).Value = StatusLabel(JGet(json, "status"))
        fws.Range(IFC_PRIO).Value   = StrConv(JGet(json, "priority"), vbProperCase)
        fws.Range(IFC_PROG).Value   = Val(JGet(json, "progress"))
        fws.Range(IFC_OWNER).Value  = JGet(json, "owner")
        fws.Range(IFC_DEPT).Value   = JGet(json, "department")
        fws.Range(IFC_DESC).Value   = JGet(json, "description")
        fws.Range(IFC_SDATE).Value  = Left(JGet(json, "startDate"), 10)
        fws.Range(IFC_EDATE).Value  = Left(JGet(json, "endDate"), 10)
    Else
        fws.Range(IFC_TITLE).Value  = ""
        fws.Range(IFC_STATUS).Value = "On Track"
        fws.Range(IFC_PRIO).Value   = "Medium"
        fws.Range(IFC_PROG).Value   = 0
        fws.Range(IFC_OWNER).Value  = ""
        fws.Range(IFC_DEPT).Value   = ""
        fws.Range(IFC_DESC).Value   = ""
        fws.Range(IFC_SDATE).Value  = ""
        fws.Range(IFC_EDATE).Value  = ""
    End If

    fws.Visible = True
    fws.Activate
    fws.Range(IFC_TITLE).Select
End Sub

Public Sub SaveInitForm()
    On Error GoTo Fail
    Dim fws As Worksheet: Set fws = ThisWorkbook.Sheets(INIT_FORM)

    Dim mode     As String: mode     = Trim(fws.Range(IFC_MODE).Value)
    Dim id       As String: id       = Trim(fws.Range(IFC_ID).Value)
    Dim title    As String: title    = Trim(fws.Range(IFC_TITLE).Value)
    Dim status   As String: status   = StatusCode(fws.Range(IFC_STATUS).Value)
    Dim priority As String: priority = LCase(Trim(fws.Range(IFC_PRIO).Value))
    Dim progress As Long:   progress = CLng(Val(fws.Range(IFC_PROG).Value))
    Dim owner    As String: owner    = Trim(fws.Range(IFC_OWNER).Value)
    Dim dept     As String: dept     = Trim(fws.Range(IFC_DEPT).Value)
    Dim desc     As String: desc     = Trim(fws.Range(IFC_DESC).Value)
    Dim sd       As String: sd       = Trim(fws.Range(IFC_SDATE).Value)
    Dim ed       As String: ed       = Trim(fws.Range(IFC_EDATE).Value)

    If title = "" Then
        MsgBox "Title is required.", vbExclamation, "Missing Field"
        fws.Range(IFC_TITLE).Select: Exit Sub
    End If
    If status = "" Then
        MsgBox "Status is required.", vbExclamation, "Missing Field"
        fws.Range(IFC_STATUS).Select: Exit Sub
    End If
    If progress < 0  Then progress = 0
    If progress > 100 Then progress = 100

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Initiatives")
    Dim body As String
    body = BuildBody(title, status, priority, progress, owner, dept, desc, sd, ed)

    If mode = "edit" Then
        Http "PATCH", "/initiatives/" & id, body
        Dim r As Long
        For r = 3 To ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
            If CStr(RowId(ws, r)) = CStr(id) Then
                Call WriteInitRow(ws, r, id, CLng(ws.Cells(r, 1).Value), _
                                  title, status, priority, progress, owner, dept, _
                                  sd, ed, Format(Now(), "mmm dd, yyyy"))
                Exit For
            End If
        Next r
        fws.Visible = False
        ws.Activate
        MsgBox "Initiative updated.", vbInformation, "Saved"
    Else
        Dim resp As String: resp = Http("POST", "/initiatives", body)
        Dim newId   As String:  newId   = JGet(resp, "id")
        Dim lastRow As Long:    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        Dim newRow  As Long:    newRow  = lastRow + 1
        Call WriteInitRow(ws, newRow, newId, newRow - 1, title, status, priority, _
                         progress, owner, dept, sd, ed, "")
        ws.Rows(newRow).RowHeight = 20
        fws.Visible = False
        ws.Activate
        ws.Cells(newRow, 1).Select
        MsgBox "Initiative created.", vbInformation, "Created"
    End If
    Exit Sub
Fail:
    MsgBox "Error: " & Err.Description, vbCritical, "Save Failed"
End Sub

Public Sub CancelInitForm()
    On Error Resume Next
    ThisWorkbook.Sheets(INIT_FORM).Visible = False
    ThisWorkbook.Sheets("Initiatives").Activate
    On Error GoTo 0
End Sub

' ─── Update Form ──────────────────────────────────────────────────────────────

Public Sub ShowUpdateForm(id As String, nm As String)
    Dim fws As Worksheet
    On Error Resume Next
    Set fws = ThisWorkbook.Sheets(UPDATE_FORM)
    On Error GoTo 0
    If fws Is Nothing Then
        MsgBox "Update Form sheet not found." & Chr(10) & _
               "Run CreateButtons to rebuild it.", vbExclamation
        Exit Sub
    End If

    fws.Range(UFC_ID).Value     = id
    fws.Range(UFC_NM).Value     = nm
    fws.Range(UFC_NOTE).Value   = ""
    fws.Range(UFC_AUTHOR).Value = CurrentUser()

    fws.Visible = True
    fws.Activate
    fws.Range(UFC_NOTE).Select
End Sub

Public Sub SaveUpdateForm()
    On Error GoTo Fail
    Dim fws As Worksheet: Set fws = ThisWorkbook.Sheets(UPDATE_FORM)

    Dim id     As String: id     = Trim(fws.Range(UFC_ID).Value)
    Dim nm     As String: nm     = Trim(fws.Range(UFC_NM).Value)
    Dim note   As String: note   = Trim(fws.Range(UFC_NOTE).Value)
    Dim author As String: author = Trim(fws.Range(UFC_AUTHOR).Value)

    If note = "" Then
        MsgBox "Please enter a note.", vbExclamation, "Missing Field"
        fws.Range(UFC_NOTE).Select: Exit Sub
    End If
    If author = "" Then author = "Unknown"

    Dim body As String
    body = "{" & Q("note")         & ":" & Q(JEsc(note)) & _
           "," & Q("author")       & ":" & Q(JEsc(author)) & _
           "," & Q("initiativeId") & ":" & Q(JEsc(id)) & "}"
    Http "POST", "/initiatives/" & id & "/updates", body

    On Error Resume Next
    Dim ws3 As Worksheet: Set ws3 = ThisWorkbook.Sheets("Updates Log")
    On Error GoTo Fail
    If Not ws3 Is Nothing Then
        ws3.Rows(3).Insert Shift:=xlDown
        Dim lastNum As Long
        On Error Resume Next: lastNum = CLng(ws3.Cells(4, 1).Value): On Error GoTo 0
        ws3.Cells(3, 1).Value = lastNum + 1
        ws3.Cells(3, 2).Value = nm
        ws3.Cells(3, 3).Value = note
        ws3.Cells(3, 4).Value = author
        ws3.Cells(3, 5).Value = Format(Now(), "mmm dd, yyyy")
        ws3.Rows(3).RowHeight = 20
    End If

    fws.Visible = False
    On Error Resume Next: ThisWorkbook.Sheets("Initiatives").Activate: On Error GoTo 0
    MsgBox "Update posted.", vbInformation, "Posted"
    Exit Sub
Fail:
    MsgBox "Error: " & Err.Description, vbCritical, "Failed"
End Sub

Public Sub CancelUpdateForm()
    On Error Resume Next
    ThisWorkbook.Sheets(UPDATE_FORM).Visible = False
    ThisWorkbook.Sheets("Initiatives").Activate
    On Error GoTo 0
End Sub

' ─── PUBLIC MACROS ────────────────────────────────────────────────────────────

Public Sub TestConnection()
    On Error GoTo Fail
    Http "GET", "/initiatives?limit=1"
    MsgBox "Connected!" & Chr(10) & "API: " & ApiBase(), vbInformation, "Connection OK"
    Exit Sub
Fail:
    MsgBox "Connection failed." & Chr(10) & Chr(10) & _
           "URL tried: " & ApiBase() & Chr(10) & Chr(10) & _
           Err.Description & Chr(10) & Chr(10) & _
           "Update the URL in the Config sheet (cell B4).", _
           vbCritical, "Connection Failed"
End Sub

' Opens the Initiative Form sheet ready to create a new initiative.
Public Sub CreateInitiative()
    ShowInitForm "create", ""
End Sub

' Opens the Initiative Form sheet pre-filled with the selected row's data.
Public Sub EditInitiative()
    Dim ws As Worksheet: Set ws = InitWS()
    If ws Is Nothing Then Exit Sub
    Dim r As Long: r = ActiveCell.Row
    If r <= 2 Then
        MsgBox "Click on a data row first (row 3 or below).", vbExclamation: Exit Sub
    End If
    Dim id As String: id = RowId(ws, r)
    If id = "" Then
        MsgBox "Initiative ID not found." & Chr(10) & _
               "Re-export from the app to embed IDs.", vbExclamation: Exit Sub
    End If
    ShowInitForm "edit", id
End Sub

' Confirms then deletes the selected initiative from the API and the sheet.
Public Sub DeleteInitiative()
    On Error GoTo Fail
    Dim ws As Worksheet: Set ws = InitWS()
    If ws Is Nothing Then Exit Sub
    Dim r As Long: r = ActiveCell.Row
    If r <= 2 Then
        MsgBox "Click on a data row first.", vbExclamation: Exit Sub
    End If
    Dim id As String: id = RowId(ws, r)
    If id = "" Then
        MsgBox "ID not found. Re-export from the app.", vbExclamation: Exit Sub
    End If
    Dim nm As String: nm = ws.Cells(r, 2).Value
    If MsgBox("Permanently delete """ & nm & """?" & Chr(10) & Chr(10) & _
              "All updates for this initiative will also be deleted." & Chr(10) & _
              "This cannot be undone.", _
              vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Delete") <> vbYes Then
        Exit Sub
    End If
    Http "DELETE", "/initiatives/" & id
    ws.Rows(r).Delete
    MsgBox "Deleted.", vbInformation, "Done"
    Exit Sub
Fail:
    MsgBox "Error: " & Err.Description, vbCritical, "Delete Failed"
End Sub

' Opens the Update Form sheet for the selected initiative.
Public Sub AddUpdate()
    Dim ws As Worksheet: Set ws = InitWS()
    If ws Is Nothing Then Exit Sub
    Dim r As Long: r = ActiveCell.Row
    If r <= 2 Then
        MsgBox "Click on a data row first.", vbExclamation: Exit Sub
    End If
    Dim id As String: id = RowId(ws, r)
    If id = "" Then
        MsgBox "ID not found. Re-export from the app.", vbExclamation: Exit Sub
    End If
    ShowUpdateForm id, ws.Cells(r, 2).Value
End Sub

' ─── Buttons ──────────────────────────────────────────────────────────────────

' RefreshAllData — pulls all data from the live API and repopulates all sheets.
' Run this any time you want to sync the workbook with the latest app data.
' Summary KPIs, every Initiatives row, and the full Updates Log are rebuilt from scratch.
Public Sub RefreshAllData()
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' ── Fetch all data ────────────────────────────────────────────────────────
    Dim initsJson   As String: initsJson   = Http("GET", "/initiatives")
    Dim summaryJson As String: summaryJson = Http("GET", "/dashboard/summary")
    Dim actJson     As String: actJson     = Http("GET", "/dashboard/recent-activity?limit=500")

    Dim inits    As Variant: inits    = JSplitArray(initsJson)
    Dim activity As Variant: activity = JSplitArray(actJson)

    ' ── Initiatives sheet ─────────────────────────────────────────────────────
    Dim ws2 As Worksheet: Set ws2 = ThisWorkbook.Sheets("Initiatives")
    Dim lr2 As Long: lr2 = ws2.Cells(ws2.Rows.Count, 2).End(xlUp).Row
    If lr2 >= 3 Then ws2.Rows("3:" & lr2).Delete

    Dim i   As Long
    Dim ini As String
    For i = 0 To UBound(inits)
        ini = CStr(inits(i))
        ws2.Rows(i + 3).RowHeight = 20
        Call WriteInitRow(ws2, i + 3, _
            JGet(ini, "id"),       i + 1, _
            JGet(ini, "title"),    JGet(ini, "status"),      JGet(ini, "priority"), _
            CLng(Val(JGet(ini, "progress"))), _
            JGet(ini, "owner"),    JGet(ini, "department"), _
            Left(JGet(ini, "startDate"), 10), _
            Left(JGet(ini, "endDate"),   10), _
            Left(JGet(ini, "updatedAt"), 10))
    Next i

    Dim nInits As Long: nInits = IIf(UBound(inits) >= 0, UBound(inits) + 1, 0)
    ws2.Cells(1, 1).Value = "  All Initiatives  " & Chr(183) & "  " & nInits & _
        " total  " & Chr(183) & _
        "  CreateInitiative | EditInitiative | DeleteInitiative | AddUpdate | RefreshAllData"

    ' ── Updates Log sheet ─────────────────────────────────────────────────────
    Dim ws3 As Worksheet: Set ws3 = ThisWorkbook.Sheets("Updates Log")
    Dim lr3 As Long: lr3 = ws3.Cells(ws3.Rows.Count, 2).End(xlUp).Row
    If lr3 >= 3 Then ws3.Rows("3:" & lr3).Delete

    Dim act As String
    For i = 0 To UBound(activity)
        act = CStr(activity(i))
        ws3.Cells(i + 3, 1).Value = i + 1
        ws3.Cells(i + 3, 2).Value = JGet(act, "initiativeTitle")
        ws3.Cells(i + 3, 3).Value = JGet(act, "note")
        ws3.Cells(i + 3, 4).Value = JGet(act, "author")
        ws3.Cells(i + 3, 5).Value = Left(JGet(act, "createdAt"), 10)
        ws3.Rows(i + 3).RowHeight = 20
    Next i

    Dim nUpdates As Long: nUpdates = IIf(UBound(activity) >= 0, UBound(activity) + 1, 0)
    ws3.Cells(1, 1).Value = "  Progress Updates  " & Chr(183) & "  " & nUpdates & " entries"

    ' ── Summary KPI cards ─────────────────────────────────────────────────────
    ' The KPI numbers live in the merged range at Excel rows 5-6, columns B-G.
    ' Writing to the top-left cell (row 5) of each merge is sufficient.
    Dim ws1 As Worksheet: Set ws1 = ThisWorkbook.Sheets("Summary")
    ws1.Cells(3, 2).Value = "Data as of " & Format(Now(), "mmmm dd, yyyy")
    ws1.Cells(5, 2).Value = CLng(Val(JGet(summaryJson, "total")))
    ws1.Cells(5, 3).Value = CLng(Val(JGet(summaryJson, "onTrack")))
    ws1.Cells(5, 4).Value = CLng(Val(JGet(summaryJson, "atRisk")))
    ws1.Cells(5, 5).Value = CLng(Val(JGet(summaryJson, "delayed")))
    ws1.Cells(5, 6).Value = CLng(Val(JGet(summaryJson, "completed")))
    ws1.Cells(5, 7).Value = JGet(summaryJson, "avgProgress") & "%"

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Refresh complete!" & Chr(10) & Chr(10) & _
           nInits   & " initiative(s)" & Chr(10) & _
           nUpdates & " update(s)", vbInformation, "Data Refreshed"
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Refresh failed:" & Chr(10) & Err.Description, vbCritical, "Error"
End Sub

' Runs on file open — recreates sidebar buttons if missing.
Public Sub Auto_Open()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Initiatives")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left(shp.Name, 9) = "MacroBtn_" Then Exit Sub
    Next shp
    CreateButtons
End Sub

' CreateButtons builds the sidebar on the Initiatives sheet AND adds Save/Cancel
' buttons to the two form sheets. Run once after importing the module; buttons
' are saved in the workbook and survive re-open.
Public Sub CreateButtons()
    ' ── Initiatives sheet sidebar ─────────────────────────────────────────────
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Initiatives")
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "Initiatives sheet not found.", vbExclamation: Exit Sub

    Dim shp As Shape, toDelete() As String, n As Long: n = 0
    For Each shp In ws.Shapes
        If Left(shp.Name, 9) = "MacroBtn_" Then
            ReDim Preserve toDelete(n): toDelete(n) = shp.Name: n = n + 1
        End If
    Next shp
    Dim j As Long
    For j = 0 To n - 1: ws.Shapes(toDelete(j)).Delete: Next j

    ws.Columns("L:L").ColumnWidth = 20

    Dim defs(5, 2) As Variant
    defs(0, 0) = "TestConnection":   defs(0, 1) = "Test Connection":   defs(0, 2) = RGB(71, 85, 105)
    defs(1, 0) = "CreateInitiative": defs(1, 1) = "Create Initiative":  defs(1, 2) = RGB(21, 128, 61)
    defs(2, 0) = "EditInitiative":   defs(2, 1) = "Edit Selected Row":  defs(2, 2) = RGB(29, 78, 216)
    defs(3, 0) = "DeleteInitiative": defs(3, 1) = "Delete Selected":    defs(3, 2) = RGB(185, 28, 28)
    defs(4, 0) = "AddUpdate":        defs(4, 1) = "Add Update Note":     defs(4, 2) = RGB(15, 37, 68)
    defs(5, 0) = "RefreshAllData":   defs(5, 1) = "Refresh All Data":   defs(5, 2) = RGB(13, 148, 136)

    Dim btnLeft As Double:   btnLeft   = ws.Range("L2").Left + 3
    Dim btnWidth As Double:  btnWidth  = ws.Range("L2").Width - 6
    Dim btnHeight As Double: btnHeight = 26
    Dim gap As Double:       gap       = 5
    Dim s As Shape
    Dim i As Integer
    For i = 0 To 5
        Dim macroName As String: macroName = CStr(defs(i, 0))
        Dim caption   As String: caption   = CStr(defs(i, 1))
        Dim fillColor As Long:   fillColor = CLng(defs(i, 2))
        Dim btnTop    As Double: btnTop    = ws.Range("A2").Top + 2 + i * (btnHeight + gap)
        Set s = ws.Shapes.AddShape(5, btnLeft, btnTop, btnWidth, btnHeight)
        With s
            .Name = "MacroBtn_" & macroName: .OnAction = macroName
            .Fill.ForeColor.RGB = fillColor: .Fill.Solid: .Line.Visible = False
            With .TextFrame
                .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
                .HorizontalAlignment = 2: .VerticalAlignment = 2
                .Characters.Text = caption
                With .Characters.Font: .Color = RGB(255,255,255): .Bold = True: .Size = 9: End With
            End With
        End With
    Next i

    ' ── Initiative Form sheet buttons ─────────────────────────────────────────
    Dim ifs As Worksheet
    On Error Resume Next: Set ifs = ThisWorkbook.Sheets(INIT_FORM): On Error GoTo 0
    If Not ifs Is Nothing Then
        ' Remove old form buttons
        Dim fd() As String: Dim fn As Long: fn = 0
        For Each shp In ifs.Shapes
            If Left(shp.Name, 8) = "FormBtn_" Then
                ReDim Preserve fd(fn): fd(fn) = shp.Name: fn = fn + 1
            End If
        Next shp
        For j = 0 To fn - 1: ifs.Shapes(fd(j)).Delete: Next j

        Dim iBtnLeft  As Double: iBtnLeft  = ifs.Range("B13").Left
        Dim iBtnTop   As Double: iBtnTop   = ifs.Range("B13").Top + 4

        ' Save (green)
        Set s = ifs.Shapes.AddShape(5, iBtnLeft, iBtnTop, 110, 28)
        With s
            .Name = "FormBtn_Save": .OnAction = "SaveInitForm"
            .Fill.ForeColor.RGB = RGB(21, 128, 61): .Fill.Solid: .Line.Visible = False
            With .TextFrame
                .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
                .HorizontalAlignment = 2: .VerticalAlignment = 2
                .Characters.Text = "Save"
                With .Characters.Font: .Color = RGB(255,255,255): .Bold = True: .Size = 10: End With
            End With
        End With

        ' Cancel (slate)
        Set s = ifs.Shapes.AddShape(5, iBtnLeft + 118, iBtnTop, 110, 28)
        With s
            .Name = "FormBtn_Cancel": .OnAction = "CancelInitForm"
            .Fill.ForeColor.RGB = RGB(71, 85, 105): .Fill.Solid: .Line.Visible = False
            With .TextFrame
                .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
                .HorizontalAlignment = 2: .VerticalAlignment = 2
                .Characters.Text = "Cancel"
                With .Characters.Font: .Color = RGB(255,255,255): .Bold = True: .Size = 10: End With
            End With
        End With
    End If

    ' ── Update Form sheet buttons ─────────────────────────────────────────────
    Dim ufs As Worksheet
    On Error Resume Next: Set ufs = ThisWorkbook.Sheets(UPDATE_FORM): On Error GoTo 0
    If Not ufs Is Nothing Then
        Dim ud() As String: Dim un As Long: un = 0
        For Each shp In ufs.Shapes
            If Left(shp.Name, 8) = "FormBtn_" Then
                ReDim Preserve ud(un): ud(un) = shp.Name: un = un + 1
            End If
        Next shp
        For j = 0 To un - 1: ufs.Shapes(ud(j)).Delete: Next j

        Dim uBtnLeft As Double: uBtnLeft = ufs.Range("B8").Left
        Dim uBtnTop  As Double: uBtnTop  = ufs.Range("B8").Top + 4

        ' Post (blue)
        Set s = ufs.Shapes.AddShape(5, uBtnLeft, uBtnTop, 130, 28)
        With s
            .Name = "FormBtn_Post": .OnAction = "SaveUpdateForm"
            .Fill.ForeColor.RGB = RGB(29, 78, 216): .Fill.Solid: .Line.Visible = False
            With .TextFrame
                .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
                .HorizontalAlignment = 2: .VerticalAlignment = 2
                .Characters.Text = "Post Update"
                With .Characters.Font: .Color = RGB(255,255,255): .Bold = True: .Size = 10: End With
            End With
        End With

        ' Cancel (slate)
        Set s = ufs.Shapes.AddShape(5, uBtnLeft + 138, uBtnTop, 110, 28)
        With s
            .Name = "FormBtn_CancelU": .OnAction = "CancelUpdateForm"
            .Fill.ForeColor.RGB = RGB(71, 85, 105): .Fill.Solid: .Line.Visible = False
            With .TextFrame
                .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
                .HorizontalAlignment = 2: .VerticalAlignment = 2
                .Characters.Text = "Cancel"
                With .Characters.Font: .Color = RGB(255,255,255): .Bold = True: .Size = 10: End With
            End With
        End With
    End If

    MsgBox "Buttons created on all sheets." & Chr(10) & _
           "Save as .xlsm to keep them.", vbInformation, "Done"
End Sub
