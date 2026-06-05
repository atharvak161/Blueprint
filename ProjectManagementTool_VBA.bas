
'=============================================================================
' PROJECT MANAGEMENT TOOL - VBA MODULE
' Version 1.0 | Import via: VBA Editor (Alt+F11) > File > Import File
'=============================================================================
Option Explicit

'─────────────────────────────────────────────────────────────────────────────
' CONSTANTS & HELPERS
'─────────────────────────────────────────────────────────────────────────────
Private Const PLAN_SHEET As String = "Project Plan"
Private Const DATA_START_ROW As Long = 12
Private Const COL_ID As Integer = 1
Private Const COL_TYPE As Integer = 2
Private Const COL_PHASE As Integer = 3
Private Const COL_DESC As Integer = 4
Private Const COL_START As Integer = 5
Private Const COL_END As Integer = 6
Private Const COL_RESOURCE As Integer = 7
Private Const COL_PCT As Integer = 8
Private Const COL_STATUS As Integer = 9
Private Const COL_NOTES As Integer = 10

'─────────────────────────────────────────────────────────────────────────────
' SETUP PROJECT UI
' Run this once after opening the workbook (or after adding new rows)
'─────────────────────────────────────────────────────────────────────────────
Sub SetupProjectUI()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(PLAN_SHEET)
    Application.ScreenUpdating = False

    ' Apply conditional formatting for Status column (I)
    Dim lastRow As Long
    lastRow = FindLastDataRow(ws)

    With ws.Range("I" & DATA_START_ROW & ":I" & lastRow + 500)
        .FormatConditions.Delete

        ' Completed - Green
        Dim cfComplete As FormatCondition
        Set cfComplete = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Completed""")
        cfComplete.Interior.Color = RGB(0, 176, 80)
        cfComplete.Font.Color = RGB(255, 255, 255)
        cfComplete.Font.Bold = True

        ' In Progress - Amber
        Dim cfInProg As FormatCondition
        Set cfInProg = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""In Progress""")
        cfInProg.Interior.Color = RGB(255, 185, 0)
        cfInProg.Font.Color = RGB(0, 0, 0)
        cfInProg.Font.Bold = True

        ' Overdue - Red
        Dim cfOverdue As FormatCondition
        Set cfOverdue = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Overdue""")
        cfOverdue.Interior.Color = RGB(255, 0, 0)
        cfOverdue.Font.Color = RGB(255, 255, 255)
        cfOverdue.Font.Bold = True

        ' Blocked - Purple
        Dim cfBlocked As FormatCondition
        Set cfBlocked = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Blocked""")
        cfBlocked.Interior.Color = RGB(112, 48, 160)
        cfBlocked.Font.Color = RGB(255, 255, 255)
        cfBlocked.Font.Bold = True

        ' Not Started - Grey
        Dim cfNotStart As FormatCondition
        Set cfNotStart = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Not Started""")
        cfNotStart.Interior.Color = RGB(165, 165, 165)
        cfNotStart.Font.Color = RGB(255, 255, 255)
        cfNotStart.Font.Bold = True
    End With

    MsgBox "Project UI setup complete. Conditional formatting applied.", vbInformation, "Setup Complete"
    Application.ScreenUpdating = True
End Sub

'─────────────────────────────────────────────────────────────────────────────
' AUTO-RENUMBER IDs
'─────────────────────────────────────────────────────────────────────────────
Sub RenumberIDs()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(PLAN_SHEET)
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim i As Long
    Dim taskID As Long
    taskID = 0

    Dim lastRow As Long
    lastRow = FindLastDataRow(ws)

    For i = DATA_START_ROW To lastRow
        Dim cellType As String
        cellType = ws.Cells(i, COL_TYPE).Value
        If cellType = "Task" Or cellType = "Subtask" Then
            taskID = taskID + 1
            ws.Cells(i, COL_ID).Value = taskID
        ElseIf cellType = "Phase" Then
            ws.Cells(i, COL_ID).Value = ""
        End If
    Next i

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

'─────────────────────────────────────────────────────────────────────────────
' IS WORKING DAY (skips weekends and UK holidays)
'─────────────────────────────────────────────────────────────────────────────
Function IsWorkingDay(dt As Date) As Boolean
    If Weekday(dt, vbMonday) >= 6 Then
        IsWorkingDay = False
        Exit Function
    End If

    Dim wsH As Worksheet
    On Error Resume Next
    Set wsH = ThisWorkbook.Sheets("Holidays")
    On Error GoTo 0

    If wsH Is Nothing Then
        IsWorkingDay = True
        Exit Function
    End If

    Dim lastHRow As Long
    lastHRow = wsH.Cells(wsH.Rows.Count, 2).End(xlUp).Row

    Dim r As Long
    For r = 4 To lastHRow
        Dim hStart As Date
        Dim hEnd As Date
        If wsH.Cells(r, 2).Value <> "" And IsDate(wsH.Cells(r, 2).Value) Then
            hStart = CDate(wsH.Cells(r, 2).Value)
            If wsH.Cells(r, 3).Value <> "" Then
                hEnd = CDate(wsH.Cells(r, 3).Value)
            Else
                hEnd = hStart
            End If
            If dt >= hStart And dt <= hEnd Then
                IsWorkingDay = False
                Exit Function
            End If
        End If
    Next r

    IsWorkingDay = True
End Function

Function NextWorkingDay(dt As Date) As Date
    Dim d As Date
    d = dt + 1
    Do While Not IsWorkingDay(d)
        d = d + 1
    Loop
    NextWorkingDay = d
End Function

Function AddWorkingDays(startDate As Date, days As Long) As Date
    Dim d As Date
    d = startDate
    Dim remaining As Long
    remaining = days - 1
    Do While remaining > 0
        d = d + 1
        If IsWorkingDay(d) Then remaining = remaining - 1
    Loop
    AddWorkingDays = d
End Function

'─────────────────────────────────────────────────────────────────────────────
' AUTO-ADJUST DATES FOR HOLIDAYS
'─────────────────────────────────────────────────────────────────────────────
Sub AdjustDatesForHolidays()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(PLAN_SHEET)
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim lastRow As Long
    lastRow = FindLastDataRow(ws)
    Dim i As Long

    For i = DATA_START_ROW To lastRow
        Dim cellType As String
        cellType = ws.Cells(i, COL_TYPE).Value
        If cellType = "Task" Or cellType = "Subtask" Then
            If IsDate(ws.Cells(i, COL_START).Value) Then
                Dim startDt As Date
                startDt = CDate(ws.Cells(i, COL_START).Value)
                ' Move start to next working day if it falls on holiday
                If Not IsWorkingDay(startDt) Then
                    Do While Not IsWorkingDay(startDt)
                        startDt = startDt + 1
                    Loop
                    ws.Cells(i, COL_START).Value = startDt
                End If
                ' Recalculate end date if duration exists (end = start + duration - 1 working days)
                ' End date cell is protected but we update it here
                If IsDate(ws.Cells(i, COL_END).Value) Then
                    Dim endDt As Date
                    endDt = CDate(ws.Cells(i, COL_END).Value)
                    If Not IsWorkingDay(endDt) Then
                        Do While Not IsWorkingDay(endDt)
                            endDt = endDt + 1
                        Loop
                        ws.Cells(i, COL_END).Value = endDt
                    End If
                End If
            End If
        End If
    Next i

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Dates adjusted. All task start/end dates now fall on working days.", vbInformation, "Date Adjustment Complete"
End Sub

'─────────────────────────────────────────────────────────────────────────────
' HELPER: Find last row with data
'─────────────────────────────────────────────────────────────────────────────
Function FindLastDataRow(ws As Worksheet) As Long
    FindLastDataRow = ws.Cells(ws.Rows.Count, COL_DESC).End(xlUp).Row
End Function

'─────────────────────────────────────────────────────────────────────────────
' COLLECT ALL TASKS INTO ARRAY FOR EXPORT
'─────────────────────────────────────────────────────────────────────────────
Function CollectTasks() As Variant
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(PLAN_SHEET)

    Dim lastRow As Long
    lastRow = FindLastDataRow(ws)

    Dim tasks() As Variant
    Dim taskCount As Long
    taskCount = 0

    Dim i As Long
    For i = DATA_START_ROW To lastRow
        If ws.Cells(i, COL_TYPE).Value <> "" Then
            taskCount = taskCount + 1
        End If
    Next i

    ReDim tasks(1 To taskCount, 1 To 10)
    Dim idx As Long
    idx = 0

    For i = DATA_START_ROW To lastRow
        If ws.Cells(i, COL_TYPE).Value <> "" Then
            idx = idx + 1
            tasks(idx, 1) = ws.Cells(i, COL_ID).Value        ' ID
            tasks(idx, 2) = ws.Cells(i, COL_TYPE).Value      ' Type
            tasks(idx, 3) = ws.Cells(i, COL_PHASE).Value     ' Phase
            tasks(idx, 4) = ws.Cells(i, COL_DESC).Value      ' Description
            If IsDate(ws.Cells(i, COL_START).Value) Then
                tasks(idx, 5) = Format(ws.Cells(i, COL_START).Value, "DD/MM/YYYY")
            Else
                tasks(idx, 5) = ""
            End If
            If IsDate(ws.Cells(i, COL_END).Value) Then
                tasks(idx, 6) = Format(ws.Cells(i, COL_END).Value, "DD/MM/YYYY")
            Else
                tasks(idx, 6) = ""
            End If
            tasks(idx, 7) = ws.Cells(i, COL_RESOURCE).Value  ' Resource
            tasks(idx, 8) = ws.Cells(i, COL_PCT).Value       ' % Done
            tasks(idx, 9) = ws.Cells(i, COL_STATUS).Value    ' Status
            tasks(idx, 10) = ws.Cells(i, COL_NOTES).Value    ' Notes
        End If
    Next i

    CollectTasks = tasks
End Function

'─────────────────────────────────────────────────────────────────────────────
' GET WORKBOOK FOLDER PATH
'─────────────────────────────────────────────────────────────────────────────
Function GetSavePath(filename As String) As String
    Dim wbPath As String
    wbPath = ThisWorkbook.Path
    If wbPath = "" Then
        wbPath = Environ("USERPROFILE") & "\Desktop"
    End If
    GetSavePath = wbPath & "\" & filename
End Function

'─────────────────────────────────────────────────────────────────────────────
' WRITE STRING TO FILE
'─────────────────────────────────────────────────────────────────────────────
Sub WriteFile(filePath As String, content As String)
    Dim fileNum As Integer
    fileNum = FreeFile()
    Open filePath For Output As #fileNum
    Print #fileNum, content
    Close #fileNum
End Sub

'─────────────────────────────────────────────────────────────────────────────
' OPEN FILE IN DEFAULT BROWSER
'─────────────────────────────────────────────────────────────────────────────
Sub OpenInBrowser(filePath As String)
    Shell "cmd /c start """" """ & filePath & """", vbHide
End Sub

'=============================================================================
' EXPORT 1: INTERACTIVE GANTT CHART
'=============================================================================
Sub ExportGanttChart()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(PLAN_SHEET)

    Dim projectName As String
    projectName = ws.Range("E2").Value
    If projectName = "" Then projectName = "Project Gantt Chart"

    Dim tasks As Variant
    tasks = CollectTasks()

    If Not IsArray(tasks) Then
        MsgBox "No task data found.", vbExclamation
        Exit Sub
    End If

    ' Build tasks JSON
    Dim jsonTasks As String
    jsonTasks = "["
    Dim i As Long
    For i = 1 To UBound(tasks, 1)
        If i > 1 Then jsonTasks = jsonTasks & ","
        Dim pct As Double
        If IsNumeric(tasks(i, 8)) Then pct = tasks(i, 8) Else pct = 0
        Dim startStr As String, endStr As String
        startStr = tasks(i, 5)
        endStr = tasks(i, 6)
        jsonTasks = jsonTasks & "{"
        jsonTasks = jsonTasks & """id"":" & IIf(tasks(i, 1) = "", "0", tasks(i, 1)) & ","
        jsonTasks = jsonTasks & """type"":""" & tasks(i, 2) & ""","
        jsonTasks = jsonTasks & """phase"":""" & EscapeJson(CStr(tasks(i, 3))) & ""","
        jsonTasks = jsonTasks & """desc"":""" & EscapeJson(CStr(tasks(i, 4))) & ""","
        jsonTasks = jsonTasks & """start"":""" & startStr & ""","
        jsonTasks = jsonTasks & """end"":""" & endStr & ""","
        jsonTasks = jsonTasks & """resource"":""" & EscapeJson(CStr(tasks(i, 7))) & ""","
        jsonTasks = jsonTasks & """pct"":" & pct & ","
        jsonTasks = jsonTasks & """status"":""" & EscapeJson(CStr(tasks(i, 9))) & ""","
        jsonTasks = jsonTasks & """notes"":""" & EscapeJson(CStr(tasks(i, 10))) & """"
        jsonTasks = jsonTasks & "}"
    Next i
    jsonTasks = jsonTasks & "]"

    Dim html As String
    html = BuildGanttHTML(projectName, jsonTasks)

    Dim savePath As String
    savePath = GetSavePath("ProjectGanttChart.html")
    WriteFile savePath, html
    OpenInBrowser savePath
    MsgBox "Gantt Chart exported to:" & vbCrLf & savePath, vbInformation, "Export Complete"
End Sub

Function EscapeJson(s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, Chr(10), "\n")
    s = Replace(s, Chr(13), "")
    EscapeJson = s
End Function

Function BuildGanttHTML(projectName As String, jsonTasks As String) As String
    Dim h As String
    h = "<!DOCTYPE html><html lang=""en""><head><meta charset=""UTF-8"">"
    h = h & "<meta name=""viewport"" content=""width=device-width,initial-scale=1"">"
    h = h & "<title>Gantt Chart - " & projectName & "</title>"
    h = h & "<style>"
    h = h & "*{box-sizing:border-box;margin:0;padding:0;font-family:'Segoe UI',Arial,sans-serif}"
    h = h & "body{background:#0d1b2a;color:#e0e8f0;min-height:100vh}"
    h = h & "header{background:linear-gradient(135deg,#1f4e79,#2e75b6);padding:20px 30px;display:flex;align-items:center;justify-content:space-between;box-shadow:0 2px 12px rgba(0,0,0,0.4)}"
    h = h & "header h1{font-size:1.6rem;font-weight:700;letter-spacing:.5px}"
    h = h & "header .meta{font-size:.85rem;opacity:.8}"
    h = h & ".controls{background:#152030;padding:14px 30px;display:flex;gap:12px;align-items:center;flex-wrap:wrap;border-bottom:1px solid #243a50}"
    h = h & ".controls input,.controls select{background:#1e3248;border:1px solid #2e75b6;color:#e0e8f0;padding:6px 12px;border-radius:6px;font-size:.85rem}"
    h = h & ".controls label{font-size:.85rem;color:#93b5d0}"
    h = h & ".legend{display:flex;gap:12px;margin-left:auto;align-items:center;flex-wrap:wrap}"
    h = h & ".leg{display:flex;align-items:center;gap:5px;font-size:.8rem}"
    h = h & ".leg-dot{width:12px;height:12px;border-radius:3px}"
    h = h & ".gantt-wrap{padding:20px 30px;overflow-x:auto}"
    h = h & ".gantt-container{min-width:1200px}"
    h = h & ".gantt-header{display:grid;background:#1a2e42;border-radius:8px 8px 0 0;border:1px solid #2e4a63}"
    h = h & ".gantt-row{display:grid;border-bottom:1px solid #1e3550;transition:background .15s}"
    h = h & ".gantt-row:hover{background:rgba(46,117,182,.15)}"
    h = h & ".gantt-row.phase-row{background:linear-gradient(90deg,#1f4e79,#1a3f63);font-weight:700;font-size:.9rem}"
    h = h & ".gantt-row.subtask-row .task-info{padding-left:28px;font-style:italic;font-size:.85rem}"
    h = h & ".task-info{padding:8px 12px;border-right:1px solid #2e4a63;font-size:.88rem;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}"
    h = h & ".task-chart{position:relative;height:36px;display:flex;align-items:center}"
    h = h & ".gantt-bar{position:absolute;height:22px;border-radius:4px;display:flex;align-items:center;padding:0 6px;font-size:.75rem;font-weight:600;color:#fff;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;box-shadow:0 2px 6px rgba(0,0,0,.3);cursor:pointer;transition:filter .2s}"
    h = h & ".gantt-bar:hover{filter:brightness(1.2)}"
    h = h & ".gantt-bar .fill{position:absolute;left:0;top:0;height:100%;background:rgba(255,255,255,.2);border-radius:4px 0 0 4px}"
    h = h & ".today-line{position:absolute;top:0;bottom:0;width:2px;background:#ff4444;z-index:10;pointer-events:none}"
    h = h & ".today-line::after{content:'TODAY';position:absolute;top:-18px;left:50%;transform:translateX(-50%);font-size:.65rem;color:#ff4444;font-weight:700;white-space:nowrap}"
    h = h & ".date-header{display:flex;border-bottom:1px solid #2e4a63;background:#152030;min-height:32px}"
    h = h & ".date-col{border-right:1px solid #2e4a63;font-size:.7rem;text-align:center;padding:4px 2px;color:#93b5d0}"
    h = h & ".status-badge{display:inline-block;padding:2px 8px;border-radius:10px;font-size:.75rem;font-weight:700}"
    h = h & ".stat-Completed{background:#00b050;color:#fff}"
    h = h & ".stat-InProgress{background:#ffb900;color:#000}"
    h = h & ".stat-NotStarted{background:#a5a5a5;color:#fff}"
    h = h & ".stat-Overdue{background:#ff0000;color:#fff}"
    h = h & ".stat-Blocked{background:#7030a0;color:#fff}"
    h = h & ".tooltip{position:fixed;background:#1a2e42;border:1px solid #2e75b6;border-radius:8px;padding:12px;font-size:.82rem;pointer-events:none;z-index:999;max-width:280px;box-shadow:0 4px 20px rgba(0,0,0,.5);display:none}"
    h = h & ".tooltip strong{color:#2e75b6;display:block;margin-bottom:6px;font-size:.9rem}"
    h = h & ".tooltip .tt-row{display:flex;justify-content:space-between;gap:12px;margin:3px 0;color:#c0d5e8}"
    h = h & ".footer{text-align:center;padding:20px;color:#3a6080;font-size:.8rem;border-top:1px solid #1e3550}"
    h = h & "</style></head><body>"
    h = h & "<header><div><h1>&#128202; " & projectName & " — Gantt Chart</h1>"
    h = h & "<div class=""meta"">Generated: " & Format(Now(), "DD MMM YYYY HH:MM") & " &nbsp;|&nbsp; Exported from Excel</div></div></header>"
    h = h & "<div class=""controls"">"
    h = h & "<label>Search: <input type=""text"" id=""searchBox"" placeholder=""Filter tasks..."" oninput=""filterTasks()""></label>"
    h = h & "<label>Status: <select id=""statusFilter"" onchange=""filterTasks()"">"
    h = h & "<option value="""">All Statuses</option>"
    h = h & "<option>Completed</option><option>In Progress</option><option>Not Started</option><option>Overdue</option><option>Blocked</option>"
    h = h & "</select></label>"
    h = h & "<label>Resource: <select id=""resourceFilter"" onchange=""filterTasks()""><option value="""">All Resources</option></select></label>"
    h = h & "<div class=""legend"">"
    h = h & "<span class=""leg""><span class=""leg-dot"" style=""background:#00b050""></span>Completed</span>"
    h = h & "<span class=""leg""><span class=""leg-dot"" style=""background:#ffb900""></span>In Progress</span>"
    h = h & "<span class=""leg""><span class=""leg-dot"" style=""background:#a5a5a5""></span>Not Started</span>"
    h = h & "<span class=""leg""><span class=""leg-dot"" style=""background:#ff0000""></span>Overdue</span>"
    h = h & "<span class=""leg""><span class=""leg-dot"" style=""background:#7030a0""></span>Blocked</span>"
    h = h & "</div></div>"
    h = h & "<div class=""gantt-wrap""><div id=""ganttContainer"" class=""gantt-container""></div></div>"
    h = h & "<div class=""tooltip"" id=""tooltip""></div>"
    h = h & "<div class=""footer"">&#128202; Interactive Gantt Chart &mdash; " & projectName & " &mdash; Generated " & Format(Now(), "DD MMM YYYY") & "</div>"
    h = h & "<script>"
    h = h & "const TASKS=" & jsonTasks & ";"
    h = h & "const STATUS_COLORS={'Completed':'#00b050','In Progress':'#ffb900','Not Started':'#a5a5a5','Overdue':'#ff0000','Blocked':'#7030a0'};"
    h = h & "function parseDate(s){if(!s)return null;const[d,m,y]=s.split('/');return new Date(y,m-1,d);}"
    h = h & "function fmtDate(d){if(!d)return'';return d.toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'});}"
    h = h & "function buildGantt(tasks){"
    h = h & "const dates=tasks.filter(t=>t.start&&t.end).map(t=>[parseDate(t.start),parseDate(t.end)]).flat().filter(Boolean);"
    h = h & "if(!dates.length)return;"
    h = h & "let minDate=new Date(Math.min(...dates));let maxDate=new Date(Math.max(...dates));"
    h = h & "minDate.setDate(minDate.getDate()-7);maxDate.setDate(maxDate.getDate()+14);"
    h = h & "const totalDays=(maxDate-minDate)/(1000*86400);"
    h = h & "const INFO_WIDTH=320;const CHART_WIDTH=Math.max(900,totalDays*24);"
    h = h & "const COL_GRID=`${INFO_WIDTH}px 80px 80px ${CHART_WIDTH}px`;"
    h = h & "const container=document.getElementById('ganttContainer');"
    h = h & "container.innerHTML='';"
    h = h & "// Date header"
    h = h & "const hdr=document.createElement('div');hdr.className='gantt-header';"
    h = h & "hdr.style.gridTemplateColumns=COL_GRID;"
    h = h & "hdr.innerHTML=`<div class='task-info' style='font-weight:700;color:#93b5d0;background:#1a2e42'>Task / Description</div><div class='task-info' style='background:#1a2e42;color:#93b5d0;text-align:center'>Start</div><div class='task-info' style='background:#1a2e42;color:#93b5d0;text-align:center'>End</div>`;"
    h = h & "const dateHdr=document.createElement('div');dateHdr.className='date-header';dateHdr.style.position='relative';"
    h = h & "const today=new Date();let cur=new Date(minDate);"
    h = h & "while(cur<=maxDate){"
    h = h & "const pct=((cur-minDate)/(maxDate-minDate))*100;"
    h = h & "const col=document.createElement('div');col.className='date-col';"
    h = h & "col.style.width=(CHART_WIDTH/totalDays)+'px';col.style.flexShrink='0';"
    h = h & "if(cur.getDate()===1||cur.getDay()===1){col.textContent=cur.toLocaleDateString('en-GB',{day:'2-digit',month:'short'});}"
    h = h & "if(cur.toDateString()===today.toDateString())col.style.background='rgba(255,68,68,.3)';"
    h = h & "dateHdr.appendChild(col);cur.setDate(cur.getDate()+1);}"
    h = h & "hdr.appendChild(dateHdr);container.appendChild(hdr);"
    h = h & "tasks.forEach(t=>{"
    h = h & "const row=document.createElement('div');row.className='gantt-row'+(t.type==='Phase'?' phase-row':t.type==='Subtask'?' subtask-row':'');"
    h = h & "row.style.gridTemplateColumns=COL_GRID;"
    h = h & "row.dataset.desc=(t.desc||'').toLowerCase();row.dataset.status=t.status||'';row.dataset.resource=t.resource||'';"
    h = h & "row.innerHTML=`<div class='task-info'>${t.type==='Phase'?'&#9660; ':''}${t.desc||''}</div>`+"
    h = h & "`<div class='task-info' style='text-align:center;font-size:.8rem'>${t.start||''}</div>`+"
    h = h & "`<div class='task-info' style='text-align:center;font-size:.8rem'>${t.end||''}</div>`;"
    h = h & "const chart=document.createElement('div');chart.className='task-chart';chart.style.position='relative';"
    h = h & "// Today line"
    h = h & "const todayPct=((today-minDate)/(maxDate-minDate))*100;"
    h = h & "if(todayPct>=0&&todayPct<=100){const tl=document.createElement('div');tl.className='today-line';tl.style.left=todayPct+'%';chart.appendChild(tl);}"
    h = h & "if(t.start&&t.end&&t.type!=='Phase'){"
    h = h & "const s=parseDate(t.start),e=parseDate(t.end);"
    h = h & "if(s&&e){const left=((s-minDate)/(maxDate-minDate))*100;const width=((e-s)/(maxDate-minDate))*100;"
    h = h & "const bar=document.createElement('div');bar.className='gantt-bar';"
    h = h & "bar.style.left=Math.max(0,left)+'%';bar.style.width=Math.max(.5,width)+'%';"
    h = h & "bar.style.background=STATUS_COLORS[t.status]||'#666';"
    h = h & "const fill=document.createElement('div');fill.className='fill';fill.style.width=(t.pct*100)+'%';bar.appendChild(fill);"
    h = h & "const lbl=document.createElement('span');lbl.style.position='relative';lbl.style.zIndex='1';"
    h = h & "lbl.textContent=Math.round(t.pct*100)+'%';bar.appendChild(lbl);"
    h = h & "bar.addEventListener('mousemove',e=>{showTooltip(e,t);});bar.addEventListener('mouseleave',hideTooltip);"
    h = h & "chart.appendChild(bar);}}"
    h = h & "row.appendChild(chart);container.appendChild(row);});"
    h = h & "}"
    h = h & "function showTooltip(e,t){"
    h = h & "const tt=document.getElementById('tooltip');"
    h = h & "tt.innerHTML=`<strong>${t.desc}</strong><div class='tt-row'><span>Status</span><span class='status-badge stat-${(t.status||'').replace(' ','')}'>${t.status}</span></div><div class='tt-row'><span>Resource</span><span>${t.resource}</span></div><div class='tt-row'><span>Start</span><span>${t.start}</span></div><div class='tt-row'><span>End</span><span>${t.end}</span></div><div class='tt-row'><span>Progress</span><span>${Math.round(t.pct*100)}%</span></div>${t.notes?'<div class=\"tt-row\"><span>Notes</span><span>'+t.notes+'</span></div>':''}`;"
    h = h & "tt.style.display='block';tt.style.left=(e.clientX+12)+'px';tt.style.top=(e.clientY-10)+'px';}"
    h = h & "function hideTooltip(){document.getElementById('tooltip').style.display='none';}"
    h = h & "function filterTasks(){"
    h = h & "const q=document.getElementById('searchBox').value.toLowerCase();"
    h = h & "const s=document.getElementById('statusFilter').value;"
    h = h & "const r=document.getElementById('resourceFilter').value;"
    h = h & "document.querySelectorAll('.gantt-row').forEach(row=>{"
    h = h & "const matchQ=!q||row.dataset.desc.includes(q);"
    h = h & "const matchS=!s||row.dataset.status===s;"
    h = h & "const matchR=!r||row.dataset.resource===r;"
    h = h & "row.style.display=(matchQ&&matchS&&matchR)||row.classList.contains('phase-row')?'':'none';});}"
    h = h & "// Populate resource filter"
    h = h & "const resources=[...new Set(TASKS.filter(t=>t.resource).map(t=>t.resource))];"
    h = h & "const rf=document.getElementById('resourceFilter');"
    h = h & "resources.forEach(r=>{const o=document.createElement('option');o.value=r;o.textContent=r;rf.appendChild(o);});"
    h = h & "buildGantt(TASKS);"
    h = h & "</script></body></html>"
    BuildGanttHTML = h
End Function

'=============================================================================
' EXPORT 2: PROJECT DASHBOARD
'=============================================================================
Sub ExportDashboard()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(PLAN_SHEET)

    Dim projectName As String
    projectName = ws.Range("E2").Value
    If projectName = "" Then projectName = "Project Dashboard"

    Dim tasks As Variant
    tasks = CollectTasks()

    If Not IsArray(tasks) Then
        MsgBox "No task data found.", vbExclamation
        Exit Sub
    End If

    Dim jsonTasks As String
    jsonTasks = "["
    Dim i As Long
    For i = 1 To UBound(tasks, 1)
        If i > 1 Then jsonTasks = jsonTasks & ","
        Dim pct As Double
        If IsNumeric(tasks(i, 8)) Then pct = tasks(i, 8) Else pct = 0
        jsonTasks = jsonTasks & "{"
        jsonTasks = jsonTasks & """id"":" & IIf(tasks(i, 1) = "", "0", tasks(i, 1)) & ","
        jsonTasks = jsonTasks & """type"":""" & EscapeJson(CStr(tasks(i, 2))) & ""","
        jsonTasks = jsonTasks & """phase"":""" & EscapeJson(CStr(tasks(i, 3))) & ""","
        jsonTasks = jsonTasks & """desc"":""" & EscapeJson(CStr(tasks(i, 4))) & ""","
        jsonTasks = jsonTasks & """start"":""" & tasks(i, 5) & ""","
        jsonTasks = jsonTasks & """end"":""" & tasks(i, 6) & ""","
        jsonTasks = jsonTasks & """resource"":""" & EscapeJson(CStr(tasks(i, 7))) & ""","
        jsonTasks = jsonTasks & """pct"":" & pct & ","
        jsonTasks = jsonTasks & """status"":""" & EscapeJson(CStr(tasks(i, 9))) & ""","
        jsonTasks = jsonTasks & """notes"":""" & EscapeJson(CStr(tasks(i, 10))) & """"
        jsonTasks = jsonTasks & "}"
    Next i
    jsonTasks = jsonTasks & "]"

    Dim html As String
    html = BuildDashboardHTML(projectName, jsonTasks)

    Dim savePath As String
    savePath = GetSavePath("ProjectDashboard.html")
    WriteFile savePath, html
    OpenInBrowser savePath
    MsgBox "Dashboard exported to:" & vbCrLf & savePath, vbInformation, "Export Complete"
End Sub

Function BuildDashboardHTML(projectName As String, jsonTasks As String) As String
    Dim h As String
    h = "<!DOCTYPE html><html lang=""en""><head><meta charset=""UTF-8"">"
    h = h & "<title>Dashboard - " & projectName & "</title>"
    h = h & "<script src=""https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js""></script>"
    h = h & "<style>"
    h = h & "*{box-sizing:border-box;margin:0;padding:0;font-family:'Segoe UI',Arial,sans-serif}"
    h = h & "body{background:#0f1923;color:#dce8f0;min-height:100vh}"
    h = h & "header{background:linear-gradient(135deg,#1f4e79 0%,#2e75b6 100%);padding:22px 32px;box-shadow:0 3px 15px rgba(0,0,0,.4)}"
    h = h & "header h1{font-size:1.7rem;font-weight:700}header p{opacity:.75;font-size:.9rem;margin-top:4px}"
    h = h & ".kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:16px;padding:24px 32px}"
    h = h & ".kpi{background:#162636;border-radius:12px;padding:20px;text-align:center;border:1px solid #2e4a63;transition:transform .2s}"
    h = h & ".kpi:hover{transform:translateY(-3px)}"
    h = h & ".kpi .num{font-size:2.4rem;font-weight:800;line-height:1}"
    h = h & ".kpi .lbl{font-size:.82rem;color:#7bafc8;margin-top:6px;text-transform:uppercase;letter-spacing:.5px}"
    h = h & ".charts-grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:20px;padding:0 32px 24px}"
    h = h & "@media(max-width:900px){.charts-grid{grid-template-columns:1fr 1fr}}"
    h = h & "@media(max-width:600px){.charts-grid{grid-template-columns:1fr}}"
    h = h & ".chart-card{background:#162636;border-radius:12px;padding:20px;border:1px solid #2e4a63}"
    h = h & ".chart-card h3{font-size:1rem;margin-bottom:14px;color:#93b5d0;font-weight:600}"
    h = h & ".phase-progress{padding:0 32px 32px}"
    h = h & ".phase-card{background:#162636;border-radius:10px;padding:16px 20px;margin-bottom:10px;border:1px solid #2e4a63}"
    h = h & ".phase-name{font-weight:700;margin-bottom:8px;font-size:.92rem}"
    h = h & ".phase-bar-bg{background:#0d1b2a;border-radius:6px;height:14px;overflow:hidden}"
    h = h & ".phase-bar{height:100%;border-radius:6px;background:linear-gradient(90deg,#00b050,#00e065);transition:width .5s}"
    h = h & ".phase-meta{display:flex;justify-content:space-between;font-size:.78rem;color:#7bafc8;margin-top:4px}"
    h = h & ".task-table-wrap{padding:0 32px 32px;overflow-x:auto}"
    h = h & "table{width:100%;border-collapse:collapse;font-size:.85rem}"
    h = h & "th{background:#1f4e79;color:#fff;padding:10px 12px;text-align:left;font-weight:600}"
    h = h & "td{padding:8px 12px;border-bottom:1px solid #1e3550}"
    h = h & "tr:hover td{background:rgba(46,117,182,.12)}"
    h = h & ".badge{display:inline-block;padding:2px 10px;border-radius:10px;font-size:.75rem;font-weight:700}"
    h = h & ".bg-green{background:#00b050;color:#fff}.bg-amber{background:#ffb900;color:#000}"
    h = h & ".bg-grey{background:#a5a5a5;color:#fff}.bg-red{background:#ff0000;color:#fff}.bg-purple{background:#7030a0;color:#fff}"
    h = h & ".progress-mini{width:80px;height:8px;background:#0d1b2a;border-radius:4px;overflow:hidden;display:inline-block;vertical-align:middle;margin-left:6px}"
    h = h & ".progress-mini-fill{height:100%;background:#2e75b6;border-radius:4px}"
    h = h & ".footer{text-align:center;padding:20px;color:#3a6080;font-size:.8rem;border-top:1px solid #1e3550}"
    h = h & "</style></head><body>"
    h = h & "<header><h1>&#128202; " & projectName & " &mdash; Project Dashboard</h1>"
    h = h & "<p>Generated: " & Format(Now(), "DD MMM YYYY HH:MM") & "</p></header>"
    h = h & "<div class=""kpi-grid"">"
    h = h & "<div class=""kpi""><div class=""num"" id=""kpiTotal"">-</div><div class=""lbl"">Total Tasks</div></div>"
    h = h & "<div class=""kpi"" style=""border-color:#00b050""><div class=""num"" style=""color:#00b050"" id=""kpiComplete"">-</div><div class=""lbl"">Completed</div></div>"
    h = h & "<div class=""kpi"" style=""border-color:#ffb900""><div class=""num"" style=""color:#ffb900"" id=""kpiInProg"">-</div><div class=""lbl"">In Progress</div></div>"
    h = h & "<div class=""kpi"" style=""border-color:#ff0000""><div class=""num"" style=""color:#ff0000"" id=""kpiOverdue"">-</div><div class=""lbl"">Overdue</div></div>"
    h = h & "<div class=""kpi"" style=""border-color:#7030a0""><div class=""num"" style=""color:#7030a0"" id=""kpiBlocked"">-</div><div class=""lbl"">Blocked</div></div>"
    h = h & "<div class=""kpi"" style=""border-color:#2e75b6""><div class=""num"" style=""color:#2e75b6"" id=""kpiPct"">-</div><div class=""lbl"">% Complete</div></div>"
    h = h & "</div>"
    h = h & "<div class=""charts-grid"">"
    h = h & "<div class=""chart-card""><h3>Status Breakdown</h3><canvas id=""statusChart"" height=""200""></canvas></div>"
    h = h & "<div class=""chart-card""><h3>Resource Workload</h3><canvas id=""resourceChart"" height=""200""></canvas></div>"
    h = h & "<div class=""chart-card""><h3>Phase Completion</h3><canvas id=""phaseChart"" height=""200""></canvas></div>"
    h = h & "</div>"
    h = h & "<div class=""phase-progress""><h3 style=""color:#93b5d0;margin-bottom:12px;font-size:1rem"">Phase Progress Breakdown</h3><div id=""phaseCards""></div></div>"
    h = h & "<div class=""task-table-wrap""><h3 style=""color:#93b5d0;margin-bottom:12px;font-size:1rem"">All Tasks</h3>"
    h = h & "<table><thead><tr><th>ID</th><th>Type</th><th>Description</th><th>Resource</th><th>Start</th><th>End</th><th>Progress</th><th>Status</th><th>Notes</th></tr></thead>"
    h = h & "<tbody id=""taskTableBody""></tbody></table></div>"
    h = h & "<div class=""footer"">&#128202; Project Dashboard &mdash; " & projectName & " &mdash; " & Format(Now(), "DD MMM YYYY") & "</div>"
    h = h & "<script>"
    h = h & "const TASKS=" & jsonTasks & ";"
    h = h & "const STATUS_COLORS={'Completed':'#00b050','In Progress':'#ffb900','Not Started':'#a5a5a5','Overdue':'#ff0000','Blocked':'#7030a0'};"
    h = h & "const workerTasks=TASKS.filter(t=>t.type==='Task'||t.type==='Subtask');"
    h = h & "const total=workerTasks.length;"
    h = h & "const byStatus={};"
    h = h & "workerTasks.forEach(t=>{byStatus[t.status]=(byStatus[t.status]||0)+1;});"
    h = h & "const completed=byStatus['Completed']||0;"
    h = h & "document.getElementById('kpiTotal').textContent=total;"
    h = h & "document.getElementById('kpiComplete').textContent=completed;"
    h = h & "document.getElementById('kpiInProg').textContent=byStatus['In Progress']||0;"
    h = h & "document.getElementById('kpiOverdue').textContent=byStatus['Overdue']||0;"
    h = h & "document.getElementById('kpiBlocked').textContent=byStatus['Blocked']||0;"
    h = h & "document.getElementById('kpiPct').textContent=total>0?Math.round(completed/total*100)+'%':'0%';"
    h = h & "// Status chart"
    h = h & "new Chart(document.getElementById('statusChart'),{type:'doughnut',data:{labels:Object.keys(byStatus),datasets:[{data:Object.values(byStatus),backgroundColor:Object.keys(byStatus).map(s=>STATUS_COLORS[s]||'#666'),borderWidth:0}]},options:{plugins:{legend:{labels:{color:'#dce8f0'}}},cutout:'65%'}});"
    h = h & "// Resource chart"
    h = h & "const byResource={};"
    h = h & "workerTasks.forEach(t=>{if(t.resource)byResource[t.resource]=(byResource[t.resource]||0)+1;});"
    h = h & "const sortedRes=Object.entries(byResource).sort((a,b)=>b[1]-a[1]).slice(0,10);"
    h = h & "new Chart(document.getElementById('resourceChart'),{type:'bar',data:{labels:sortedRes.map(r=>r[0]),datasets:[{label:'Tasks',data:sortedRes.map(r=>r[1]),backgroundColor:'#2e75b6',borderRadius:4}]},options:{indexAxis:'y',plugins:{legend:{display:false}},scales:{x:{ticks:{color:'#7bafc8'},grid:{color:'#1e3550'}},y:{ticks:{color:'#dce8f0'},grid:{color:'transparent'}}}}});"
    h = h & "// Phase chart & cards"
    h = h & "const phases=[...new Set(TASKS.filter(t=>t.type==='Phase').map(t=>t.phase))];"
    h = h & "const phaseComps=phases.map(p=>{const pt=TASKS.filter(t=>t.phase===p&&(t.type==='Task'||t.type==='Subtask'));return pt.length>0?Math.round(pt.filter(t=>t.status==='Completed').length/pt.length*100):0;});"
    h = h & "new Chart(document.getElementById('phaseChart'),{type:'bar',data:{labels:phases.map(p=>p.replace(/Phase \\d+: /,'')),datasets:[{label:'% Done',data:phaseComps,backgroundColor:phaseComps.map(v=>v===100?'#00b050':v>50?'#ffb900':'#2e75b6'),borderRadius:4}]},options:{plugins:{legend:{display:false}},scales:{y:{max:100,ticks:{callback:v=>v+'%',color:'#7bafc8'},grid:{color:'#1e3550'}},x:{ticks:{color:'#dce8f0',maxRotation:45},grid:{color:'transparent'}}}}});"
    h = h & "const phaseCards=document.getElementById('phaseCards');"
    h = h & "phases.forEach((p,i)=>{"
    h = h & "const pTasks=TASKS.filter(t=>t.phase===p&&(t.type==='Task'||t.type==='Subtask'));"
    h = h & "const pComp=pTasks.filter(t=>t.status==='Completed').length;"
    h = h & "const pPct=pTasks.length>0?Math.round(pComp/pTasks.length*100):0;"
    h = h & "const card=document.createElement('div');card.className='phase-card';"
    h = h & "card.innerHTML=`<div class='phase-name'>${p}</div><div class='phase-bar-bg'><div class='phase-bar' style='width:${pPct}%;background:${pPct===100?'linear-gradient(90deg,#00b050,#00e065)':pPct>50?'linear-gradient(90deg,#ffb900,#ffd000)':'linear-gradient(90deg,#2e75b6,#4a9fd6)'}'></div></div><div class='phase-meta'><span>${pComp} of ${pTasks.length} tasks complete</span><span>${pPct}%</span></div>`;"
    h = h & "phaseCards.appendChild(card);});"
    h = h & "// Task table"
    h = h & "const tbody=document.getElementById('taskTableBody');"
    h = h & "const statusBadge={'Completed':'bg-green','In Progress':'bg-amber','Not Started':'bg-grey','Overdue':'bg-red','Blocked':'bg-purple'};"
    h = h & "workerTasks.forEach(t=>{"
    h = h & "const tr=document.createElement('tr');"
    h = h & "const pct=Math.round((t.pct||0)*100);"
    h = h & "tr.innerHTML=`<td>${t.id||''}</td><td>${t.type}</td><td>${t.desc}</td><td>${t.resource}</td><td>${t.start}</td><td>${t.end}</td>`+"
    h = h & "`<td><span>${pct}%</span><span class='progress-mini'><span class='progress-mini-fill' style='width:${pct}%'></span></span></td>`+"
    h = h & "`<td><span class='badge ${statusBadge[t.status]||""}'>${t.status}</span></td><td style='color:#7bafc8'>${t.notes||''}</td>`;"
    h = h & "tbody.appendChild(tr);});"
    h = h & "</script></body></html>"
    BuildDashboardHTML = h
End Function

'=============================================================================
' EXPORT 3: RAG STATUS REPORT
'=============================================================================
Sub ExportRAGReport()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(PLAN_SHEET)

    Dim projectName As String
    projectName = ws.Range("E2").Value
    If projectName = "" Then projectName = "Project RAG Report"

    Dim tasks As Variant
    tasks = CollectTasks()

    If Not IsArray(tasks) Then
        MsgBox "No task data found.", vbExclamation
        Exit Sub
    End If

    Dim jsonTasks As String
    jsonTasks = "["
    Dim i As Long
    For i = 1 To UBound(tasks, 1)
        If i > 1 Then jsonTasks = jsonTasks & ","
        Dim pct As Double
        If IsNumeric(tasks(i, 8)) Then pct = tasks(i, 8) Else pct = 0
        jsonTasks = jsonTasks & "{"
        jsonTasks = jsonTasks & """id"":" & IIf(tasks(i, 1) = "", "0", tasks(i, 1)) & ","
        jsonTasks = jsonTasks & """type"":""" & EscapeJson(CStr(tasks(i, 2))) & ""","
        jsonTasks = jsonTasks & """phase"":""" & EscapeJson(CStr(tasks(i, 3))) & ""","
        jsonTasks = jsonTasks & """desc"":""" & EscapeJson(CStr(tasks(i, 4))) & ""","
        jsonTasks = jsonTasks & """start"":""" & tasks(i, 5) & ""","
        jsonTasks = jsonTasks & """end"":""" & tasks(i, 6) & ""","
        jsonTasks = jsonTasks & """resource"":""" & EscapeJson(CStr(tasks(i, 7))) & ""","
        jsonTasks = jsonTasks & """pct"":" & pct & ","
        jsonTasks = jsonTasks & """status"":""" & EscapeJson(CStr(tasks(i, 9))) & ""","
        jsonTasks = jsonTasks & """notes"":""" & EscapeJson(CStr(tasks(i, 10))) & """"
        jsonTasks = jsonTasks & "}"
    Next i
    jsonTasks = jsonTasks & "]"

    Dim html As String
    html = BuildRAGHTML(projectName, jsonTasks)

    Dim savePath As String
    savePath = GetSavePath("ProjectRAGReport.html")
    WriteFile savePath, html
    OpenInBrowser savePath
    MsgBox "RAG Status Report exported to:" & vbCrLf & savePath, vbInformation, "Export Complete"
End Sub

Function BuildRAGHTML(projectName As String, jsonTasks As String) As String
    Dim h As String
    h = "<!DOCTYPE html><html lang=""en""><head><meta charset=""UTF-8"">"
    h = h & "<title>RAG Report - " & projectName & "</title>"
    h = h & "<style>"
    h = h & "*{box-sizing:border-box;margin:0;padding:0;font-family:'Segoe UI',Arial,sans-serif}"
    h = h & "body{background:#fff;color:#222;min-height:100vh}"
    h = h & "@media print{.no-print{display:none}body{background:#fff}}"
    h = h & "header{background:#1f4e79;color:#fff;padding:24px 40px;print-color-adjust:exact}"
    h = h & "header h1{font-size:1.8rem;font-weight:700}"
    h = h & "header .sub{display:flex;gap:24px;margin-top:8px;font-size:.9rem;opacity:.85}"
    h = h & ".rag-summary{display:flex;gap:0;margin:20px 40px;border-radius:12px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,.15)}"
    h = h & ".rag-band{flex:1;padding:20px;text-align:center;color:#fff;print-color-adjust:exact}"
    h = h & ".rag-band .num{font-size:2.5rem;font-weight:800;line-height:1}"
    h = h & ".rag-band .lbl{font-size:.85rem;text-transform:uppercase;letter-spacing:.5px;margin-top:4px}"
    h = h & ".rag-R{background:#d00000}.rag-A{background:#e89b00}.rag-G{background:#00863f}"
    h = h & ".section-title{font-size:1.1rem;font-weight:700;margin:24px 40px 10px;color:#1f4e79;border-bottom:2px solid #1f4e79;padding-bottom:6px}"
    h = h & ".phase-section{margin:0 40px 20px}"
    h = h & ".phase-hdr{display:flex;align-items:center;gap:12px;padding:10px 16px;border-radius:8px 8px 0 0;font-weight:700;font-size:.95rem;cursor:pointer;print-color-adjust:exact}"
    h = h & ".phase-hdr.rag-R{background:#d00000;color:#fff}"
    h = h & ".phase-hdr.rag-A{background:#e89b00;color:#fff}"
    h = h & ".phase-hdr.rag-G{background:#00863f;color:#fff}"
    h = h & ".phase-hdr.rag-B{background:#1f4e79;color:#fff}"
    h = h & ".rag-dot{width:16px;height:16px;border-radius:50%;border:2px solid rgba(255,255,255,.5)}"
    h = h & ".phase-body{border:1px solid #ddd;border-top:none;border-radius:0 0 8px 8px;overflow:hidden}"
    h = h & "table{width:100%;border-collapse:collapse;font-size:.85rem}"
    h = h & "th{background:#f0f5fb;color:#1f4e79;padding:8px 12px;text-align:left;font-weight:600;border-bottom:2px solid #1f4e79}"
    h = h & "td{padding:7px 12px;border-bottom:1px solid #eee}"
    h = h & "tr:last-child td{border-bottom:none}"
    h = h & "tr.subtask td:nth-child(3){padding-left:24px;font-style:italic;color:#555}"
    h = h & ".rag-pill{display:inline-block;padding:3px 12px;border-radius:12px;font-size:.75rem;font-weight:700;color:#fff;print-color-adjust:exact}"
    h = h & ".pill-R{background:#d00000}.pill-A{background:#e89b00;color:#000}.pill-G{background:#00863f}.pill-B{background:#1f4e79}"
    h = h & ".pbar{width:80px;height:8px;background:#eee;border-radius:4px;display:inline-block;vertical-align:middle;margin-left:6px;overflow:hidden}"
    h = h & ".pbar-fill{height:100%;border-radius:4px;print-color-adjust:exact}"
    h = h & ".pbar-fill.R{background:#d00000}.pbar-fill.A{background:#e89b00}.pbar-fill.G{background:#00863f}"
    h = h & ".notes-cell{color:#e65c00;font-style:italic;font-size:.82rem}"
    h = h & ".btn-print{position:fixed;bottom:24px;right:24px;background:#1f4e79;color:#fff;border:none;padding:12px 22px;border-radius:8px;cursor:pointer;font-size:.9rem;font-weight:600;box-shadow:0 4px 12px rgba(0,0,0,.25)}"
    h = h & ".btn-print:hover{background:#2e75b6}"
    h = h & ".footer{text-align:center;padding:20px;color:#888;font-size:.8rem;border-top:1px solid #eee;margin-top:20px}"
    h = h & ".overall-rag{display:inline-flex;align-items:center;gap:10px;padding:10px 20px;border-radius:8px;font-weight:700;font-size:1.1rem;margin:16px 40px;color:#fff;print-color-adjust:exact}"
    h = h & "</style></head><body>"
    h = h & "<header>"
    h = h & "<h1>&#128994;&#128993;&#128308; RAG Status Report &mdash; " & projectName & "</h1>"
    h = h & "<div class=""sub""><span>&#128197; Generated: " & Format(Now(), "DD MMM YYYY HH:MM") & "</span>"
    h = h & "<span>&#128203; This report reflects live data from the project plan workbook.</span></div>"
    h = h & "</header>"
    h = h & "<div id=""overallRAG""></div>"
    h = h & "<div class=""rag-summary""><div class=""rag-band rag-R""><div class=""num"" id=""cntR"">-</div><div class=""lbl"">&#128308; Red — Overdue / Blocked</div></div>"
    h = h & "<div class=""rag-band rag-A""><div class=""num"" id=""cntA"">-</div><div class=""lbl"">&#128993; Amber — In Progress</div></div>"
    h = h & "<div class=""rag-band rag-G""><div class=""num"" id=""cntG"">-</div><div class=""lbl"">&#128994; Green — Completed</div></div></div>"
    h = h & "<div id=""phaseSections""></div>"
    h = h & "<button class=""btn-print no-print"" onclick=""window.print()"">&#128424; Print / Save PDF</button>"
    h = h & "<div class=""footer"">RAG Status Report &mdash; " & projectName & " &mdash; Confidential &mdash; " & Format(Now(), "DD MMM YYYY") & "</div>"
    h = h & "<script>"
    h = h & "const TASKS=" & jsonTasks & ";"
    h = h & "function getRAG(status){"
    h = h & "if(status==='Overdue'||status==='Blocked')return'R';"
    h = h & "if(status==='In Progress'||status==='Not Started')return'A';"
    h = h & "if(status==='Completed')return'G';"
    h = h & "return'B';}"
    h = h & "const phases=[...new Set(TASKS.filter(t=>t.type==='Phase').map(t=>t.phase))];"
    h = h & "const workerTasks=TASKS.filter(t=>t.type==='Task'||t.type==='Subtask');"
    h = h & "let cntR=0,cntA=0,cntG=0;"
    h = h & "workerTasks.forEach(t=>{const r=getRAG(t.status);if(r==='R')cntR++;else if(r==='A')cntA++;else if(r==='G')cntG++;});"
    h = h & "document.getElementById('cntR').textContent=cntR;"
    h = h & "document.getElementById('cntA').textContent=cntA;"
    h = h & "document.getElementById('cntG').textContent=cntG;"
    h = h & "const total=workerTasks.length;"
    h = h & "const overallRag=cntR>0?'R':cntA>0?'A':'G';"
    h = h & "const ragLabel={'R':'RED — Project has critical issues requiring immediate attention','A':'AMBER — Project in progress, some items need monitoring','G':'GREEN — Project on track, all tasks completed or progressing well'};"
    h = h & "const ragBg={'R':'#d00000','A':'#e89b00','G':'#00863f'};"
    h = h & "document.getElementById('overallRAG').innerHTML=`<div class='overall-rag' style='background:${ragBg[overallRag]}'><span style='font-size:1.4rem'>${overallRag==='R'?'&#128308;':overallRag==='A'?'&#128993;':'&#128994;'}</span> OVERALL STATUS: ${ragLabel[overallRag]}</div>`;"
    h = h & "const container=document.getElementById('phaseSections');"
    h = h & "container.innerHTML='<div class=""section-title"">Phase-by-Phase RAG Breakdown</div>';"
    h = h & "phases.forEach(phase=>{"
    h = h & "const pTasks=TASKS.filter(t=>t.phase===phase&&(t.type==='Task'||t.type==='Subtask'));"
    h = h & "const pRags=pTasks.map(t=>getRAG(t.status));"
    h = h & "const phaseRag=pRags.includes('R')?'R':pRags.includes('A')?'A':pRags.length>0?'G':'B';"
    h = h & "const ragIcon={'R':'&#128308;','A':'&#128993;','G':'&#128994;','B':'&#9898;'};"
    h = h & "const phaseSection=document.createElement('div');phaseSection.className='phase-section';"
    h = h & "let tableRows='';"
    h = h & "pTasks.forEach((t,i)=>{"
    h = h & "const tRag=getRAG(t.status);"
    h = h & "const pct=Math.round((t.pct||0)*100);"
    h = h & "const barColor=tRag==='R'?'R':tRag==='A'?'A':'G';"
    h = h & "tableRows+=`<tr class='${t.type===\"Subtask\"?\"subtask\":\"\"}'>`+"
    h = h & "`<td>${t.id||''}</td><td>${t.type}</td><td>${t.desc}</td><td>${t.resource}</td>`+"
    h = h & "`<td>${t.start}</td><td>${t.end}</td>`+"
    h = h & "`<td>${pct}%<span class='pbar'><span class='pbar-fill ${barColor}' style='width:${pct}%'></span></span></td>`+"
    h = h & "`<td><span class='rag-pill pill-${tRag}'>${tRag==='R'?'&#128308; RED':tRag==='A'?'&#128993; AMBER':'&#128994; GREEN'}</span></td>`+"
    h = h & "`<td class='notes-cell'>${t.notes||''}</td></tr>`;"
    h = h & "});"
    h = h & "phaseSection.innerHTML=`<div class='phase-hdr rag-${phaseRag}' onclick='this.nextElementSibling.style.display=this.nextElementSibling.style.display===\"none\"?\"\":\"none\"'>`+"
    h = h & "`<span>${ragIcon[phaseRag]}</span><span>${phase}</span>`+"
    h = h & "`<span style='margin-left:auto;font-size:.8rem;opacity:.8'>${pTasks.length} tasks &mdash; click to collapse</span></div>`+"
    h = h & "`<div class='phase-body'><table><thead><tr><th>ID</th><th>Type</th><th>Description</th><th>Resource</th><th>Start</th><th>End</th><th>Progress</th><th>RAG</th><th>Notes / Blockers</th></tr></thead>`+"
    h = h & "`<tbody>${tableRows}</tbody></table></div>`;"
    h = h & "container.appendChild(phaseSection);});"
    h = h & "</script></body></html>"
    BuildRAGHTML = h
End Function
