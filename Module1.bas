Option Explicit

Public Sub SendCalibrationAlerts()

    Const RECIPIENTS As String = _
        "xihao.han@mevion.com;yiqian.lv@mevion.com;bin.yu@mevion.com;peng.cui@mevion.com;wendong.tian@mevion.com"
    Const ALERT_COOLDOWN As Long = 7
    Const HEADER_ROW As Long = 3
    Const DATA_START_ROW As Long = 4

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim outlookApp As Object
    Dim email As Object
    Dim statusVal As String
    Dim lastAlert As Variant
    Dim daysSinceAlert As Long
    Dim body As String
    Dim alertCount As Long
    Dim partNo As String       ' Renamed from toolId
    Dim toolName As String
    Dim location As String
    Dim nextCalDate As Variant
    Dim lastAlertStr As String

    ' Dynamic Column Variables
    Dim colPartNo As Long      ' Renamed from colToolID
    Dim colToolName As Long
    Dim colLocation As Long
    Dim colStatus As Long
    Dim colNextCal As Long
    Dim colLastAlert As Long

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Worksheets("Calibrated Tools Register")

    ' Refresh TODAY() and all formulas before checking.
    Application.CalculateFull
    DoEvents

    ' --- 1. Identify Column Indexes Dynamically ---
    ' We use Application.Match to find the column number based on the header text in Row 3.
    ' If a header is missing, the macro will alert you and stop.
    
    On Error Resume Next ' Suppress error if Match fails so we can check it manually
    
    colPartNo = Application.Match("Part No.", ws.Rows(HEADER_ROW), 0)
    colToolName = Application.Match("Tool Name", ws.Rows(HEADER_ROW), 0)
    colLocation = Application.Match("Location", ws.Rows(HEADER_ROW), 0)
    colStatus = Application.Match("Calibration Status", ws.Rows(HEADER_ROW), 0)
    colNextCal = Application.Match("Next Calibration Date", ws.Rows(HEADER_ROW), 0)
    colLastAlert = Application.Match("Last Alert Sent", ws.Rows(HEADER_ROW), 0)
    
    On Error GoTo ErrorHandler ' Re-enable standard error handling

    ' Check if all columns were found
    If IsError(colPartNo) Or IsError(colToolName) Or IsError(colLocation) Or _
       IsError(colStatus) Or IsError(colNextCal) Or IsError(colLastAlert) Then
        MsgBox "Error: One or more required headers were not found in Row " & HEADER_ROW & "." & vbCrLf & _
               "Please ensure the following headers exist exactly as spelled:" & vbCrLf & _
               "- Part No." & vbCrLf & _
               "- Tool Name" & vbCrLf & _
               "- Location" & vbCrLf & _
               "- Calibration Status" & vbCrLf & _
               "- Next Calibration Date" & vbCrLf & _
               "- Last Alert Sent", vbCritical
        Exit Sub
    End If

    ' --- 2. Find Last Row ---
    ' Using the dynamic Part No. column to find the bottom of the data
    lastRow = ws.Cells(ws.Rows.Count, colPartNo).End(xlUp).Row
    If lastRow < DATA_START_ROW Then lastRow = DATA_START_ROW

    ' --- 3. Build Email Body ---
    body = "<html><body>" & _
           "<p>The following calibration tasks have status <strong>Expiring Soon</strong> " & _
           "and are eligible for alert (no alert sent in the past " & ALERT_COOLDOWN & _
           " days or never alerted):</p>" & _
           "<table border='1' cellpadding='5' cellspacing='0'>" & _
           "<tr style='background:#1F4E78;color:white;'>" & _
           "<th>No.</th><th>Part No.</th><th>Tool Name</th>" & _
           "<th>Location</th><th>Next Calibration Date</th>" & _
           "<th>Status</th><th>Last Alert Sent</th></tr>"

    alertCount = 0

    For i = DATA_START_ROW To lastRow

        ' Get Calibration Status (Dynamic Column)
        statusVal = Trim(CStr(ws.Cells(i, colStatus).value))

        ' Only process rows where status = "Expiring Soon"
        If statusVal = "Expiring Soon" Then

            ' Get Last Alert date (Dynamic Column)
            lastAlert = ws.Cells(i, colLastAlert).value

            ' Check if alert should be sent:
            ' - If no last alert date recorded (empty/null), OR
            ' - If last alert date is >= ALERT_COOLDOWN days ago
            If IsEmpty(lastAlert) Or IsNull(lastAlert) Or _
               CStr(lastAlert) = "" Or CStr(lastAlert) = "0" Then
                daysSinceAlert = ALERT_COOLDOWN + 1  ' Force alert (never alerted)
            ElseIf IsDate(lastAlert) Then
                daysSinceAlert = DateDiff("d", CDate(lastAlert), Date)
                If daysSinceAlert < 0 Then daysSinceAlert = 0
            Else
                daysSinceAlert = ALERT_COOLDOWN + 1  ' Invalid date, force alert
            End If

            If daysSinceAlert >= ALERT_COOLDOWN Then

                ' Gather cell values for email body using Dynamic Columns
                partNo = CStr(ws.Cells(i, colPartNo).value)
                toolName = CStr(ws.Cells(i, colToolName).value)
                location = CStr(ws.Cells(i, colLocation).value)
                nextCalDate = ws.Cells(i, colNextCal).value

                ' Format last alert string for display
                If IsEmpty(lastAlert) Or IsNull(lastAlert) Or _
                   CStr(lastAlert) = "" Or CStr(lastAlert) = "0" Then
                    lastAlertStr = "Never"
                ElseIf IsDate(lastAlert) Then
                    lastAlertStr = Format(CDate(lastAlert), "yyyy-mm-dd")
                Else
                    lastAlertStr = CStr(lastAlert)
                End If

                alertCount = alertCount + 1

                body = body & "<tr>" & _
                    "<td>" & alertCount & "</td>" & _
                    "<td>" & HtmlEncode(partNo) & "</td>" & _
                    "<td>" & HtmlEncode(toolName) & "</td>" & _
                    "<td>" & HtmlEncode(location) & "</td>" & _
                    "<td>" & Format(nextCalDate, "yyyy-mm-dd") & "</td>" & _
                    "<td style='color:red;font-weight:bold;'>" & _
                    HtmlEncode(statusVal) & "</td>" & _
                    "<td>" & Format(Date, "yyyy-mm-dd") & "</td></tr>"

                ' Record today's date as last alert sent (Dynamic Column)
                ws.Cells(i, colLastAlert).value = Date
            End If
        End If
    Next i

    If alertCount = 0 Then
        MsgBox "No tools found with status ""Expiring Soon"" that are eligible for alert." & _
               vbCrLf & vbCrLf & _
               "Note: Alerts are only sent if no alert was sent in the past " & _
               ALERT_COOLDOWN & " days.", vbInformation
        ' Still save to persist any Last Alert dates written above
        ThisWorkbook.Save
        Exit Sub
    End If

    body = body & "</table>" & _
           "<p>Please review and arrange the required calibration.</p>" & _
           "</body></html>"

    ' Connect to Classic Outlook.
    On Error Resume Next
    Set outlookApp = GetObject(, "Outlook.Application")
    If outlookApp Is Nothing Then
        Set outlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo ErrorHandler

    If outlookApp Is Nothing Then
        Err.Raise vbObjectError + 1000, , _
            "Classic Outlook could not be started."
    End If

    Set email = outlookApp.CreateItem(0)

    With email
        .To = RECIPIENTS
        .Subject = "Calibration Alert - Expiring Soon - " & Format(Date, "yyyy-mm-dd")
        .HTMLBody = body
        .Send
        ' Use .Display instead of .Send while testing.
    End With

    ' Add delay for Outlook to process
    Application.Wait (Now + TimeValue("0:00:30"))

    ' Save workbook to persist Last Alert dates
    ThisWorkbook.Save

    MsgBox "Calibration alert email sent! Total items: " & alertCount, vbInformation
    Exit Sub

ErrorHandler:
    Debug.Print "Calibration alert error: " & Err.Number & " - " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical

End Sub

Private Function HtmlEncode(ByVal value As String) As String
    value = Replace(value, "&", "&amp;")
    value = Replace(value, "<", "&lt;")
    value = Replace(value, ">", "&gt;")
    value = Replace(value, """", "&quot;")
    HtmlEncode = value
End Function


