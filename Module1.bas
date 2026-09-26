Attribute VB_Name = "Module1"
Option Explicit

Public Sub SendCalibrationAlerts()

    Const RECIPIENTS As String = _
        "xihao.han@mevion.com;yiqian.lv@mevion.com;bin.yu@mevion.com;peng.cui@mevion.com;wendong.tian@mevion.com"
    Const ALERT_COOLDOWN As Long = 7

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
    Dim toolId As String
    Dim toolName As String
    Dim location As String
    Dim nextCalDate As Variant
    Dim lastAlertStr As String

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Worksheets("Calibrated Tools Register")

    ' Refresh TODAY() and all formulas before checking.
    Application.CalculateFull
    DoEvents

    ' Find the last row with data in column B (Tool ID)
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    If lastRow < 4 Then lastRow = 4

    body = "<html><body>" & _
           "<p>The following calibration tasks have status <strong>Expiring Soon</strong> " & _
           "and are eligible for alert (no alert sent in the past " & ALERT_COOLDOWN & _
           " days or never alerted):</p>" & _
           "<table border='1' cellpadding='5' cellspacing='0'>" & _
           "<tr style='background:#1F4E78;color:white;'>" & _
           "<th>No.</th><th>Tool ID</th><th>Tool Name</th>" & _
           "<th>Location</th><th>Next Calibration Date</th>" & _
           "<th>Status</th><th>Last Alert Sent</th></tr>"

    alertCount = 0

    For i = 4 To lastRow

        ' Get Calibration Status (column N)
        statusVal = Trim(CStr(ws.Cells(i, 14).value))

        ' Only process rows where status = "Expiring Soon"
        If statusVal = "Expiring Soon" Then

            ' Get Last Alert date (column R)
            lastAlert = ws.Cells(i, 18).value

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

                ' Gather cell values for email body
                toolId = CStr(ws.Cells(i, 2).value)
                toolName = CStr(ws.Cells(i, 3).value)
                location = CStr(ws.Cells(i, 9).value)
                nextCalDate = ws.Cells(i, 13).value

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
                    "<td>" & HtmlEncode(toolId) & "</td>" & _
                    "<td>" & HtmlEncode(toolName) & "</td>" & _
                    "<td>" & HtmlEncode(location) & "</td>" & _
                    "<td>" & Format(nextCalDate, "yyyy-mm-dd") & "</td>" & _
                    "<td style='color:red;font-weight:bold;'>" & _
                    HtmlEncode(statusVal) & "</td>" & _
                    "<td>" & Format(Date, "yyyy-mm-dd") & "</td></tr>"

                ' Record today's date as last alert sent
                ws.Cells(i, 18).value = Date
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


