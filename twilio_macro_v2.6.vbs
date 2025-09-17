' VBA Macro for sending Twilio WhatsApp Messages (Hybrid-Variable Version)
' Version   : 2.6 (Final)
'---------------------------------------------------------------------------------------
' Procedure : SendTwilioMessages
' Author    : Jules
' Date      : 2025-09-17
' Purpose   : Main subroutine to send messages using Twilio's Content API.
'             Supports a hybrid variable model and rate-limiting.
'---------------------------------------------------------------------------------------
Public Sub SendTwilioMessages()
10    On Error GoTo ErrorHandler

20    Dim wsDashboard As Worksheet, wsContacts As Worksheet, wsTemplates As Worksheet, wsSettings As Worksheet, wsLog As Worksheet
30    Dim btnSend As Object, originalButtonState As Boolean
40    Dim accountSid As String, authToken As String, fromNumber As String
50    Dim messageUid As String, contentSid As String, mps As Long
60    Dim templateRow As Variant, baseVariables(1 To 15) As String
70    Dim http As Object
80    Dim lastRow As Long, i As Long, j As Long, k As Long
90    Dim messagesSentThisSecond As Long

    ' --- Configuration ---
100   Set wsDashboard = ThisWorkbook.Sheets("Dashboard")
110   Set wsContacts = ThisWorkbook.Sheets("Contacts")
120   Set wsTemplates = ThisWorkbook.Sheets("Templates")
130   Set wsSettings = ThisWorkbook.Sheets("Settings")
140   Set wsLog = ThisWorkbook.Sheets("Log")

    ' --- Disable Button ---
150   On Error Resume Next
160   Set btnSend = wsDashboard.Shapes("btnSendMessages")
170   If Not btnSend Is Nothing Then
180       originalButtonState = btnSend.OLEFormat.Object.Enabled
190       btnSend.OLEFormat.Object.Enabled = False
200       btnSend.OLEFormat.Object.Caption = "Sending..."
210   End If
220   On Error GoTo ErrorHandler

    ' --- Read Settings & Inputs ---
230   accountSid = wsSettings.Range("B2").Value
240   authToken = wsSettings.Range("B3").Value
250   fromNumber = wsSettings.Range("B4").Value
260   messageUid = wsDashboard.Range("C5").Value

270   Dim mpsValue As Variant
280   mpsValue = wsDashboard.Range("C6").Value
290   If Not IsNumeric(mpsValue) Then
300       MsgBox "Error: 'Messages Per Second (MPS)' must be a valid number.", vbCritical, "Config Error"
310       GoTo Cleanup
320   End If
330   mps = CLng(mpsValue)

    ' --- Input Validation ---
340   If accountSid = "" Or authToken = "" Or fromNumber = "" Or messageUid = "" Or mps <= 0 Then
350       MsgBox "Error: Please ensure all settings and inputs on the Dashboard and Settings sheets are filled in correctly.", vbCritical, "Config Error"
360       GoTo Cleanup
370   End If

    ' --- Find Template and Get Base Variables ---
380   templateRow = Application.Match(messageUid, wsTemplates.Columns(1), 0)
390   If IsError(templateRow) Then
400       MsgBox "Error: The Message UID '" & messageUid & "' was not found in the 'Templates' sheet.", vbCritical, "Template Error"
410       GoTo Cleanup
420   End If

430   contentSid = wsTemplates.Cells(templateRow, 2).Value
440   If contentSid = "" Then
450       MsgBox "Error: The ContentSid for Message UID '" & messageUid & "' is empty.", vbCritical, "Template Error"
460       GoTo Cleanup
470   End If

480   For j = 1 To 15
490       baseVariables(j) = wsTemplates.Cells(templateRow, 5 + j).Value
500   Next j

    ' --- Main Loop ---
510   Set http = CreateObject("MSXML2.XMLHTTP")
520   lastRow = wsContacts.Cells(wsContacts.Rows.Count, "A").End(xlUp).Row
530   messagesSentThisSecond = 0

540   For i = 2 To lastRow
550       If UCase(wsContacts.Cells(i, 3).Value) = "YES" Or UCase(wsContacts.Cells(i, 3).Value) = "TRUE" Then
            ' Rate Limiting
560           If messagesSentThisSecond >= mps Then
570               Application.Wait (Now + TimeValue("0:00:01"))
580               messagesSentThisSecond = 0
590           End If

600           Dim toName As String, toNumber As String, statusCell As Range
610           toName = wsContacts.Cells(i, 1).Value
620           toNumber = wsContacts.Cells(i, 2).Value
630           Set statusCell = wsContacts.Cells(i, 4)
640           statusCell.Value = ""

650           If toNumber <> "" Then
660               Dim resolvedVariables(1 To 15) As String, finalJson As String
670               Dim url As String, body As String, statusText As String, responseText As String

                ' CORRECTED Two-Tiered Variable Replacement
680               For j = 1 To 15
690                   Dim tempVar As String
700                   tempVar = baseVariables(j) ' Start with the base variable
                    ' Check if the base variable itself is a placeholder for a contact variable
710                   If Left(tempVar, 2) = "{{" And Right(tempVar, 2) = "}}" Then
720                       If tempVar = "{{name}}" Then
730                           tempVar = toName
740                       Else
750                           For k = 1 To 15
760                               If tempVar = "{{contact_var_" & k & "}}" Then
770                                   tempVar = wsContacts.Cells(i, 4 + k).Value
780                                   Exit For
790                               End If
800                           Next k
810                       End If
820                   End If
830                   resolvedVariables(j) = tempVar
840               Next j

                ' Build JSON and API Body
850               finalJson = BuildContentVariablesJson(resolvedVariables)
860               url = "https://api.twilio.com/2010-04-01/Accounts/" & accountSid & "/Messages.json"
870               body = "To=" & UrlEncode("whatsapp:" & toNumber) & _
                      "&From=" & UrlEncode("whatsapp:" & fromNumber) & _
                      "&ContentSid=" & contentSid & _
                      "&ContentVariables=" & UrlEncode(finalJson)

                ' Send Request
880               http.Open "POST", url, False
890               http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
900               http.setRequestHeader "Authorization", "Basic " & Base64Encode(accountSid & ":" & authToken)
910               http.send body
920               messagesSentThisSecond = messagesSentThisSecond + 1

                ' Process Response
930               If http.Status >= 200 And http.Status < 300 Then
940                   statusText = "Success"
950                   statusCell.Value = "Sent (" & http.Status & ")"
960               Else
970                   statusText = "Failed"
980                   statusCell.Value = "Failed: " & http.Status
990               End If
1000              responseText = http.responseText

                ' Write to Log
1010              Dim logRow As Long
1020              logRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1
1030              wsLog.Cells(logRow, 1).Value = Now()
1040              wsLog.Cells(logRow, 2).Value = toName
1050              wsLog.Cells(logRow, 3).Value = toNumber
1060              wsLog.Cells(logRow, 4).Value = messageUid
1070              wsLog.Cells(logRow, 5).Value = statusText
1080              wsLog.Cells(logRow, 6).Value = responseText
1090          Else
1100              statusCell.Value = "Skipped: No number"
1110          End If
1120      End If
1130  Next i

1140  MsgBox "Message sending process complete. Please check the 'Log' sheet for detailed results.", vbInformation, "Process Complete"

Cleanup:
1150  If Not btnSend Is Nothing Then
1160      btnSend.OLEFormat.Object.Enabled = originalButtonState
1170      btnSend.OLEFormat.Object.Caption = "Send Messages"
1180  End If
1190  Set http = Nothing
1200  Exit Sub

ErrorHandler:
1210  MsgBox "An unexpected error occurred: " & vbCrLf & "Error " & Err.Number & " on line " & Erl & ": " & Err.Description, vbCritical, "Runtime Error"
1220  GoTo Cleanup
End Sub

'---------------------------------------------------------------------------------------
' Procedure : BuildContentVariablesJson
' Author    : Jules
' Date      : 2025-09-17
' Purpose   : Builds a JSON string for Twilio's ContentVariables parameter.
'---------------------------------------------------------------------------------------
Private Function BuildContentVariablesJson(vars() As String) As String
    Dim jsonBuilder As String
    Dim i As Long
    jsonBuilder = "{"

    For i = 1 To UBound(vars)
        If vars(i) <> "" Then
            If Len(jsonBuilder) > 1 Then
                jsonBuilder = jsonBuilder & ","
            End If
            jsonBuilder = jsonBuilder & """" & i & """:""" & JsonEncode(vars(i)) & """"
        End If
    Next i

    jsonBuilder = jsonBuilder & "}"
    BuildContentVariablesJson = jsonBuilder
End Function

'---------------------------------------------------------------------------------------
' Procedure : JsonEncode
' Author    : Jules
' Date      : 2025-09-17
' Purpose   : Escapes characters for safe inclusion in a JSON string.
'---------------------------------------------------------------------------------------
Private Function JsonEncode(str As String) As String
    str = Replace(str, "\", "\\")
    str = Replace(str, """", "\""")
    str = Replace(str, vbCrLf, "\n")
    str = Replace(str, vbLf, "\n")
    str = Replace(str, vbCr, "\n")
    str = Replace(str, vbTab, "\t")
    JsonEncode = str
End Function

'---------------------------------------------------------------------------------------
' Procedure : UrlEncode
' Author    : Jules
' Date      : 2025-09-17
' Purpose   : A simplified URL Encoder.
'---------------------------------------------------------------------------------------
Private Function UrlEncode(str As String) As String
    Dim i As Long, tempStr As String, char As String
    For i = 1 To Len(str)
        char = Mid(str, i, 1)
        Select Case char
            Case "a" To "z", "A" To "Z", "0" To "9", "-", "_", ".", "~"
                tempStr = tempStr & char
            Case " "
                tempStr = tempStr & "+"
            Case Else
                tempStr = tempStr & "%" & Right("0" & Hex(Asc(char)), 2)
        End Select
    Next i
    UrlEncode = tempStr
End Function

'---------------------------------------------------------------------------------------
' Procedure : Base64Encode
' Author    : Jules
' Date      : 2025-09-17
' Purpose   : Base64 encodes a string using MSXML2.
'---------------------------------------------------------------------------------------
Private Function Base64Encode(str As String) As String
    Dim arrData() As Byte
    arrData = StrConv(str, vbFromUnicode)
    Dim objXML As Object, objNode As Object
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    Base64Encode = objNode.text
    Set objNode = Nothing
    Set objXML = Nothing
End Function
