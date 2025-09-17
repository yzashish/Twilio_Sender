' VBA Macro for sending Twilio WhatsApp Messages (Hybrid-Variable Version)

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
80    Dim lastRow As Long, i As Long, j As Long
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

270   If Not IsNumeric(wsDashboard.Range("C6").Value) Then
280       MsgBox "Error: 'Messages Per Second (MPS)' must be a valid number.", vbCritical, "Config Error"
290       GoTo Cleanup
300   End If
310   mps = wsDashboard.Range("C6").Value

    ' --- Input Validation ---
320   If accountSid = "" Or authToken = "" Or fromNumber = "" Or messageUid = "" Or mps <= 0 Then
330       MsgBox "Error: Please ensure all settings and inputs on the Dashboard and Settings sheets are filled in correctly.", vbCritical, "Config Error"
340       GoTo Cleanup
350   End If

    ' --- Find Template and Get Base Variables ---
360   templateRow = Application.Match(messageUid, wsTemplates.Columns(1), 0)
370   If IsError(templateRow) Then
380       MsgBox "Error: The Message UID '" & messageUid & "' was not found in the 'Templates' sheet.", vbCritical, "Template Error"
390       GoTo Cleanup
400   End If

410   contentSid = wsTemplates.Cells(templateRow, 2).Value
420   If contentSid = "" Then
430       MsgBox "Error: The ContentSid for Message UID '" & messageUid & "' is empty.", vbCritical, "Template Error"
440       GoTo Cleanup
450   End If

460   For j = 1 To 15
470       baseVariables(j) = wsTemplates.Cells(templateRow, 2 + j).Value
480   Next j

    ' --- Main Loop ---
490   Set http = CreateObject("MSXML2.XMLHTTP")
500   lastRow = wsContacts.Cells(wsContacts.Rows.Count, "A").End(xlUp).Row
510   messagesSentThisSecond = 0

520   For i = 2 To lastRow
530       If UCase(wsContacts.Cells(i, 3).Value) = "YES" Or UCase(wsContacts.Cells(i, 3).Value) = "TRUE" Then
            ' Rate Limiting
540           If messagesSentThisSecond >= mps Then
550               Application.Wait (Now + TimeValue("0:00:01"))
560               messagesSentThisSecond = 0
570           End If

580           Dim toName As String, toNumber As String, statusCell As Range
590           toName = wsContacts.Cells(i, 1).Value
600           toNumber = wsContacts.Cells(i, 2).Value
610           Set statusCell = wsContacts.Cells(i, 4)
620           statusCell.Value = ""

630           If toNumber <> "" Then
640               Dim resolvedVariables(1 To 15) As String, finalJson As String
650               Dim url As String, body As String, statusText As String, responseText As String

                ' Two-Tiered Variable Replacement
660               For j = 1 To 15
670                   Dim tempVar As String
680                   tempVar = baseVariables(j)
690                   tempVar = Replace(tempVar, "{{name}}", toName)
700                   For k = 1 To 15
710                       tempVar = Replace(tempVar, "{{contact_var_" & k & "}}", wsContacts.Cells(i, 4 + k).Value)
720                   Next k
730                   resolvedVariables(j) = tempVar
740               Next j

                ' Build JSON and API Body
750               finalJson = BuildContentVariablesJson(resolvedVariables)
760               url = "https://api.twilio.com/2010-04-01/Accounts/" & accountSid & "/Messages.json"
770               body = "To=" & UrlEncode("whatsapp:" & toNumber) & _
                      "&From=" & UrlEncode("whatsapp:" & fromNumber) & _
                      "&ContentSid=" & contentSid & _
                      "&ContentVariables=" & UrlEncode(finalJson)

                ' Send Request
780               http.Open "POST", url, False
790               http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
800               http.setRequestHeader "Authorization", "Basic " & Base64Encode(accountSid & ":" & authToken)
810               http.send body
820               messagesSentThisSecond = messagesSentThisSecond + 1

                ' Process Response
830               If http.Status >= 200 And http.Status < 300 Then
840                   statusText = "Success"
850                   statusCell.Value = "Sent (" & http.Status & ")"
860               Else
870                   statusText = "Failed"
880                   statusCell.Value = "Failed: " & http.Status
890               End If
900               responseText = http.responseText

                ' Write to Log
910               Dim logRow As Long
920               logRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1
930               wsLog.Cells(logRow, 1).Value = Now()
940               wsLog.Cells(logRow, 2).Value = toName
950               wsLog.Cells(logRow, 3).Value = toNumber
960               wsLog.Cells(logRow, 4).Value = messageUid
970               wsLog.Cells(logRow, 5).Value = statusText
980               wsLog.Cells(logRow, 6).Value = responseText
990           Else
1000              statusCell.Value = "Skipped: No number"
1010          End If
1020      End If
1030  Next i

1040  MsgBox "Message sending process complete. Please check the 'Log' sheet for detailed results.", vbInformation, "Process Complete"

Cleanup:
1050  If Not btnSend Is Nothing Then
1060      btnSend.OLEFormat.Object.Enabled = originalButtonState
1070      btnSend.OLEFormat.Object.Caption = "Send Messages"
1080  End If
1090  Set http = Nothing
1100  Exit Sub

ErrorHandler:
1110  MsgBox "An unexpected error occurred: " & vbCrLf & "Error " & Err.Number & " on line " & Erl & ": " & Err.Description, vbCritical, "Runtime Error"
1120  GoTo Cleanup
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
