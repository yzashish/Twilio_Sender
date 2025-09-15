' VBA Macro for sending Twilio WhatsApp Messages

' Main Subroutine to be assigned to the button
Public Sub SendTwilioMessages()
    On Error GoTo ErrorHandler

    Dim wsDashboard As Worksheet
    Dim wsContacts As Worksheet
    Dim wsTemplates As Worksheet
    Dim wsSettings As Worksheet

    Dim btnSend As Object
    Dim originalButtonState As Boolean

    ' --- Configuration ---
    Set wsDashboard = ThisWorkbook.Sheets("Dashboard")
    Set wsContacts = ThisWorkbook.Sheets("Contacts")
    Set wsTemplates = ThisWorkbook.Sheets("Templates")
    Set wsSettings = ThisWorkbook.Sheets("Settings")

    ' Get the send button from the Dashboard to disable it
    On Error Resume Next
    Set btnSend = wsDashboard.Shapes("btnSendMessages")
    If Not btnSend Is Nothing Then
        originalButtonState = btnSend.OLEFormat.Object.Enabled
        btnSend.OLEFormat.Object.Enabled = False
        btnSend.OLEFormat.Object.Caption = "Sending..."
    End If
    On Error GoTo ErrorHandler ' Re-enable default error handling

    ' --- Read Settings ---
    Dim accountSid As String
    Dim authToken As String
    Dim fromNumber As String

    accountSid = wsSettings.Range("B1").Value
    authToken = wsSettings.Range("B2").Value
    fromNumber = wsSettings.Range("B3").Value

    If accountSid = "" Or authToken = "" Or fromNumber = "" Then
        MsgBox "Error: Please provide Account SID, Auth Token, and Twilio 'From' number in the 'Settings' sheet.", vbCritical, "Configuration Error"
        GoTo Cleanup
    End If

    ' --- Read Selected Template ---
    Dim messageUid As String
    Dim messageTemplate As String
    Dim templateRow As Variant

    ' The user specifies the Message UID in a dedicated cell on the 'Dashboard'.
    messageUid = wsDashboard.Range("C5").Value

    If messageUid = "" Then
        MsgBox "Error: Please specify a Message UID in cell C5 on the 'Dashboard' sheet.", vbCritical, "Template Error"
        GoTo Cleanup
    End If

    ' Find the row number of the matching UID in the Templates sheet (Column A)
    templateRow = Application.Match(messageUid, wsTemplates.Columns(1), 0)

    ' Check if a match was found
    If IsError(templateRow) Then
        MsgBox "Error: The Message UID '" & messageUid & "' was not found in the 'Templates' sheet.", vbCritical, "Template Error"
        GoTo Cleanup
    Else
        ' Get the template content from Column B of the found row
        messageTemplate = wsTemplates.Cells(templateRow, 2).Value
    End If

    If messageTemplate = "" Then
        MsgBox "Error: The found template for UID '" & messageUid & "' is empty.", vbCritical, "Template Error"
        GoTo Cleanup
    End If

    ' --- Create HTTP Object ---
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")

    ' --- Loop Through Contacts ---
    Dim lastRow As Long
    Dim i As Long
    Dim toName As String
    Dim toNumber As String
    Dim sendFlag As String
    Dim statusCell As Range

    lastRow = wsContacts.Cells(wsContacts.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow
        toName = wsContacts.Cells(i, 1).Value
        toNumber = wsContacts.Cells(i, 2).Value
        sendFlag = UCase(wsContacts.Cells(i, 3).Value)
        Set statusCell = wsContacts.Cells(i, 4)

        statusCell.Value = "" ' Clear previous status

        If sendFlag = "YES" Or sendFlag = "TRUE" Then
            If toNumber <> "" Then
                ' --- Prepare API Request ---
                Dim url As String
                Dim body As String
                Dim finalMessage As String

                ' Simple placeholder replacement
                finalMessage = Replace(messageTemplate, "{{1}}", wsContacts.Cells(i, 5).Value)
                finalMessage = Replace(finalMessage, "{{2}}", wsContacts.Cells(i, 6).Value)
                finalMessage = Replace(finalMessage, "{{3}}", wsContacts.Cells(i, 7).Value)
                finalMessage = Replace(finalMessage, "{{4}}", wsContacts.Cells(i, 8).Value)

                url = "https://api.twilio.com/2010-04-01/Accounts/" & accountSid & "/Messages.json"

                body = "To=" & UrlEncode("whatsapp:" & toNumber) & _
                       "&From=" & UrlEncode("whatsapp:" & fromNumber) & _
                       "&Body=" & UrlEncode(finalMessage)

                http.Open "POST", url, False
                http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
                http.setRequestHeader "Authorization", "Basic " & Base64Encode(accountSid & ":" & authToken)

                ' --- Send Request ---
                http.send body

                ' --- Log Status ---
                If http.Status >= 200 And http.Status < 300 Then
                    statusCell.Value = "Sent (" & http.Status & ")"
                    statusCell.Font.Color = RGB(0, 128, 0) ' Green
                Else
                    statusCell.Value = "Failed: " & http.Status & " - " & http.responseText
                    statusCell.Font.Color = RGB(255, 0, 0) ' Red
                End If
            Else
                statusCell.Value = "Skipped: No number"
            End If
        End If
    Next i

    MsgBox "Message sending process complete. Please check the 'Status' column in the 'Contacts' sheet for details.", vbInformation, "Process Complete"

Cleanup:
    ' Re-enable the button
    If Not btnSend Is Nothing Then
        btnSend.OLEFormat.Object.Enabled = originalButtonState
        btnSend.OLEFormat.Object.Caption = "Send Messages"
    End If
    Set http = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An unexpected error occurred: " & vbCrLf & "Error " & Err.Number & ": " & Err.Description, vbCritical, "Runtime Error"
    GoTo Cleanup
End Sub


' --- Helper Functions ---

Private Function UrlEncode(str As String) As String
    ' A simplified URL Encoder
    Dim i As Long
    Dim tempStr As String
    Dim char As String

    For i = 1 To Len(str)
        char = Mid(str, i, 1)
        Select Case char
            Case "a" To "z", "A" To "Z", "0" To "9", "-", "_", "."
                tempStr = tempStr & char
            Case " "
                tempStr = tempStr & "+"
            Case Else
                tempStr = tempStr & "%" & Hex(Asc(char))
        End Select
    Next i
    UrlEncode = tempStr
End Function

Private Function Base64Encode(str As String) As String
    ' Base64 encodes a string using MSXML2
    Dim arrData() As Byte
    arrData = StrConv(str, vbFromUnicode)

    Dim objXML As Object
    Dim objNode As Object

    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")

    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData

    Base64Encode = objNode.text

    Set objNode = Nothing
    Set objXML = Nothing
End Function
