VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HTTPServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents Winsock As Winsock
Attribute Winsock.VB_VarHelpID = -1
Private parentForm As Form
Private ServerRunning As Boolean

Private Declare Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal codepage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cchMultiByte As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long) As Long
Private Const CP_UTF8 As Long = 65001

Public Function ByteArrayToString(byteData() As Byte, Optional ByVal encoding As String = "UTF-8") As String
    Dim strResult As String
    Dim Length As Long
    Dim byteLen As Long
    ' Calculate byte length
    byteLen = UBound(byteData) - LBound(byteData) + 1
    If byteLen <= 0 Then
        ByteArrayToString = ""
        Exit Function
    End If
    ' Get required buffer size for UTF-16
    Length = MultiByteToWideChar(CP_UTF8, 0, VarPtr(byteData(0)), byteLen, 0, 0)
    strResult = String$(Length, vbNullChar)
    ' Convert to UTF-16 string
    MultiByteToWideChar CP_UTF8, 0, VarPtr(byteData(0)), byteLen, StrPtr(strResult), Length
    ByteArrayToString = strResult
End Function
'-------------------

Private Sub Class_Initialize()
    Set Winsock = Form1.Winsock1
    ServerRunning = False
End Sub

Private Sub Class_Terminate()
    Set Winsock = Nothing
    Set parentForm = Nothing
End Sub

Public Sub Initialize(frm As Form)
    Set parentForm = frm
End Sub

' Start HTTP server
Public Sub StartServer()
    With Winsock
        .Close
        .LocalPort = 62007
        .Bind 62007, "0.0.0.0"
        .Listen
    End With
    Exit Sub
ErrorHandler:
    MsgBox "Error starting server: " & Err.Description, vbCritical
End Sub

Private Sub Winsock_ConnectionRequest(ByVal requestID As Long)
    Dim newSock As Winsock
    Set newSock = Form1.Winsock1
    If newSock.State <> sckClosed Then newSock.Close
    newSock.Accept requestID
End Sub

Private Sub Winsock_Close()
    StartServer
End Sub

Private Function UTF8BytesToString(bytes() As Byte) As String
    With CreateObject("ADODB.Stream")
        .Type = 1 ' adTypeBinary
        .Open
        .Write bytes
        .Position = 0
        .Type = 2 ' adTypeText
        .Charset = "UTF-8"
        UTF8BytesToString = .ReadText
        .Close
    End With
End Function

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
'On Error GoTo ErrorHandler
    Dim byteData() As Byte
    Dim strData As String
    Dim jsonObj As Object
    Dim encodeText As String
    Dim systemTemplate As String
    Dim modelName As String
    Dim temperature As Double
    Dim apiUrl As String

    Winsock.GetData byteData, vbArray + vbByte
    strData = UTF8BytesToString(byteData)
    
    Dim bodyStart As Long
    bodyStart = InStr(strData, vbCrLf & vbCrLf)
    If bodyStart > 0 Then
        strData = Mid$(strData, bodyStart + 4)
    End If

    ' Parse JSON
    On Error GoTo PassErrorHandler
    Set jsonObj = JsonConverter.ParseJson(strData)
    If jsonObj Is Nothing Then
        SendResponse Winsock, "Invalid JSON", 400
        Exit Sub
    End If

    ' Check encode_text field
    If Not jsonObj.Exists("encode_text") Then
        MsgBox "Request does not contain encode_text field", vbExclamation
        SendResponse Winsock, "Missing encode_text field", 400
        Exit Sub
    End If
    encodeText = jsonObj("encode_text")
    
    RichTextboxUniText(parentForm.Text3) = encodeText

    ' Check system_template field
    If Not jsonObj.Exists("system_template") Then
        MsgBox "Request does not contain system_template field", vbExclamation
        SendResponse Winsock, "Missing system_template field", 400
        Exit Sub
    End If
    systemTemplate = jsonObj("system_template")

    ' Check model field
    If Not jsonObj.Exists("model") Then
        MsgBox "Request does not contain model field", vbExclamation
        SendResponse Winsock, "Missing model field", 400
        Exit Sub
    End If
    modelName = jsonObj("model")

    ' Check temperature field
    If Not jsonObj.Exists("temperature") Then
        MsgBox "Request does not contain temperature field", vbExclamation
        SendResponse Winsock, "Missing temperature field", 400
        Exit Sub
    End If
    temperature = jsonObj("temperature")

    ' Check api_url field
    If Not jsonObj.Exists("api_url") Then
        MsgBox "Request does not contain api_url field", vbExclamation
        SendResponse Winsock, "Missing api_url field", 400
        Exit Sub
    End If
    apiUrl = jsonObj("api_url")

    ' Send translation request with parameters
    SendTranslationRequest encodeText, systemTemplate, modelName, temperature, apiUrl

    ' Send response to client
    SendResponse Winsock, "Translation request sent", 200
    Exit Sub
ErrorHandler:
    SendResponse Winsock, "Internal Server Error", 500
    MsgBox "Error processing request: " & Err.Description, vbCritical
PassErrorHandler:
End Sub

Private Sub SendTranslationRequest(textToTranslate As String, systemTemplate As String, modelName As String, temperature As Double, apiUrl As String)
    On Error GoTo ErrorHandler
    ' Show window
    parentForm.TrayHelper_TrayDblClick
    parentForm.Text1.Text = "" ' Clear text box

    Dim jsonData As String
    Dim dict As Object
    Dim messages As Object
    Dim systemMessage As Object
    Dim userMessage As Object

    ' Create JSON object
    Set dict = CreateObject("Scripting.Dictionary")
    Set messages = CreateObject("Scripting.Dictionary")

    ' Set system message
    Set systemMessage = CreateObject("Scripting.Dictionary")
    systemMessage.Add "role", "system"
    systemMessage.Add "content", systemTemplate

    ' Set user message
    Set userMessage = CreateObject("Scripting.Dictionary")
    userMessage.Add "role", "user"
    userMessage.Add "content", "<translate_input>" & textToTranslate & "</translate_input>"

    ' Add parameters to JSON
    dict.Add "model", modelName
    dict.Add "stream", True
    dict.Add "keep_alive", "0m"
    dict.Add "temperature", temperature
    dict.Add "messages", Array(systemMessage, userMessage)

    ' Convert to JSON string
    jsonData = JsonConverter.ConvertToJson(dict)

    ' Initialize stream processor
    WinINetHelper.InitializeStreamProcessor parentForm

    ' Send request to API
    WinINetHelper.SendStreamRequest apiUrl, jsonData
    Exit Sub
ErrorHandler:
    MsgBox "Error sending translation request: " & Err.Description, vbCritical
End Sub

' Send HTTP response
Private Sub SendResponse(ws As Winsock, message As String, statusCode As Integer)
    On Error Resume Next
    Dim response As String
    Dim statusText As String

    Select Case statusCode
        Case 200
            statusText = "OK"
        Case 400
            statusText = "Bad Request"
        Case 500
            statusText = "Internal Server Error"
        Case Else
            statusText = "Unknown Status"
    End Select

    response = "HTTP/1.1 " & statusCode & " " & statusText & vbCrLf
    response = response & "Content-Type: text/plain" & vbCrLf
    response = response & "Content-Length: " & Len(message) & vbCrLf
    response = response & "Connection: close" & vbCrLf
    response = response & vbCrLf
    response = response & message

    ws.SendData response
    On Error GoTo 0
End Sub




