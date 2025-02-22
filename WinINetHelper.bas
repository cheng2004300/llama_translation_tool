Attribute VB_Name = "WinINetHelper"
Option Explicit

' Windows API constants for window management
Private Const WM_SETREDRAW = &HB
Private Const RDW_INVALIDATE = &H1
Private Const RDW_UPDATENOW = &H100
Private Const RDW_ALLCHILDREN = &H80

' WinINet API constants for HTTP operations
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_SERVICE_HTTP = 3
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
Private Const HTTP_ADDREQ_FLAG_ADD = &H20000000
Private Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000

' Windows API declarations
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Any) As Long

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

' WinINet API declarations for internet connectivity
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
    (ByVal lpszAgent As String, ByVal dwAccessType As Long, _
    ByVal lpszProxyName As String, ByVal lpszProxyBypass As String, _
    ByVal dwFlags As Long) As Long

Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
    (ByVal hInternet As Long, ByVal lpszServerName As String, _
    ByVal nServerPort As Long, ByVal lpszUsername As String, _
    ByVal lpszPassword As String, ByVal dwService As Long, _
    ByVal dwFlags As Long, ByVal dwContext As Long) As Long

Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" _
    (ByVal hConnect As Long, ByVal lpszVerb As String, _
    ByVal lpszObjectName As String, ByVal lpszVersion As String, _
    ByVal lpszReferer As String, ByVal lpszAcceptTypes As Long, _
    ByVal dwFlags As Long, ByVal dwContext As Long) As Long

Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" _
    (ByVal hRequest As Long, ByVal lpszHeaders As String, _
    ByVal dwHeadersLength As Long, ByVal lpOptional As String, _
    ByVal dwOptionalLength As Long) As Long

Private Declare Function InternetReadFile Lib "wininet.dll" _
    (ByVal hFile As Long, ByVal lpBuffer As Long, _
    ByVal dwNumberOfBytesToRead As Long, lpdwNumberOfBytesRead As Long) As Long

Private Declare Function InternetCloseHandle Lib "wininet.dll" _
    (ByVal hInternet As Long) As Long

Private Declare Function HttpAddRequestHeaders Lib "wininet.dll" Alias "HttpAddRequestHeadersA" _
    (ByVal hRequest As Long, ByVal lpszHeaders As String, _
    ByVal dwHeadersLength As Long, ByVal dwModifiers As Long) As Long

Private Declare Function InternetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" _
    (ByVal hInternet As Long, ByVal dwOption As Long, lpBuffer As Any, ByVal dwBufferLength As Long) As Boolean

Const INTERNET_OPTION_CONNECT_TIMEOUT = 2

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal Length As Long)

' Private module-level variables
Private hInternet As Long        ' Handle for internet session
Private hConnect As Long         ' Handle for server connection
Private hRequest As Long         ' Handle for HTTP request
Private parentForm As Object     ' Reference to the parent form
Private combinedContent As String ' Accumulated response content
Private isProcessing As Boolean  ' Flag to prevent re-entrant processing

' Initializes the streaming processor with a parent form
Public Sub InitializeStreamProcessor(parentFormObj As Object)
    Set parentForm = parentFormObj
    combinedContent = ""
    isProcessing = False
    
    ' Open an internet session using WinINet
    hInternet = InternetOpen("VB6 Stream Client", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    If hInternet = 0 Then
        MsgBox "Failed to initialize Internet session"
        Exit Sub
    End If
End Sub

' Sends an HTTP POST request with streaming capability
Public Sub SendStreamRequest(url As String, postData As String)
    On Error Resume Next
    
    Dim serverName As String
    Dim port As Long
    Dim path As String
    Dim lTimeout As Long
    
    ' Parse URL into components
    ParseUrl url, serverName, port, path
    

    lTimeout = 100 'set connect timeout

    InternetSetOption hInternet, INTERNET_OPTION_CONNECT_TIMEOUT, lTimeout, 4

    ' Connect to the server
    hConnect = InternetConnect(hInternet, serverName, port, vbNullString, vbNullString, INTERNET_SERVICE_HTTP, 0, 0)
    If hConnect = 0 Then
        MsgBox "Failed to connect to server"
        Cleanup
        Exit Sub
    End If
    

    ' Create an HTTP POST request
    hRequest = HttpOpenRequest(hConnect, "POST", path, "HTTP/1.1", vbNullString, 0, _
        INTERNET_FLAG_RELOAD Or INTERNET_FLAG_NO_CACHE_WRITE, 0)
    If hRequest = 0 Then
        MsgBox "Failed to create HTTP request"
        Cleanup
        Exit Sub
    End If
    
    ' Set request headers for JSON communication
    Dim headers As String
    headers = "Content-Type: application/json" & vbCrLf & _
              "Accept: application/json" & vbCrLf & _
              "Accept-Charset: utf-8" & vbCrLf
    HttpAddRequestHeaders hRequest, headers, Len(headers), HTTP_ADDREQ_FLAG_ADD Or HTTP_ADDREQ_FLAG_REPLACE
    
    ' Send the request with POST data
    Dim result As Long
    result = HttpSendRequest(hRequest, vbNullString, 0, postData, Len(postData))
    If result = 0 Then
        MsgBox "Failed to send HTTP request"
        Cleanup
        Exit Sub
    End If
    
    ' Start a timer to process the stream periodically
    parentForm.Timer1.Interval = 100
    parentForm.Timer1.Enabled = True
    
    On Error GoTo 0
End Sub

' Processes incoming stream data
Public Sub ProcessStream()
    On Error Resume Next
    
    If isProcessing Then Exit Sub
    isProcessing = True
    
    Dim byteBuffer() As Byte
    Dim bytesRead As Long
    Dim success As Long
    
    ' Allocate a 4KB buffer for reading data
    ReDim byteBuffer(4095)
    
    ' Read data from the HTTP stream
    success = InternetReadFile(hRequest, VarPtr(byteBuffer(0)), UBound(byteBuffer) + 1, bytesRead)
    
    If success <> 0 And bytesRead > 0 Then
        ' Resize buffer to actual data size
        ReDim Preserve byteBuffer(bytesRead - 1)
        
        ' Convert bytes to UTF-8 string
        Dim newData As String
        newData = BytesToString(byteBuffer)
        
        ' Process the received data
        ProcessNewData newData
    End If
    
    ' Check if stream has ended
    If success <> 0 And bytesRead = 0 Then
        ' Update secondary textbox with final content and clean up
        RichTextboxUniText(parentForm.Text2) = combinedContent
        parentForm.Timer1.Enabled = False
        Cleanup
    End If
    
    isProcessing = False
    On Error GoTo 0
End Sub

' Processes new data chunks from the stream
Private Sub ProcessNewData(newData As String)
    Dim lines() As String
    lines = Split(newData, vbLf)
    
    Dim i As Integer
    For i = 0 To UBound(lines)
        Dim chunk As String
        chunk = Trim(lines(i))
        
        If Len(chunk) > 0 Then
            ' Strip "data: " prefix if present
            If Left(chunk, 6) = "data: " Then
                chunk = Mid(chunk, 7)
            End If
            
            ' Exit if end marker is found
            If InStr(chunk, "[DONE]") > 0 Then
                Exit Sub
            End If
            
            ' Parse JSON content
            Dim jsonObj As Object
            On Error Resume Next
            Set jsonObj = JsonConverter.ParseJson(chunk)
            
            If Not jsonObj Is Nothing Then
                If Not IsEmpty(jsonObj("choices")) Then
                    If Not IsEmpty(jsonObj("choices")(1)("delta")) Then
                        If Not IsEmpty(jsonObj("choices")(1)("delta")("content")) Then
                            ' Append new content to combined result
                            combinedContent = combinedContent & jsonObj("choices")(1)("delta")("content")
                            
                            ' Update primary textbox with minimal flicker
                            SendMessage parentForm.Text1.hWnd, WM_SETREDRAW, 0, 0
                            RichTextboxUniText(parentForm.Text1) = combinedContent
                            parentForm.Text1.ScrollToCaret
                            parentForm.Text1.SelStart = 9999
                            SendMessage parentForm.Text1.hWnd, WM_SETREDRAW, 1, 0
                            RedrawWindow parentForm.Text1.hWnd, ByVal 0&, 0, RDW_INVALIDATE Or RDW_UPDATENOW Or RDW_ALLCHILDREN
                        End If
                    End If
                End If
            End If
        End If
    Next i
End Sub

' Returns the accumulated content
Public Function GetCombinedContent() As String
    GetCombinedContent = combinedContent
End Function

' Cleans up all open handles and references
Public Sub Cleanup()
    If hRequest <> 0 Then
        InternetCloseHandle hRequest
        hRequest = 0
    End If
    If hConnect <> 0 Then
        InternetCloseHandle hConnect
        hConnect = 0
    End If
    If hInternet <> 0 Then
        InternetCloseHandle hInternet
        hInternet = 0
    End If
    Set parentForm = Nothing
End Sub

' Parses a URL into server name, port, and path
Private Sub ParseUrl(url As String, ByRef serverName As String, ByRef port As Long, ByRef path As String)
    ' Remove "http://" prefix from the URL
    url = Replace(url, "http://", "")

    ' Split the URL by "/" into at most two parts: server part and path part
    Dim parts() As String
    parts = Split(url, "/", 2)

    ' Extract the server and port portion
    Dim serverAndPort As String
    serverAndPort = parts(0)

    ' Set the path: if there a path part, prepend "/"; otherwise, use "/"
    If UBound(parts) > 0 Then
        path = "/" & parts(1)
    Else
        path = "/"
    End If

    ' Split the server part by ":" into at most two parts: server name and port
    Dim serverParts() As String
    serverParts = Split(serverAndPort, ":", 2)

    ' Set the server name
    serverName = serverParts(0)

    ' Set the port: if a port part exists, try converting it to a number; otherwise, use default port 80
    If UBound(serverParts) > 0 Then
        Dim portStr As String
        portStr = serverParts(1)
        If IsNumeric(portStr) Then
            port = CLng(portStr)
        Else
            port = 80 ' Use default port if the port is not numeric
        End If
    Else
        port = 80
    End If
End Sub

' Extracts a subset of an array starting from a given index
Private Function SliceArray(arr() As String, startIndex As Long) As String()
    Dim result() As String
    ReDim result(UBound(arr) - startIndex)
    Dim i As Long
    For i = startIndex To UBound(arr)
        result(i - startIndex) = arr(i)
    Next i
    SliceArray = result
End Function

' Converts a byte array to a UTF-8 string using ADODB.Stream
Private Function BytesToString(byteArray() As Byte) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    
    On Error Resume Next
    
    stream.Type = 1 ' Binary
    stream.Open
    stream.Write byteArray
    stream.Position = 0
    stream.Type = 2 ' Text
    stream.Charset = "utf-8"
    BytesToString = stream.ReadText
    stream.Close
    
    Set stream = Nothing
    On Error GoTo 0
End Function

