Attribute VB_Name = "StreamProcessor"
Option Explicit

' Variable declarations grouped by usage
Private combinedContent As String           ' Stores accumulated response content
Private xmlhttp As Object                   ' Handles HTTP request operations
Private parentForm As Object                ' Reference to the parent form for UI updates
Private lastProcessedLength As Long         ' Tracks length of processed response
Private isProcessing As Boolean             ' Controls processing state to prevent re-entry

' Initializes the stream processor with required objects
Public Sub InitializeStreamProcessor(xmlhttpObj As Object, parentFormObj As Object)
    Set xmlhttp = xmlhttpObj
    Set parentForm = parentFormObj
    combinedContent = ""
    lastProcessedLength = 0
    isProcessing = False
End Sub

' Processes the streaming response from xmlhttp
Public Sub ProcessStream()
    If isProcessing Then Exit Sub           ' Prevent reentrant calls
    isProcessing = True
    
    If xmlhttp.readyState >= 3 Then         ' Check if data is being received (3) or complete (4)
        Dim currentResponse As String
        currentResponse = ""
        
        On Error Resume Next                ' Handle potential errors in response retrieval
        currentResponse = xmlhttp.responseText
        On Error GoTo 0
        
        If Len(currentResponse) > lastProcessedLength Then
            Dim newData As String
            newData = Mid$(currentResponse, lastProcessedLength + 1)   ' Extract new data portion
            lastProcessedLength = Len(currentResponse)                 ' Update processed length
            ProcessNewData newData                                     ' Process the new data
        End If
    End If
    
    If xmlhttp.readyState = 4 Then          ' Handle completion state
        parentForm.Text2.Text = combinedContent    ' Update UI with final content
        parentForm.Timer1.Enabled = False          ' Disable timer on completion
    End If
    
    isProcessing = False                    ' Reset processing flag
    On Error GoTo 0                         ' Reset error handling
End Sub

' Processes newly received data chunks
Private Sub ProcessNewData(newData As String)
    Dim lines() As String
    lines = Split(newData, vbLf)            ' Split data into lines
    
    Dim i As Integer
    For i = 0 To UBound(lines)
        Dim chunk As String
        chunk = Trim(lines(i))              ' Remove leading/trailing whitespace
        
        If Len(chunk) > 0 Then              ' Process non-empty chunks
            If Left(chunk, 6) = "data: " Then    ' Remove "data: " prefix if present
                chunk = Mid(chunk, 7)
            End If
            
            If InStr(chunk, "[DONE]") > 0 Then   ' Exit if completion marker found
                Exit Sub
            End If
            
            Dim jsonObj As Object
            On Error Resume Next            ' Handle potential JSON parsing errors
            Set jsonObj = JsonConverter.ParseJson(chunk)
            On Error GoTo 0
            
            If Not jsonObj Is Nothing Then  ' Process valid JSON objects
                If Not IsEmpty(jsonObj("choices")) Then
                    If Not IsEmpty(jsonObj("choices")(1)("delta")) Then
                        If Not IsEmpty(jsonObj("choices")(1)("delta")("content")) Then
                            combinedContent = combinedContent & jsonObj("choices")(1)("delta")("content")    ' Append content
                        End If
                    End If
                End If
            End If
            
            parentForm.Text1.Text = parentForm.Text1.Text & chunk & vbCrLf    ' Update UI with raw chunk
        End If
    Next i
End Sub

' Returns the accumulated content
Public Function GetCombinedContent() As String
    GetCombinedContent = combinedContent
End Function

' Cleans up object references
Public Sub Cleanup()
    Set xmlhttp = Nothing
    Set parentForm = Nothing
End Sub
