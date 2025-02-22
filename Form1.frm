VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "LlamaTranslation"
   ClientHeight    =   6540
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   9060
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   3240
      Top             =   6000
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2640
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox Text2 
      Height          =   3255
      Left            =   9240
      TabIndex        =   4
      Top             =   1080
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5741
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":324A
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   2655
      Left            =   240
      TabIndex        =   3
      Top             =   3360
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   4683
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      RightMargin     =   1
      TextRTF         =   $"Form1.frx":32D9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "PMingLiU"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox Text3 
      Height          =   2655
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   4683
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      RightMargin     =   1
      TextRTF         =   $"Form1.frx":336C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "PMingLiU"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   11400
      TabIndex        =   1
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2040
      Top             =   6000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   9720
      TabIndex        =   0
      Top             =   4560
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Server As HTTPServer
Private WithEvents TrayHelper As MyTray
Attribute TrayHelper.VB_VarHelpID = -1

' API declarations for window manipulation
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

' Constants for window styles and positioning
Private Const GWL_STYLE = (-16)
Private Const WS_THICKFRAME = &H40000
Private Const HWND_TOPMOST = -1
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_SHOWWINDOW = &H40

' Button click handler to set up Unicode text in a RichTextBox
Private Sub Command2_Click()
    SetupRichTextboxForUnicode Text3
    RichTextboxUniText(Text3) = ChrW(&H6C49) & ChrW(&H5B57) ' Displays "??" (Chinese characters)
End Sub

' Timer event to process network stream
Private Sub Timer1_Timer()
    On Error Resume Next
    WinINetHelper.ProcessStream
    On Error GoTo 0
End Sub

' Form load event to initialize window and server
Private Sub Form_Load()
    Dim lStyle As Long
    ' Remove thick frame (resize border) from the window
    lStyle = GetWindowLong(Me.hWnd, GWL_STYLE)
    lStyle = lStyle And Not WS_THICKFRAME
    SetWindowLong Me.hWnd, GWL_STYLE, lStyle
    
    ' Set window to topmost without changing size or position
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
    
    ' Bring window to foreground
    SetForegroundWindow Me.hWnd
    
    ' Initialize and start the HTTP server
    Set Server = New HTTPServer
    Server.Initialize Me
    Server.StartServer
    
    ' Initialize system tray helper
    Set TrayHelper = New MyTray
    TrayHelper.Initialize Me, , "LlamaTranslation"
End Sub

' Handle form resize to minimize to system tray
Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        TrayHelper.AddToTray
        Me.Hide
    End If
End Sub

Private Sub Timer2_Timer()
    Server.StartServer
End Sub

' Restore form from system tray on double-click
Public Sub TrayHelper_TrayDblClick()
    Me.Visible = True
    Me.WindowState = vbNormal
    Me.Show
    
    ' Small delay (100ms) to ensure smooth UI update
    Dim startTime As Double
    startTime = Timer
    Do While Timer < startTime + 0.1
        DoEvents
    Loop
    
    TrayHelper.RemoveFromTray
End Sub

' Cleanup on form unload
Private Sub Form_Unload(Cancel As Integer)
    TrayHelper.RemoveFromTray
    Set TrayHelper = Nothing
End Sub
