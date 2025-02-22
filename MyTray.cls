VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' MyTray.cls
Option Explicit

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" _
    (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 128
End Type

Private Const NIM_ADD As Long = &H0
Private Const NIM_MODIFY As Long = &H1
Private Const NIM_DELETE As Long = &H2
Private Const NIF_MESSAGE As Long = &H1
Private Const NIF_ICON As Long = &H2
Private Const NIF_TIP As Long = &H4
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_LBUTTONDBLCLK As Long = &H203

Private WithEvents m_Form As Form
Attribute m_Form.VB_VarHelpID = -1
Private m_nid As NOTIFYICONDATA
Public Event TrayDblClick()

Public Sub Initialize(ByRef frm As Form, Optional icon As StdPicture, Optional tip As String = "")
    Set m_Form = frm
    With m_nid
        .cbSize = Len(m_nid)
        .hwnd = m_Form.hwnd
        .uID = 1
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        
        If icon Is Nothing Then
            .hIcon = m_Form.icon.Handle
        Else
            .hIcon = icon.Handle
        End If
        
        .szTip = tip & vbNullChar
    End With
End Sub


Public Sub AddToTray()
    Shell_NotifyIcon NIM_ADD, m_nid
End Sub

Public Sub RemoveFromTray()
    Shell_NotifyIcon NIM_DELETE, m_nid
End Sub

Private Sub m_Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Static msg As Long
    msg = CLng(x / Screen.TwipsPerPixelX)
    
    If msg = WM_LBUTTONDBLCLK Then
        RaiseEvent TrayDblClick
    End If
End Sub


