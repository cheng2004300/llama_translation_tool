Attribute VB_Name = "utf8lng"
Option Explicit
'
Private Type GETTEXTEX
    cb As Long
    flags As Long
    codepage As Long
    lpDefaultChar As Long
    lpUsedDefChar As Long
End Type
'
Private Type GETTEXTLENGTHEX
    flags As Long
    codepage As Long
End Type
'
Private Type SETTEXTEX
    flags As Long
    codepage As Long
End Type
'
Private Declare Function DefWindowProcW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub PutMem4 Lib "msvbvm60" (Destination As Any, Value As Any)
Private Declare Function SysAllocStringLen Lib "oleaut32" (ByVal OleStr As Long, ByVal bLen As Long) As Long
Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32.dll" () As Long
Private Declare Function CloseClipboard Lib "user32.dll" () As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SendMessageWLng Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'
' The following is from MSDN help:
'
' UNICODE: International Standards Organization (ISO) character standard.
' Unicode uses a 16-bit (2-byte) coding scheme that allows for 65,536 distinct character spaces.
' Unicode includes representations for punctuation marks, mathematical symbols, and dingbats,
' with substantial room for future expansion.
'
' vbUnicode constant:     Converts the string toUnicode using the default code page of the system.
' vbFromUnicode constant: Converts the string from Unicode to the default code page of the system.
'
' LCID: The LocaleID, if different than the system LocaleID. (The system LocaleID is the default.)
'

Public Property Let UniCaption(ctrl As Object, sUniCaption As String)
    Const WM_SETTEXT As Long = &HC
    ' USAGE: UniCaption(SomeControl) = s
    '
    ' This is known to work on Form, MDIForm, Checkbox, CommandButton, Frame, & OptionButton.
    ' Other controls are not known.
    '
    ' As a tip, build your Unicode caption using ChrW.
    ' Also note the careful way we pass the string to the unicode API call to circumvent VB6's auto-ASCII-conversion.
    DefWindowProcW ctrl.hWnd, WM_SETTEXT, 0&, ByVal StrPtr(sUniCaption)
End Property

Public Property Get UniCaption(ctrl As Object) As String
    Const WM_GETTEXT As Long = &HD
    Const WM_GETTEXTLENGTH As Long = &HE
    ' USAGE: s = UniCaption(SomeControl)
    '
    ' This is known to work on Form, MDIForm, Checkbox, CommandButton, Frame, & OptionButton.
    ' Other controls are not known.
    Dim lLen As Long
    Dim lPtr As Long
    '
    lLen = DefWindowProcW(ctrl.hWnd, WM_GETTEXTLENGTH, 0&, ByVal 0&) ' Get length of caption.
    If lLen Then ' Must have length.
        lPtr = SysAllocStringLen(0&, lLen) ' Create a BSTR of that length.
        PutMem4 ByVal VarPtr(UniCaption), ByVal lPtr ' Make the property return the BSTR.
        DefWindowProcW ctrl.hWnd, WM_GETTEXT, lLen + 1&, ByVal lPtr ' Call the default Unicode window procedure to fill the BSTR.
    End If
End Property

Public Property Let UniClipboard(sUniText As String)
    ' Puts a VB string in the clipboard without converting it to ASCII.
    Dim iStrPtr As Long
    Dim iLen As Long
    Dim iLock As Long
    Const GMEM_MOVEABLE As Long = &H2
    Const GMEM_ZEROINIT As Long = &H40
    Const CF_UNICODETEXT As Long = &HD
    '
    OpenClipboard 0&
    EmptyClipboard
    iLen = LenB(sUniText) + 2&
    iStrPtr = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, iLen)
    iLock = GlobalLock(iStrPtr)
    lstrcpy iLock, StrPtr(sUniText)
    GlobalUnlock iStrPtr
    SetClipboardData CF_UNICODETEXT, iStrPtr
    CloseClipboard
End Property

Public Property Get UniClipboard() As String
    ' Gets a UNICODE string from the clipboard and puts it in a standard VB string (which is UNICODE).
    Dim iStrPtr As Long
    Dim iLen As Long
    Dim iLock As Long
    Dim sUniText As String
    Const CF_UNICODETEXT As Long = 13&
    '
    OpenClipboard 0&
    If IsClipboardFormatAvailable(CF_UNICODETEXT) Then
        iStrPtr = GetClipboardData(CF_UNICODETEXT)
        If iStrPtr Then
            iLock = GlobalLock(iStrPtr)
            iLen = GlobalSize(iStrPtr)
            sUniText = String$(iLen \ 2& - 1&, vbNullChar)
            lstrcpy StrPtr(sUniText), iLock
            GlobalUnlock iStrPtr
        End If
        UniClipboard = sUniText
    End If
    CloseClipboard
End Property

Public Sub SetupRichTextboxForUnicode(rtb As RichTextBox)
    ' Call this so that the rtb doesn't try to do any RTF interpretation.  We will just be using it for Unicode display.
    ' Once this is called, the following two procedures will work with the rtb.
    Const TM_PLAINTEXT As Long = 1&
    Const EM_SETTEXTMODE As Long = &H459
    SendMessage rtb.hWnd, EM_SETTEXTMODE, TM_PLAINTEXT, 0& ' Set the control to use "plain text" mode so RTF isn't interpreted.
End Sub

Public Property Let RichTextboxUniText(rtb As RichTextBox, sUniText As String)
    ' Usage: Just assign any VB6 string to the rtb.
    '        If the string contains Unicode (which VB6 strings are capable of), it will be correctly handled.
    Dim stUnicode As SETTEXTEX
    Const EM_SETTEXTEX As Long = &H461
    Const RTBC_DEFAULT As Long = 0&
    Const CP_UNICODE As Long = 1200&
    '
    stUnicode.flags = RTBC_DEFAULT ' This could be otherwise.
    stUnicode.codepage = CP_UNICODE
    SendMessageWLng rtb.hWnd, EM_SETTEXTEX, VarPtr(stUnicode), StrPtr(sUniText)
End Property

Public Property Get RichTextboxUniText(rtb As RichTextBox) As String
    Dim uGTL As GETTEXTLENGTHEX
    Dim uGT As GETTEXTEX
    Dim iChars As Long
    Const EM_GETTEXTEX As Long = &H45E
    Const EM_GETTEXTLENGTHEX As Long = &H45F
    Const CP_UNICODE As Long = 1200&
    Const GTL_USECRLF As Long = 1&
    Const GTL_PRECISE As Long = 2&
    Const GTL_NUMCHARS As Long = 8&
    Const GT_USECRLF As Long = 1&
    '
    uGTL.flags = GTL_USECRLF Or GTL_PRECISE Or GTL_NUMCHARS
    uGTL.codepage = CP_UNICODE
    iChars = SendMessageWLng(rtb.hWnd, EM_GETTEXTLENGTHEX, VarPtr(uGTL), 0&)
    '
    uGT.cb = (iChars + 1) * 2
End Property
