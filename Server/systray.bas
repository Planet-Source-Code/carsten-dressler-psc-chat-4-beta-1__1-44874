Attribute VB_Name = "systray"
Public Type NOTIFYICONDATA
   cbSize As Long
   hwnd As Long
   uID As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip As String * 128
   dwState As Long
   dwStateMask As Long
   szInfo As String * 256
   uTimeout As Long
   szInfoTitle As String * 64
   dwInfoFlags As Long
End Type

 Public m_IconData As NOTIFYICONDATA
Public Const NOTIFYICON_VERSION = 3       'V5 style taskbar
Public Const NOTIFYICON_OLDVERSION = 0    'Win95 style taskbar

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2

Public Const NIM_SETFOCUS = &H3
Public Const NIM_SETVERSION = &H4

Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Public Const NIF_STATE = &H8
Public Const NIF_INFO = &H10
 

Public Const NIS_HIDDEN = &H1
Public Const NIS_SHAREDICON = &H2
 

Public Const NIIF_NONE = &H0
Public Const NIIF_WARNING = &H2
Public Const NIIF_ERROR = &H3
Public Const NIIF_INFO = &H1
 Public Const NIIF_GUID = &H4


Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
 
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" _
(ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public Function DoConv(Number As String)
Dim DoConv3 As String
DoConv3 = Number
Dim part As Variant
On Error Resume Next
    If DoConv3 > 1023 Then
        DoConv3 = DoConv3 / 1024
    Else
        DoConv = dec(DoConv3, 2) & " B"
        Exit Function
    End If
    If DoConv3 > 1023 Then
        DoConv3 = DoConv3 / 1024
    Else
        DoConv = dec(DoConv3, 2) & " KB"
        Exit Function
    End If
    If DoConv3 > 1023 Then
        DoConv3 = DoConv3 / 1024
    Else
        DoConv = dec(DoConv3, 2) & " MB"
        Exit Function
    End If
    If DoConv3 > 1023 Then
        DoConv3 = DoConv3 / 1024
    Else
        DoConv = dec(DoConv3, 2) & " GB"
        Exit Function
    End If
    If DoConv3 > 1023 Then
        DoConv3 = DoConv3 / 1024
    Else
        DoConv = dec(DoConv3, 2) & " TB"
        Exit Function
    End If
        If DoConv3 > 1023 Then
        DoConv3 = DoConv3 / 1024
    Else
        DoConv = dec(DoConv3, 2) & " PB"
        Exit Function
    End If
    If DoConv3 > 1023 Then
        DoConv3 = DoConv3 / 1024
    Else
        DoConv = dec(DoConv3, 2) & " EB"
        Exit Function
    End If
    If DoConv3 > 1023 Then
        DoConv3 = DoConv3 / 1024
    Else
        DoConv = dec(DoConv3, 2) & " ZB"
        Exit Function
    End If
    If DoConv3 > 1023 Then
        DoConv3 = DoConv3 / 1024
    Else
        DoConv = dec(DoConv3, 2) & " YB"
        Exit Function
    End If
    DoConv = "YOUR CRAZY"
End Function
Public Function dec(Number As String, lod As Integer)
Dim x As Variant
x = Split(Number, ".")
If UBound(x) < 1 Then
    dec = Number
Else
    If Len(x(1)) < lod Then
        dec = Number
    Else
        dec = x(0) & "." & Left(x(1), lod)
    End If
End If
End Function




