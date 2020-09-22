Attribute VB_Name = "URLDetection"
'---------------------------------
'Copyright (C) IntraDream, Inc 2002 - 2003
'If you have any questions Please Contact
'IntraDreams Support; Support@intradream.com
'
'This program is free software; you can redistribute
'it and/or modify it under the terms of the GNU General
'Public License as published by the Free Software
'Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will
'be useful, but WITHOUT ANY WARRANTY; without even the
'implied warranty of MERCHANTABILITY or FITNESS FOR A
'PARTICULAR PURPOSE. See the GNU General Public License
'for more details.
'
'You should have received a copy of the GNU General Public
'License along with this program; if not, write to the Free
'Software Foundation, Inc., 59 Temple Place, Suite 330,
'Boston, MA 02111-1307 USA
'----------------------------------

Option Explicit
Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
 
Public Const GWL_WNDPROC = (-4)
Public Const WM_USER = &H400
Public Const WM_NOTIFY = &H4E
Public Const WM_LBUTTONDOWNL = &H201
Public Const EM_GETEVENTMASK = WM_USER + 59
Public Const EM_GETTEXTRANGE = WM_USER + 75
Public Const EM_SETEVENTMASK = WM_USER + 69
Public Const EM_AUTOURLDETECT = WM_USER + 91
Public Const EN_LINK = &H70B
Public Const ENM_LINK = &H4000000
Public Const SW_SHOWNORMAL = 1

Type tagNMHDR
    hwndFrom As Long
    idFrom   As Long
    code     As Long
End Type

Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type

Type ENLINK
    nmhdr  As tagNMHDR
    msg    As Long
    wParam As Long
    lParam As Long
    chrg   As CHARRANGE
End Type

Type TEXTRANGE
    chrg      As CHARRANGE
    lpstrText As Long
End Type

Public glnglpOriginalWndProc As Long
Public glngOriginalhWnd As Long

Function RichTextBoxSubProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim udtNMHDR               As tagNMHDR
Dim udtENLINK              As ENLINK
Dim udtTEXTRANGE           As TEXTRANGE
Dim strBuffer              As String * 128
Dim strOperation           As String
Dim strFileName            As String
Dim strDefaultDirectory    As String
Dim lngHInstanceExecutable As Long
Dim lngWin32apiResultCode  As Long
 
 
 If uMsg = WM_NOTIFY Then
    RtlMoveMemory udtNMHDR, ByVal lParam, Len(udtNMHDR)
    If udtNMHDR.hwndFrom = frmMain.RTBRecive.hWnd And udtNMHDR.code = EN_LINK Then
        RtlMoveMemory udtENLINK, ByVal lParam, Len(udtENLINK)
        If udtENLINK.msg = WM_LBUTTONDOWN Then
            strBuffer = ""
            
            With udtTEXTRANGE
                .chrg.cpMin = udtENLINK.chrg.cpMin
                .chrg.cpMax = udtENLINK.chrg.cpMax
                .lpstrText = StrPtr(strBuffer)
            End With
 
            With frmMain.RTBRecive
                lngWin32apiResultCode = SendMessage(.hWnd, EM_GETTEXTRANGE, 0, udtTEXTRANGE)
            End With

            RtlMoveMemory ByVal strBuffer, ByVal udtTEXTRANGE.lpstrText, Len(strBuffer)
            strOperation = "open"
            strFileName = strBuffer
            lngHInstanceExecutable = ShellExecute(frmMain.hWnd, strOperation, strFileName, vbNullString, strDefaultDirectory, SW_SHOWNORMAL)

        End If
    End If
End If
  
RichTextBoxSubProc = CallWindowProc(glnglpOriginalWndProc, hWnd, uMsg, wParam, lParam)

End Function

