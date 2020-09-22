Attribute VB_Name = "Camera"
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
Public HasCam As Boolean
Public Declare Function BMPToJPG Lib "converter.dll" (ByVal InputFilename As String, ByVal OutputFilename As String, ByVal Quality As Long) As Integer
Public Const WM_CAP_DRIVER_CONNECT As Long = 1034
Public Const WM_CAP_DRIVER_DISCONNECT As Long = 1035
Public Const WM_CAP_GRAB_FRAME As Long = 1084
Public Const WM_CAP_EDIT_COPY As Long = 1054
Public Const WM_CAP_DLG_VIDEOFORMAT As Long = 1065
Public Const WM_CAP_DLG_VIDEOSOURCE As Long = 1066
Public Const WM_CLOSE = &H10
Public Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public mCapHwnd As Long

Public Sub StopCam()
    HasCam = False
    SendMessage mCapHwnd, WM_CLOSE, 0, 0
    SendMessage mCapHwnd, WM_CAP_DRIVER_DISCONNECT, 0, 0
End Sub

Public Sub StartCam()
    Dim Ret As Long
    If HasCam = True Then Exit Sub
    mCapHwnd = capCreateCaptureWindow("IntraDream", 0, 0, 0, 0, 0, 0, 0)
    Ret = SendMessage(mCapHwnd, WM_CAP_DRIVER_CONNECT, 0, 0)
    If Ret = 1 Then
        HasCam = True
    Else
        HasCam = False
        StopCam
    End If
End Sub

Public Sub SetCamSize()
    SendMessage mCapHwnd, WM_CAP_DLG_VIDEOFORMAT, 0, 0
End Sub

Public Sub SetCamSource()
    SendMessage mCapHwnd, WM_CAP_DLG_VIDEOSOURCE, 0, 0
End Sub

Public Function CamToBMP(filename As String)
    On Error GoTo error
    CamToBMP = 0
    SendMessage mCapHwnd, WM_CAP_GRAB_FRAME, 0, 0
    SendMessage mCapHwnd, WM_CAP_EDIT_COPY, 0, 0
    SavePicture Clipboard.GetData, filename
    Exit Function
error:
    CamToBMP = 1
End Function

Public Function CamToClip()
    On Error GoTo error
    CamToClip = 0
    SendMessage mCapHwnd, WM_CAP_GRAB_FRAME, 0, 0
    SendMessage mCapHwnd, WM_CAP_EDIT_COPY, 0, 0
    Exit Function
error:
    CamToClip = 1
End Function

Public Function CamToJPG(filename As String, Quality As Integer)
    On Error GoTo exits
    CamToJPG = 0
    If Quality > 100 Or Quality < 1 Then Quality = 100
        If CamToBMP(App.Path & "\Cam.bmp") = 0 Then
            If BMPToJPG(App.Path & "\Cam.bmp", filename, Quality) = 0 Then
                Exit Function
            End If
        End If
exits:
    CamToJPG = 1
End Function
Public Function BMPFileToJpg(filename As String, output As String, Quality As Integer)
    On Error Resume Next
    If Quality > 100 Or Quality < 1 Then Quality = 100
    BMPToJPG filename, output, Quality
    
End Function

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
    DoConv = "N/A"
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

