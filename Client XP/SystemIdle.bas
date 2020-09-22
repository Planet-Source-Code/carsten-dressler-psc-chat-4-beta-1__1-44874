Attribute VB_Name = "SystemIdle"
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Type POINTAPI
    x As Long
    y As Long
End Type
Dim intOldX As Integer
Dim intOldY As Integer
Dim intX As Integer
Dim intY As Integer
Dim intA As Integer
Dim intState As Integer
Dim intKeyBoardIdle As Integer
Dim intMouseIdle As Integer
Dim lngKeyBoardIdleCheck As Long
Dim lngMouseIdleCheck As Long

Public Function GetX() As Long
    Dim objPoint As POINTAPI
    GetCursorPos objPoint
    GetX = objPoint.x
End Function
Public Function GetY() As Long
    Dim objPoint As POINTAPI
    GetCursorPos objPoint
    GetY = objPoint.y
End Function

Public Function SystemIdleCheck() As Boolean

   'Checks If KeyBoard is used
   For intA = 1 To 256
      'DoEvents
      intState = GetAsyncKeyState(intA)
      
      'If KeyBoard Key is down
      If intState = -32767 Then
         intKeyBoardIdle = 0
         lngKeyBoardIdleCheck = 0
         Exit For
      End If
   Next intA
   
   If intState <> -32767 Then
      lngKeyBoardIdleCheck = lngKeyBoardIdleCheck + 1
   End If
   
   If lngKeyBoardIdleCheck > 20 Then
      intKeyBoardIdle = 1
   End If
   
   'Starts Mouse Idle Check
   intX = GetX
   intY = GetY
   
   If intOldX <> intX And intOldY <> intY Then
      intMouseIdle = 0
      lngMouseIdleCheck = 0
   End If
   
   If intOldX = intX And intOldY = intY Then
      lngMouseIdleCheck = lngMouseIdleCheck + 1
   End If
   
   If lngMouseIdleCheck > 20 Then
      intMouseIdle = 1
   End If
   
   intOldX = GetX
   intOldY = GetY
      
   If intKeyBoardIdle = 0 Or intMouseIdle = 0 Then
       SystemIdleCheck = False
     Else
       SystemIdleCheck = True
   End If

  
End Function

