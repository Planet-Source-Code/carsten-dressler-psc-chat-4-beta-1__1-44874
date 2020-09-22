VERSION 5.00
Begin VB.Form frmTakePicture 
   Caption         =   "Take a Picture"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3165
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTakePicture.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2805
   ScaleWidth      =   3165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTakePic 
      Caption         =   "Take Picture"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Timer tmrFrames 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Live"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.Image piclive 
         BorderStyle     =   1  'Fixed Single
         Height          =   1695
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmTakePicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub cmdTakePic_Click()
   frmLogin.picme.Picture = Me.piclive
   frmLogin.intPicChange = 1
   frmLogin.SetFocus
   Unload Me
End Sub

Private Sub Form_Load()
   'Starts Takeing Pictures
   tmrFrames.Enabled = True
   StartCam
End Sub

Private Sub tmrFrames_Timer()
    On Error Resume Next
    CamToClip
    piclive.Picture = Clipboard.GetData
End Sub
