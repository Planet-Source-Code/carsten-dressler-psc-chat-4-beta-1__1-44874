VERSION 5.00
Begin VB.Form frmFaces 
   BackColor       =   &H80000009&
   ClientHeight    =   510
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4365
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   510
   ScaleWidth      =   4365
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   8
      Left            =   3960
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmFaces.frx":0000
      Top             =   120
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   7
      Left            =   3480
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmFaces.frx":0342
      Top             =   120
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   6
      Left            =   3000
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmFaces.frx":0684
      Top             =   120
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   5
      Left            =   2520
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmFaces.frx":09C6
      Top             =   120
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   4
      Left            =   2040
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmFaces.frx":0D08
      Top             =   120
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   3
      Left            =   1560
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmFaces.frx":104A
      Top             =   120
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   2
      Left            =   1080
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmFaces.frx":138C
      Top             =   120
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   1
      Left            =   600
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmFaces.frx":16CE
      Top             =   120
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   0
      Left            =   120
      MousePointer    =   10  'Up Arrow
      Picture         =   "frmFaces.frx":1A10
      Top             =   120
      Width           =   240
   End
End
Attribute VB_Name = "frmFaces"
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

Private Sub imgIcon_Click(Index As Integer)
  Select Case Index
    Case 0
      frmMain.RTBSend.Text = frmMain.RTBSend.Text & " :) "
    Case 1
      frmMain.RTBSend.Text = frmMain.RTBSend.Text & " :( "
    Case 2
      frmMain.RTBSend.Text = frmMain.RTBSend.Text & " :l "
    Case 3
      frmMain.RTBSend.Text = frmMain.RTBSend.Text & " :o "
    Case 4
      frmMain.RTBSend.Text = frmMain.RTBSend.Text & " :z "
    Case 5
      frmMain.RTBSend.Text = frmMain.RTBSend.Text & " :t "
    Case 6
      frmMain.RTBSend.Text = frmMain.RTBSend.Text & " ;@ "
    Case 7
      frmMain.RTBSend.Text = frmMain.RTBSend.Text & " |) "
    Case 8
      frmMain.RTBSend.Text = frmMain.RTBSend.Text & " :k "
          
  End Select
  
  Unload Me
End Sub
':),:(,:o,;@,:l,:<,:t,|),:z,:k
