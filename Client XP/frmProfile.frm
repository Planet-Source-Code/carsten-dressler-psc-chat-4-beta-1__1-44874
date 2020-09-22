VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmProfile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "< > Profile - PSC Chat 4"
   ClientHeight    =   2775
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6240
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProfile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "      Profile"
      Height          =   2055
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   3495
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "xxx"
         Height          =   210
         Index           =   5
         Left            =   1440
         TabIndex        =   11
         Top             =   1680
         Width           =   270
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "xxx"
         Height          =   210
         Index           =   4
         Left            =   1440
         TabIndex        =   10
         Top             =   1320
         Width           =   270
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "xxx"
         Height          =   210
         Index           =   3
         Left            =   1440
         TabIndex        =   9
         Top             =   960
         Width           =   270
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "xxx"
         Height          =   210
         Index           =   2
         Left            =   1440
         TabIndex        =   8
         Top             =   600
         Width           =   270
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "xxx"
         Height          =   210
         Index           =   1
         Left            =   1440
         TabIndex        =   7
         Top             =   240
         Width           =   270
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         Caption         =   "Full Name:"
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   315
         Width           =   855
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         Caption         =   "Age:"
         Height          =   210
         Index           =   3
         Left            =   240
         TabIndex        =   5
         Top             =   645
         Width           =   855
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         Caption         =   "Hobbies:"
         Height          =   210
         Index           =   4
         Left            =   240
         TabIndex        =   4
         Top             =   975
         Width           =   855
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         Caption         =   "Skills:"
         Height          =   210
         Index           =   5
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         Caption         =   "Email:"
         Height          =   210
         Index           =   6
         Left            =   285
         TabIndex        =   2
         Top             =   1710
         Width           =   405
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   105
         Picture         =   "frmProfile.frx":038A
         Top             =   -15
         Width           =   240
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Image Picture1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1965
      Left            =   120
      Picture         =   "frmProfile.frx":0714
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2430
   End
End
Attribute VB_Name = "frmProfile"
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

Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Inet1.Cancel
  Close #200
End Sub

Private Sub lblItem_Click(Index As Integer)
  Select Case Index
     Case 1
       MsgBox lblItem(1).Caption
     Case 2
       MsgBox lblItem(2).Caption
     Case 3
       MsgBox lblItem(3).Caption
     Case 4
       MsgBox lblItem(4).Caption
     Case 5
       MsgBox lblItem(5).Caption
  End Select
End Sub
