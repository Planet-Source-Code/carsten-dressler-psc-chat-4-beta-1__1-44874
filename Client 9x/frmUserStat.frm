VERSION 5.00
Begin VB.Form frmUserStat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Status"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5760
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUserStat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "Close"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "?"
      Height          =   255
      Index           =   3
      Left            =   3000
      TabIndex        =   9
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "?"
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   8
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "?"
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   7
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "?"
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   6
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Last Loged in at:"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sent Bytes Since logged in:"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Accounts Under Same Name:"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Current IP Address:"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxxxx Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmUserStat"
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

Private Sub cmdOK_Click()
  Unload Me
End Sub

