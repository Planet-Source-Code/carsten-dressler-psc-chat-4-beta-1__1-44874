VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmNewAccount 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creating A New Account"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4770
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create Now"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   2880
      Width           =   1935
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4683
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      MouseIcon       =   "frmNewAccount.frx":038A
      TabCaption(0)   =   " User Login"
      TabPicture(0)   =   "frmNewAccount.frx":03A6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   " Profile"
      TabPicture(1)   =   "frmNewAccount.frx":0740
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblText(6)"
      Tab(1).Control(1)=   "lblText(5)"
      Tab(1).Control(2)=   "lblText(4)"
      Tab(1).Control(3)=   "lblText(3)"
      Tab(1).Control(4)=   "lblText(2)"
      Tab(1).Control(5)=   "txtProfile(5)"
      Tab(1).Control(6)=   "txtProfile(4)"
      Tab(1).Control(7)=   "txtProfile(3)"
      Tab(1).Control(8)=   "txtProfile(2)"
      Tab(1).Control(9)=   "txtProfile(1)"
      Tab(1).ControlCount=   10
      Begin VB.Frame Frame1 
         Caption         =   "Account"
         Height          =   1815
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   3975
         Begin VB.TextBox txtRepass 
            Height          =   315
            Left            =   1080
            TabIndex        =   5
            Top             =   1320
            Width           =   2655
         End
         Begin VB.TextBox txtID 
            Height          =   315
            Left            =   1080
            TabIndex        =   1
            Top             =   360
            Width           =   2655
         End
         Begin VB.TextBox txtPass 
            Height          =   315
            Left            =   1080
            TabIndex        =   3
            Top             =   840
            Width           =   2655
         End
         Begin VB.Label lblText 
            AutoSize        =   -1  'True
            Caption         =   "Retype Pas:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   7
            Left            =   150
            TabIndex        =   18
            Top             =   1350
            Width           =   795
         End
         Begin VB.Label lblText 
            AutoSize        =   -1  'True
            Caption         =   "User Name:"
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   405
            Width           =   840
         End
         Begin VB.Label lblText 
            AutoSize        =   -1  'True
            Caption         =   " Password:"
            Height          =   210
            Index           =   1
            Left            =   105
            TabIndex        =   11
            Top             =   840
            Width           =   840
         End
      End
      Begin VB.TextBox txtProfile 
         Height          =   315
         Index           =   1
         Left            =   -73920
         TabIndex        =   2
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox txtProfile 
         Height          =   315
         Index           =   2
         Left            =   -73920
         TabIndex        =   9
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox txtProfile 
         Height          =   315
         Index           =   3
         Left            =   -73920
         TabIndex        =   8
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox txtProfile 
         Height          =   315
         Index           =   4
         Left            =   -73920
         TabIndex        =   7
         Top             =   1560
         Width           =   3255
      End
      Begin VB.TextBox txtProfile 
         Height          =   315
         Index           =   5
         Left            =   -73920
         TabIndex        =   6
         Top             =   1920
         Width           =   3255
      End
      Begin VB.Label lblText 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Full Name:"
         Height          =   210
         Index           =   2
         Left            =   -74880
         TabIndex        =   17
         Top             =   510
         Width           =   855
      End
      Begin VB.Label lblText 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Age:"
         Height          =   210
         Index           =   3
         Left            =   -74880
         TabIndex        =   16
         Top             =   915
         Width           =   855
      End
      Begin VB.Label lblText 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Hobbies:"
         Height          =   210
         Index           =   4
         Left            =   -74880
         TabIndex        =   15
         Top             =   1245
         Width           =   855
      End
      Begin VB.Label lblText 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Skills:"
         Height          =   210
         Index           =   5
         Left            =   -74880
         TabIndex        =   14
         Top             =   1590
         Width           =   855
      End
      Begin VB.Label lblText 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Email:"
         Height          =   210
         Index           =   6
         Left            =   -74430
         TabIndex        =   13
         Top             =   1980
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmNewAccount"
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

Private Sub cmdCreate_Click()
   'Check if Winsock Connected
   If frmMain.Winsock1.State = 7 Then
      
      'Puts the Users Profile Together
      Dim strFinal As String
      Dim lngA As Long
      
      For lngA = 1 To txtProfile.Count
         DoEvents
         strFinal = strFinal & "/Ö/" & txtProfile(lngA).Text
      Next
      
      'Sends Msg and Updates statusbar
      frmMain.Winsock1.SendData "NewAccount/Ñ/" & txtID.Text & "/Ñ/" & txtPass.Text & "/Ñ/" & strFinal & "/Ñ/" & intPicChange
      frmMain.StaBar.Panels(1).Text = "Creating Account..."
      
      'Updates the MyName String
      frmMain.MyName = txtID.Text
      
      'Update MyInfo on frmMain
      frmMain.MyInfo = txtID.Text & "/Ö/" & txtPass.Text & strFinal
      
      Unload Me
    Else
      'Attempts to connect again
      frmMain.Winsock1.Close
      frmMain.Winsock1.Connect frmMain.cmbAddress.Text, "15009"
      
      'Tells users not connected
      MsgBox "Psc Chat wasnt able to connect to a server!", vbExclamation, "Psc Chat 4.0"
   End If
End Sub



