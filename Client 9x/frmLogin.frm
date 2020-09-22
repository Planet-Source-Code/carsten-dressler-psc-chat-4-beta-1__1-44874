VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Psc Chat 4.0 Login"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4755
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Height          =   360
      Left            =   3120
      TabIndex        =   3
      Top             =   2880
      Width           =   1455
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
      TabHeight       =   520
      MouseIcon       =   "frmLogin.frx":038A
      TabCaption(0)   =   " User Login"
      TabPicture(0)   =   "frmLogin.frx":03A6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   " Profile"
      TabPicture(1)   =   "frmLogin.frx":0740
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblText(2)"
      Tab(1).Control(1)=   "lblText(3)"
      Tab(1).Control(2)=   "lblText(4)"
      Tab(1).Control(3)=   "lblText(5)"
      Tab(1).Control(4)=   "lblText(6)"
      Tab(1).Control(5)=   "txtProfile(1)"
      Tab(1).Control(6)=   "txtProfile(2)"
      Tab(1).Control(7)=   "txtProfile(3)"
      Tab(1).Control(8)=   "txtProfile(4)"
      Tab(1).Control(9)=   "txtProfile(5)"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   " Picture"
      TabPicture(2)   =   "frmLogin.frx":0ADA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "picme"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdBrowse"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdTakePic"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdClearPic"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.CommandButton cmdClearPic 
         Caption         =   "Clear Picture"
         Height          =   375
         Left            =   -72705
         TabIndex        =   19
         Top             =   1185
         Width           =   1935
      End
      Begin VB.CommandButton cmdTakePic 
         Caption         =   "Take a Picture"
         Height          =   375
         Left            =   -72705
         TabIndex        =   18
         Top             =   825
         Width           =   1935
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Broswe for Picture"
         Height          =   375
         Left            =   -72705
         TabIndex        =   17
         Top             =   465
         Width           =   1935
      End
      Begin VB.TextBox txtProfile 
         Height          =   315
         Index           =   5
         Left            =   -73920
         TabIndex        =   16
         Top             =   1920
         Width           =   3255
      End
      Begin VB.TextBox txtProfile 
         Height          =   315
         Index           =   4
         Left            =   -73920
         TabIndex        =   15
         Top             =   1560
         Width           =   3255
      End
      Begin VB.TextBox txtProfile 
         Height          =   315
         Index           =   3
         Left            =   -73920
         TabIndex        =   14
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox txtProfile 
         Height          =   315
         Index           =   2
         Left            =   -73920
         TabIndex        =   13
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox txtProfile 
         Height          =   315
         Index           =   1
         Left            =   -73920
         TabIndex        =   12
         Top             =   480
         Width           =   3255
      End
      Begin VB.Frame Frame1 
         Caption         =   "Account"
         Height          =   1455
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   3975
         Begin VB.TextBox txtPass 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1080
            PasswordChar    =   "*"
            TabIndex        =   2
            Top             =   840
            Width           =   2655
         End
         Begin VB.TextBox txtID 
            Height          =   315
            Left            =   1080
            TabIndex        =   1
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label lblText 
            AutoSize        =   -1  'True
            Caption         =   " Password:"
            Height          =   210
            Index           =   1
            Left            =   105
            TabIndex        =   6
            Top             =   840
            Width           =   840
         End
         Begin VB.Label lblText 
            AutoSize        =   -1  'True
            Caption         =   "User Name:"
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   405
            Width           =   840
         End
      End
      Begin VB.Image picme 
         BorderStyle     =   1  'Fixed Single
         Height          =   1215
         Left            =   -74880
         Stretch         =   -1  'True
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   $"frmLogin.frx":342C
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -74760
         TabIndex        =   20
         Top             =   1920
         Width           =   3975
      End
      Begin VB.Label lblText 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Email:"
         Height          =   210
         Index           =   6
         Left            =   -74430
         TabIndex        =   11
         Top             =   1980
         Width           =   405
      End
      Begin VB.Label lblText 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Skills:"
         Height          =   210
         Index           =   5
         Left            =   -74880
         TabIndex        =   10
         Top             =   1590
         Width           =   855
      End
      Begin VB.Label lblText 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Hobbies:"
         Height          =   210
         Index           =   4
         Left            =   -74880
         TabIndex        =   9
         Top             =   1245
         Width           =   855
      End
      Begin VB.Label lblText 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Age:"
         Height          =   210
         Index           =   3
         Left            =   -74880
         TabIndex        =   8
         Top             =   915
         Width           =   855
      End
      Begin VB.Label lblText 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Full Name:"
         Height          =   210
         Index           =   2
         Left            =   -74880
         TabIndex        =   7
         Top             =   510
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmLogin"
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

Public intPicChange As Integer

Private Sub cmdBrowse_Click()
    On Error Resume Next
   frmMain.cd1.Filter = "Jpg/Gif|*.gif;*.jpg"
   frmMain.cd1.ShowOpen
   If Len(frmMain.cd1.filename) Then
        If FileLen(frmMain.cd1.filename) > 15360 Then
             intPicChange = 1
             picme.Picture = LoadPicture(frmMain.cd1.filename)
        Else
             intPicChange = 1
             MsgBox "Invalid Pic.. make sure its a jpg/gif under 15k.", , "IntraDream"
        End If
    End If
End Sub

Private Sub cmdClearPic_Click()
  intPicChange = 0
End Sub

Private Sub cmdLogin_Click()
   'Check if Winsock Connected
   If frmMain.Winsock1.State = 7 Then
      
      'Puts the Users Profile Together
      Dim strFinal As String
      Dim lngA As Long
      
      For lngA = 1 To txtProfile.Count
         DoEvents
         strFinal = strFinal & "/Ö/" & txtProfile(lngA).Text
      Next
      
      'if Pic Changes then change frmMainPIC
      If intPicChange = 1 Then
         frmMain.tmppic.Picture = frmLogin.picme.Picture
      End If
      
      'Sends Msg and Updates statusbar
      frmMain.Winsock1.SendData "Login/Ñ/" & txtID.Text & "/Ñ/" & txtPass.Text & "/Ñ/" & strFinal & "/Ñ/" & intPicChange & "/Ñ/" & frmMain.mnuNotifyReg.Checked
      frmMain.StaBar.Panels(1).Text = "Logining in..."
      
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


Private Sub cmdTakePic_Click()
   frmTakePicture.Show
End Sub

Private Sub Form_Load()
   'IntPicChange is used to let the Server
   'Know if the Picture has changed; once
   'You Click the TakePic/Broswe for pic
   'The intpicchange will change to 1
   intPicChange = 0
End Sub
