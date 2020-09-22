VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPrivateM 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "<CDressler> Private Massage"
   ClientHeight    =   5055
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   8340
   Icon            =   "frmPrivateM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   7920
      Top             =   3120
   End
   Begin VB.CommandButton cmdCancelTran 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5670
      TabIndex        =   23
      Top             =   2280
      Width           =   2685
   End
   Begin VB.PictureBox picAccept 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   5655
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton cmdDecline 
         Caption         =   "Decline"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   9
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdAccept 
         Caption         =   "Accept"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   8
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblAccept 
         BackStyle       =   0  'Transparent
         Caption         =   "Accept Video Transmission?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   2175
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   8040
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   5760
      Max             =   100
      Min             =   1
      TabIndex        =   5
      Top             =   2640
      Value           =   40
      Width           =   2595
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   5760
      TabIndex        =   4
      Top             =   3000
      Width           =   2580
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Receiving FPS :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Sending FPS :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Bytes Received :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Bytes Sent :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   18
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   16
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   15
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   14
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Send Speed :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   12
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Receive Speed :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   1440
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6960
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrivateM.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrivateM.frx":0724
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrivateM.frx":0ABE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrivateM.frx":0E58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrivateM.frx":11F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrivateM.frx":158C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrivateM.frx":1926
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrivateM.frx":1CC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrivateM.frx":205A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrivateM.frx":23F4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7560
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4440
      TabIndex        =   2
      Top             =   3720
      Width           =   1140
   End
   Begin RichTextLib.RichTextBox rtbS 
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   4200
      Width           =   5670
      _ExtentX        =   10001
      _ExtentY        =   1296
      _Version        =   393217
      MultiLine       =   0   'False
      TextRTF         =   $"frmPrivateM.frx":4D46
   End
   Begin RichTextLib.RichTextBox rtbR 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   6376
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   2
      TextRTF         =   $"frmPrivateM.frx":4DC8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar 
      Height          =   540
      Left            =   0
      TabIndex        =   3
      Top             =   3600
      Width           =   5670
      _ExtentX        =   10001
      _ExtentY        =   953
      ButtonWidth     =   1191
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Font"
            Key             =   "Font"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Color"
            Key             =   "color"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Size"
            Key             =   "size"
            ImageIndex      =   3
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   7
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "s8"
                  Text            =   "8"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "s9"
                  Text            =   "9"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "s10"
                  Text            =   "10"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "s11"
                  Text            =   "11"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "s12"
                  Text            =   "12"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "s13"
                  Text            =   "13"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "s14"
                  Text            =   "14"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Wboard"
            Key             =   "Wboard"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cam"
            Key             =   "cam"
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSetSource 
      Caption         =   "Select Source"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5670
      TabIndex        =   22
      Top             =   2040
      Width           =   2685
   End
   Begin VB.Image Load 
      Height          =   615
      Left            =   5715
      Picture         =   "frmPrivateM.frx":4E3F
      Stretch         =   -1  'True
      Top             =   60
      Width           =   735
   End
   Begin VB.Image Receive 
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   5685
      Picture         =   "frmPrivateM.frx":13AA5
      Stretch         =   -1  'True
      Top             =   30
      Width           =   2655
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
      End
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Begin VB.Menu mnuWBoard 
         Caption         =   "White Board"
      End
      Begin VB.Menu mnuCam 
         Caption         =   "Start Camera"
      End
      Begin VB.Menu mnuSp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCamara 
         Caption         =   "Camera"
         Begin VB.Menu mnuSource 
            Caption         =   "Select Video Source"
         End
         Begin VB.Menu mnuSize 
            Caption         =   "Select Video Size"
         End
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmPrivateM"
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

Dim Rec As String
Dim BRec As String
Dim BSent As String
Dim CalcSpeed As String
Dim CalcUpSpeed As String
Dim Fpsr As String
Dim Fpss As String
Public strPort As String
Public strIP As String

Private Sub cmdAccept_Click()
   'Disables the Appect button and renames other button
   Me.cmdDecline.Caption = "Cancel"
   Me.cmdAccept.Enabled = False
   
   BRec = 0
   BSent = 0
   CalcSpeed = 0
   CalcUpSpeed = 0
   
   StartCam
   Me.Width = 8520
   
   'Connects
    Winsock1.Close
    Winsock1.Connect strIP, strPort
    
End Sub

Private Sub cmdCancelTran_Click()
   Winsock1.Close
   Me.Width = 5755
End Sub

Private Sub cmdDecline_Click()
   Select Case cmdDecline.Caption
   Case "Decline"
   picAccept.Visible = False
      
   Dim strName1 As String
       strName1 = Right(Replace(Me.Caption, " > Private Message", ""), Len(Replace(Me.Caption, " > Private Message", "")) - 2)
       
       If frmMain.Winsock1.State = 7 Then
         frmMain.Winsock1.SendData "CamDeclined/Ñ/" & strName
       End If
     
   Case "Cancel"
      picAccept.Visible = False
   End Select
End Sub

Private Sub cmdSend_Click()
   'Checks if Extra Command Are Entered
   If rtbS.Text = "\WBoard" Then
       
       Dim strName1 As String
       strName1 = Right(Replace(Me.Caption, " > Private Message", ""), Len(Replace(Me.Caption, " > Private Message", "")) - 2)
       
       If frmMain.Winsock1.State = 7 Then
         frmMain.Winsock1.SendData "CamDeclined/Ñ/" & strName
       End If
   End If
   
   'Chekcs to see if anything is in RTBS
   If rtbS.Text <> "" Then
     Dim strSend As String
     
     'WARNING: IF YOU REMOVE THIS YOU WILL
     'BE BANNED FROM THE SERVER!
     strSend = Replace(rtbS.TextRTF, vbNewLine, "|ÿ|?|Ö|?|ÿ|")
     
     'Phases UserName from Caption
     Dim strName2 As String
     strName2 = Right(Replace(Me.Caption, " > Private Message", ""), Len(Replace(Me.Caption, " > Private Message", "")) - 2)
    
     'Checks to see if Winsock is connected
     If frmMain.Winsock1.State = 7 Then
         frmMain.Winsock1.SendData "PM/Ñ/" & strName2 & "/Ñ/" & strSend
     End If
     
     'Adds Name Where it came from and
     'Adds the Message to the RTBRecive
     Dim lngStart As Long
     lngStart = InStr(1, frmMain.MyInfo, "/Ö/")
          
     rtbR.SelColor = vbBlue
     rtbR.SelText = "<" & Left(frmMain.MyInfo, lngStart - 1) & "> "
             
     rtbR.SelRTF = rtbS.TextRTF
     rtbR.SelStart = Len(rtbR.Text)
     
     'Clearing the RTBSend
     rtbS.Text = ""
   End If
End Sub

Private Sub cmdSetSource_Click()
   SetCamSource
End Sub



Private Sub Form_Load()
   Me.Width = 5755
   
End Sub

Private Sub mnuAbout_Click()
  frmAbout.Show
  
End Sub

Private Sub mnuCam_Click()
   If frmMain.Winsock1.State = 7 Then
              
     'Resets Caculations Data
     BRec = 0
     BSent = 0
     CalcSpeed = 0
     CalcUpSpeed = 0
                   
     Me.Width = 8505
     
     Winsock1.Close
     Winsock1.LocalPort = InputBox("Please Enter a Port")
     Winsock1.Listen
     
     frmMain.Winsock1.SendData "CamRequest/Ñ/" & Right(Replace(Me.Caption, " > Private Message", ""), Len(Replace(Me.Caption, " > Private Message", "")) - 2) & "/Ñ/" & Winsock1.LocalPort
  End If
End Sub

Private Sub mnuexit_Click()
   Unload Me
End Sub

Private Sub mnuWBoard_Click()
   'Gets Users Name
   Dim strName As String
   strName = Right(Replace(Me.Caption, " > Private Message", ""), Len(Replace(Me.Caption, " > Private Message", "")) - 2)
             
   'Loads Whiteboard Form
   If frmWhiteBoard.Visible = False Then
      
       'Asks for a Port to use
       Dim intPort As Integer
       intPort = InputBox("Please Enter a Open Port")
               
       frmWhiteBoard.SckServer.LocalPort = intPort
       frmWhiteBoard.SckServer.Listen
      
       If Winsock1.State = 7 Then
          Winsock1.SendData "WhiteBoard/Ñ/" & strName & "/Ñ/" & intPort
       End If
      
       frmWhiteBoard.Caption = "< " & strName & " > WhiteBoard"
       frmWhiteBoard.Show
      
       frmWhiteBoard.intType = 1
       Exit Sub
     Else
       MsgBox "WhiteBoard is in use", vbOKCancel + vbCritical
   End If
End Sub

Private Sub rtbR_Change()
  On Error Resume Next
  rtbR.SelStart = Len(rtbR.Text)
End Sub

Private Sub rtbR_KeyPress(KeyAscii As Integer)
  If KeyAscii = 3 Then Exit Sub
  rtbS.SetFocus
  rtbS.Text = rtbR.Text & Chr(KeyAscii)
  rtbS.SelStart = Len(rtbR.Text)
  KeyAscii = 0
  
  RTBRecive.SelStart = Len(RTBRecive.Text)
End Sub

Private Sub rtbS_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmdSend_Click
   End If
End Sub

Private Sub TBar_ButtonClick(ByVal Button As MSComctlLib.Button)
   On Error Resume Next
   
   Select Case LCase(Button.Key)
       'Shows the Font Dialog
       Case "font"
          cd1.Flags = &H1
          cd1.ShowFont
          
          If cd1.FontName <> "" Then
             rtbS.SelFontName = cd1.FontName
          End If
          
          rtbS.SelBold = cd1.FontBold
          rtbS.SelItalic = cd1.FontItalic
          rtbS.SelFontSize = cd1.FontSize
          
       'Shows the Size Dialog
       Case "size"
          
       'Shows the Color Dialog
       Case "color"
          cd1.ShowColor
          
          'Adds the Colour to the RTB
          If cd1.Color <> "" Then
             rtbS.SelColor = cd1.Color
          End If
       Case "wboard"
             'Gets Users Name
             Dim strName As String
             strName = Right(Replace(Me.Caption, " > Private Message", ""), Len(Replace(Me.Caption, " > Private Message", "")) - 2)
             
             'Loads Whiteboard Form
             If frmWhiteBoard.Visible = False Then
      
                'Asks for a Port to use
                Dim intPort As Integer
                intPort = InputBox("Please Enter a Open Port")
                
                If intPort <> "" Then
                  frmWhiteBoard.SckServer.LocalPort = intPort
                  frmWhiteBoard.SckServer.Listen
      
                  If Winsock1.State = 7 Then
                     Winsock1.SendData "WhiteBoard/Ñ/" & strName & "/Ñ/" & intPort
                  End If
      
                  frmWhiteBoard.Caption = "< " & strName & " > WhiteBoard"
                  frmWhiteBoard.Show
      
                  frmWhiteBoard.intType = 1
                  Exit Sub
                End If
              Else
                MsgBox "WhiteBoard is in use", vbOKCancel + vbCritical
      
            End If
       Case "cam"
           If frmMain.Winsock1.State = 7 Then
              
              'Resets Caculations Data
              BRec = 0
              BSent = 0
              CalcSpeed = 0
              CalcUpSpeed = 0
              StartCam
              
              Winsock1.Close
              
              Dim intPortcam As String
              intPortcam = InputBox("Please Enter a Port")
              
              If intPortcam <> "" Then
                Me.Width = 8520
                
                StartCam
                Winsock1.LocalPort = intPortcam
                Winsock1.Listen
                frmMain.Winsock1.SendData "CamRequest/Ñ/" & Right(Replace(Me.Caption, " > Private Message", ""), Len(Replace(Me.Caption, " > Private Message", "")) - 2) & "/Ñ/" & intPortcam
              End If
           End If
   End Select
   
End Sub

Private Sub Timer1_Timer()
   On Error Resume Next
   Label2(0) = DoConv((BRec - CalcSpeed) / 3) & "/s"
   CalcSpeed = BRec
   Label2(3) = DoConv((BSent - CalcUpSpeed) / 3) & "/s"
   CalcUpSpeed = BSent
   Label2(1) = dec(Fpsr / 3, 2)
   Fpsr = 0
   Label2(4) = dec(Fpss / 3, 2)
   Fpss = 0
   Label2(2) = DoConv(BSent)
   Label2(5) = DoConv(BRec)
End Sub

Private Sub Winsock1_Close()
   Winsock1.Close
   Me.Width = 5805
End Sub

Private Sub Winsock1_Connect()
    
    Me.picAccept.Visible = False
    Me.cmdDecline.Caption = "Decline"
    SendPic
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    Winsock1.Close
    Winsock1.Accept requestID
    SendPic
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim Sata As String
On Error Resume Next
        Dim i As Integer
        Dim a As String
        Dim file As String
If Winsock1.State <> 7 Then Exit Sub
Winsock1.GetData Sata
DoEvents
BRec = BRec + bytesTotal
Dim S2 As Variant
If InStr(Right(Sata, 20), "@CAM@") Then
        Dim S1 As Variant
        S1 = Split(Sata, "@CAM@") '
        file = S1(0)
        Rec = Rec & file
        Fpsr = Fpsr + 1
        
        Open App.Path & "\" & Me.Caption & "Rec.jpg" For Binary As #2
        Put #2, , Rec
        Rec = ""
        Close #2
        
        Receive.Picture = LoadPicture(App.Path & "\" & Me.Caption & "Rec.jpg")
        Receive.Refresh
        'If UBound(S1) = 0 Then Exit Sub
        'file = S1(1)
        'a = Replace(Left(file, 4), " ´", "") & Right(file, Len(file) - 4)
        'Rec = Rec & a
        Exit Sub
End If
    a = Replace(Left(Sata, 4), " ´", "") & Right(Sata, Len(file) - 4)
    Rec = Rec & a
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Winsock1.Close
End Sub

Sub SendPic()
    On Error Resume Next
    If HasCam = False Then Exit Sub
    Dim file As String
    DoEvents
    If CamToJPG(App.Path & "\" & Me.Caption & "Cam.jpg", HScroll1.Value) = 0 Then
        Open App.Path & "\" & Me.Caption & "Cam.jpg" For Binary As #1
        file = Space$(LOF(1))
        Get #1, , file
        Close #1
        If Winsock1.State = 7 Then
            Winsock1.SendData file & "@CAM@"
        End If
    Else
        DoEvents
        SendPic
    End If
End Sub

Private Sub Winsock1_SendComplete()
    On Error Resume Next
    Load.Picture = LoadPicture(App.Path & "\" & Me.Caption & "Cam.jpg")
    Fpss = Fpss + 1
    Kill App.Path & "\" & Me.Caption & "Cam.jpg"
    SendPic
End Sub

Private Sub Winsock1_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    On Error Resume Next
    BSent = BSent + bytesSent
End Sub

