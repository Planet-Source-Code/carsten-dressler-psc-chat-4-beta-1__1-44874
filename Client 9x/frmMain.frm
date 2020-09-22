VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "IntraDream PSC Chat 4.0"
   ClientHeight    =   7050
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9075
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   9075
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrIdle 
      Interval        =   60
      Left            =   120
      Top             =   960
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   8520
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7920
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":46B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4A4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4DE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5180
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":551A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":58B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5C4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5FE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6382
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":671C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6AB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6DD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C5C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C71C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageCombo cmbStatus 
      Height          =   330
      Left            =   6480
      TabIndex        =   7
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Indentation     =   1
      Text            =   "My Status"
      ImageList       =   "ImageList1"
   End
   Begin VB.ComboBox cmbAddress 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      ItemData        =   "frmMain.frx":CAB6
      Left            =   120
      List            =   "frmMain.frx":CAC0
      TabIndex        =   6
      Text            =   "Pscchat4.ods.org"
      Top             =   120
      Width           =   6255
   End
   Begin MSComctlLib.StatusBar StaBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   6795
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2718
            MinWidth        =   2718
            Text            =   "Disconnected..."
            TextSave        =   "Disconnected..."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3598
            MinWidth        =   3598
            Text            =   "Topic: "
            TextSave        =   "Topic: "
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "Warning(s): None"
            TextSave        =   "Warning(s): None"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4763
            Text            =   "Last Message:"
            TextSave        =   "Last Message:"
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar TBar 
      Height          =   540
      Left            =   150
      TabIndex        =   4
      Top             =   5475
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   953
      ButtonWidth     =   926
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Font"
            Key             =   "Font"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Color"
            Key             =   "Color"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Size"
            Key             =   "Size"
            ImageIndex      =   13
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
            Caption         =   "Faces"
            Key             =   "Faces"
            ImageIndex      =   14
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   615
      Left            =   8280
      TabIndex        =   3
      Top             =   6120
      Width           =   735
   End
   Begin RichTextLib.RichTextBox RTBSend 
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   6120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   1085
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      TextRTF         =   $"frmMain.frx":CAE8
   End
   Begin RichTextLib.RichTextBox RTBRecive 
      Height          =   4935
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   8705
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":CB5F
   End
   Begin MSComctlLib.ListView lvUsers 
      Height          =   5535
      Left            =   6480
      TabIndex        =   0
      Top             =   480
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   9763
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.PictureBox tmppic 
      Height          =   1095
      Left            =   6480
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ImageList faces 
      Left            =   120
      Top             =   960
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
            Picture         =   "frmMain.frx":CBD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CF28
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D27A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D5CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D91E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DC70
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DFC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E314
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E666
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E9B8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox WinMP 
      Height          =   255
      Left            =   8640
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   9
      Top             =   5760
      Width           =   255
   End
   Begin VB.Menu mnuRClick 
      Caption         =   "mnuRClick"
      Visible         =   0   'False
      Begin VB.Menu mnuIgnore 
         Caption         =   "Ignore"
      End
      Begin VB.Menu mnuUnIgnore 
         Caption         =   "UnIgnore"
      End
      Begin VB.Menu mnusp2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWhiteBoard 
         Caption         =   "&WhiteBoard"
      End
      Begin VB.Menu mnuProfile 
         Caption         =   "&Profile"
      End
      Begin VB.Menu mnuS4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOps 
         Caption         =   "Operators"
         Begin VB.Menu mnuKick 
            Caption         =   "&Kick"
         End
         Begin VB.Menu mnuPBans 
            Caption         =   "Perment Bans"
         End
         Begin VB.Menu mnuBan60 
            Caption         =   "&Ban for 60 minutes"
         End
         Begin VB.Menu mnuMuzzle 
            Caption         =   "&Muzzle"
         End
         Begin VB.Menu mnuS3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuClear60Min 
            Caption         =   "Clear All 60 Min Bans"
         End
         Begin VB.Menu mnuStats 
            Caption         =   "Status"
         End
         Begin VB.Menu mnuNotifyReg 
            Caption         =   "Notify New Registrations"
         End
      End
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLogin 
         Caption         =   "Login"
      End
      Begin VB.Menu mnuLogOff 
         Caption         =   "LogOff"
      End
      Begin VB.Menu mnus1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewAccount 
         Caption         =   "Create Account"
      End
      Begin VB.Menu mnuRenameAccount 
         Caption         =   "Rename Account"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Account"
      End
      Begin VB.Menu mnus2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuLastCode 
         Caption         =   "View Lastest Code"
      End
      Begin VB.Menu mnuCodeSearch 
         Caption         =   "Code Search"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuBWhiteBoard 
         Caption         =   "Block All WhiteBoard Requests"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuNotiMini 
         Caption         =   "Notify When Message Recived (Sound)"
      End
      Begin VB.Menu mnusp7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewIgn 
         Caption         =   "View Ignored List"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
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

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long
Private Const WM_PASTE = &H302

Public MyName As String
Public MyInfo As String 'Just Used to save Your Profile/Name/Pass Once u quit
Public Sent As Long
Public size As Long
Dim frmPM(25) As New frmPrivateM
Dim Ignored As String
Dim Similies As String
Public BJustJoined As Integer
Dim strCount As String
Dim intStatAt As Integer
Dim blnChanged As Boolean




Private Sub cmbStatus_Click()
'Sends Status Change to Server
   If Winsock1.State = 7 Then
      Winsock1.SendData "StatusChange/Ñ/" & cmbStatus.SelectedItem.Index
   End If
   'MsgBox cmbStatus.SelectedItem.Index
End Sub

Private Sub cmdSend_Click()
      
   'Looks for Extra Commands
   If RTBSend.Text = "/OPS" Then
      If Winsock1.State = 7 Then
        Winsock1.SendData "/OPS"
        RTBSend.Text = ""
        Exit Sub
      End If
   End If
   
   If RTBSend.Text = "/Stats" Then
      If Winsock1.State = 7 Then
        Winsock1.SendData "/Stats"
        RTBSend.Text = ""
        Exit Sub
      End If
   End If
      
       
   'Chekcs to see if anything is in RTBSend
   If RTBSend.Text <> "" Then
     
     'Checks for TextLength
     'NOTE: Server Checks also ;)...Just as a little warning if you attempt
     'to remove this and u send over 100 chr, the server will BAN your IP.
     If Len(RTBSend.Text) >= 200 Or Len(txtrtf) > 4096 Then
       MsgBox "Msg is to long to send.", vbExclamation
       Exit Sub
     End If
     
     Dim strSend As String
     
     'WARNING: IF YOU REMOVE THIS YOU WILL
     'BE BANNED FROM THE SERVER!
     strSend = Replace(RTBSend.TextRTF, vbNewLine, "|ÿ|?|Ö|?|ÿ|")
     
         
     'Checks to see if Winsock is connected
     If Winsock1.State = 7 Then
         Winsock1.SendData "SendToAll/Ñ/" & strSend
     End If
     
     'Clearing the RTBSend
     RTBSend.Text = ""
   End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
MsgBox Asc(Right(RTBRecive.Text, 1))
End Sub

Private Sub Form_Load()
   On Error Resume Next
   Me.Caption = "IntraDream PSC Chat 4.0 BUILD: " & App.Major & "." & App.Minor & "." & App.Revision
   Open App.Path & "\Settings.ini" For Input As #122
   
   Dim strInBuffer As String
   
   Input #122, strInBuffer
   MyInfo = strInBuffer
      
   Input #122, strInBuffer
   Ignored = strInBuffer
   
   Close #122
   
   Open App.Path & "\BackupServers.ini" For Input As #122
   Do Until EOF(122) = True
     If EOF(122) = True Then Exit Do
     
     DoEvents
     Input #122, strInBuffer
     cmbAddress.AddItem strInBuffer
   Loop
   
   Close #122
   'Adds The MyStatus Settings
   cmbStatus.ComboItems.Add , , "Available", 3
   cmbStatus.ComboItems.Add , , "Away", 5
   cmbStatus.ComboItems.Add , , "On The Phone", 6
   cmbStatus.ComboItems.Add , , "Eating", 7
   cmbStatus.ComboItems.Add , , "Be Right Back", 8
   cmbStatus.ComboItems(1).Selected = True
   
   With m_IconData
        .cbSize = Len(m_IconData)
        .hWnd = Me.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "PSC Chat 4" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
         End With
   Shell_NotifyIcon NIM_ADD, m_IconData
   
   Similies = ":),:(,:o,;@,:l,:<,:t,|),:z,:k"
   
   'Loads Settings from Registry
   mnuPopupBallon.Checked = GetSetting("PSC Chat 4", "Settings", "mnuPopupBallon", False)
   mnuBWhiteBoard.Checked = GetSetting("PSC Chat 4", "Settings", "mnuBWhiteBoard", False)
   mnuNotiMini.Checked = GetSetting("PSC Chat 4", "Settings", "mnuNotiMini", True)
   mnuNotifyReg.Checked = GetSetting("PSC Chat 4", "Settings", "NotifyReg", True)
   RTBSend.Font = GetSetting("PSC Chat 4", "Settings", "Font", "Arial")
   RTBSend.Font.Bold = GetSetting("PSC Chat 4", "Settings", "FontBold", False)
   RTBSend.Font.Italic = GetSetting("PSC Chat 4", "Settings", "FontItalic", False)
   RTBSend.Font.size = GetSetting("PSC Chat 4", "Settings", "FontSize", 8)
   RTBSend.SelColor = GetSetting("PSC Chat 4", "Settings", "FontColor", vbBlack)
   
   
   SaveSetting "PSC Chat 4", "Settings", "mnuNotiMini", mnuNotiMini.Checked
   SaveSetting "PSC Chat 4", "Settings", "mnuBWhiteBoard", mnuBWhiteBoard.Checked
   SaveSetting "PSC Chat 4", "Settings", "mnuPopupBallon", mnuPopupBallon.Checked
   strCount = "0"
   
   Dim x As Long
   x = InitCommonControls
   
   blnChanged = False
   
   'Initalizes the URL Link
   'Dim lngEventMask   As Long
   'Dim lngWin32apiResultCode As Long
   '
   'With RTBRecive
   ' lngEventMask = SendMessage(.hWnd, EM_GETEVENTMASK, 0, ByVal CLng(0))
   '
   ' If lngEventMask Xor ENM_LINK Then
   '     lngEventMask = lngEventMask Or ENM_LINK
   ' End If
   '
   ' lngWin32apiResultCode = SendMessage(.hWnd, EM_SETEVENTMASK, 0, ByVal CLng(lngEventMask))
   ' lngWin32apiResultCode = SendMessage(.hWnd, EM_AUTOURLDETECT, CLng(1), ByVal CLng(0))
   'End With
   '
   'glngOriginalhWnd = Me.hWnd
   'glnglpOriginalWndProc = SetWindowLong(glngOriginalhWnd, GWL_WNDPROC, AddressOf RichTextBoxSubProc)
   
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  Dim msgCallBackMessage As Long
  msgCallBackMessage = x / Screen.TwipsPerPixelX
  
  Select Case msgCallBackMessage
    Case WM_LBUTTONDOWN
       WindowState = vbNormal
       Me.Show
    Case WM_RBUTTONDOWN
  End Select
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  If Me.WindowState = 1 Then
    Me.Hide
  End If
  
  RTBRecive.Width = Me.ScaleWidth - 2820
  cmbAddress.Width = RTBRecive.Width
  TBar.Width = RTBRecive.Width
  RTBSend.Width = Me.ScaleWidth - 1020
  cmbStatus.Left = cmbAddress.Width + 225
  lvUsers.Left = cmbAddress.Width + 205
  cmdSend.Left = RTBSend.Width + 230
  
  RTBRecive.Height = Me.ScaleHeight - 2250
  TBar.Top = (RTBRecive.Top + RTBRecive.Height) + 60
  lvUsers.Height = (TBar.Top + TBar.Height) - lvUsers.Top
  RTBSend.Top = (TBar.Top + TBar.Height) + 60
  cmdSend.Top = RTBSend.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SaveSetting "PSC Chat 4", "Settings", "NotifyReg", mnuNotifyReg.Checked
   SaveSetting "PSC Chat 4", "Settings", "mnuNotiMini", mnuNotiMini.Checked
   SaveSetting "PSC Chat 4", "Settings", "mnuBWhiteBoard", mnuBWhiteBoard.Checked
   'SaveSetting "PSC Chat 4", "Settings", "mnuPopupBallon", mnuPopupBallon.Checked
  
   Open App.Path & "\Settings.ini" For Output As #122
   Write #122, MyInfo
   Write #122, Ignored
   Close #122
   
   Open App.Path & "\BackupServers.ini" For Output As #122
   Dim lngA As Long
   For lngA = 1 To cmbAddress.ListCount - 1
     Write #122, cmbAddress.List(lngA)
   Next lngA
   Close #122
   StopCam
   Shell_NotifyIcon NIM_DELETE, m_IconData
   
   Dim lngWin32apiResultCode As Long
   lngWin32apiResultCode = SetWindowLong(glngOriginalhWnd, GWL_WNDPROC, glnglpOriginalWndProc)
   End
   
End Sub


Private Sub lvUsers_DblClick()
 ' On Error Resume Next
 
  'Checks for avible form
  Dim lngA As Long
  For lngA = 1 To 24
     DoEvents
     Me.Caption = lngA
     If frmPM(lngA).Visible = False Then
       If InStr(1, frmPM(lngA).Caption, lvUsers.SelectedItem.Text) Then
          frmPM(lngA).Show
          Exit Sub
       End If
     End If
  Next lngA
  
  For lngA = 1 To 24
     DoEvents
     If frmPM(lngA).Visible = False Then
        frmPM(lngA).Caption = "< " & lvUsers.SelectedItem.Text & " > Private Message"
        frmPM(lngA).Show
        Exit Sub
     End If
  Next lngA
     
   
End Sub

Private Sub lvUsers_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then
      PopupMenu mnuRClick
   End If
End Sub

Private Sub mnuAbout_Click()
  frmAbout.Show
End Sub

Private Sub mnuBan60_Click()
   On Error Resume Next
   
   Dim strResone As String
   strResone = InputBox("Please Enter a Resone.")
   
   'If REMOVED Server will not ban!
   If strResone = "" Then Exit Sub
   
   'Bans User for 60 Minutes
   If Winsock1.State = 7 Then
      Winsock1.SendData "Ban60/Ñ/" & lvUsers.SelectedItem.Text & "/Ñ/" & strResone
   End If
End Sub

Private Sub mnuBWhiteBoard_Click()
   If mnuBWhiteBoard.Checked = True Then
     mnuBWhiteBoard.Checked = False
     Exit Sub
   End If
   
   If mnuBWhiteBoard.Checked = False Then mnuBWhiteBoard.Checked = True
   
End Sub

Private Sub mnuClear60Min_Click()
   If Winsock1.State = 7 Then
      Winsock1.SendData "Clear60minBan/Ñ/"
   End If
End Sub

Private Sub mnuDelete_Click()
   On Error Resume Next
      
   If Winsock1.State = 7 Then
      'Prompts User for UserName/Pass
      Dim strRemove As String
      strRemove = InputBox("Please Enter Account Name, that you want removed.", "PSC Remove Account")
      strRemove = strRemove & "/Ö/" & InputBox("Please Enter The Password to the Account.", "PSC Remove Account")
   
     'Sends Removal Request
     If Winsock1.State = 7 Then
        Winsock1.SendData "RemoveAcc/Ñ/" & strRemove
     End If
   Else
     MsgBox "You Most Be Connected to Server!", vbExclamation, "PSC Chat 4.0"
  End If
End Sub



Private Sub mnuIgnore_Click()
   On Error Resume Next
   'Checks if users name is already blocked
   Dim var As Variant
   var = Split(Ignored, "ÖÖ")
   
   Dim lngA As Long
   For lngA = 0 To UBound(var)
     DoEvents
     If var(lngA) = lvUsers.SelectedItem.Text Then
       Exit Sub
     End If
   Next lngA
     
   'Changes the User Icon to Ignored Icon
   lvUsers.SelectedItem.SmallIcon = 4
   
   'Adds User to Ignored List
   Ignored = Ignored & "ÖÖ" & lvUsers.SelectedItem.Text
   
   
End Sub

Private Sub mnuKick_Click()
   On Error Resume Next
   'Kicks User
   If Winsock1.State = 7 Then
      Winsock1.SendData "Kick/Ñ/" & lvUsers.SelectedItem.Text
   End If
  
   
End Sub

Private Sub mnuLastCode_Click()
   frmLastestCode.Show
End Sub

Private Sub mnuLogin_Click()
   On Error Resume Next
   
   'Closes Connection
   Winsock1.Close
   
   RTBRecive.Text = "Disconnected..."
   lvUsers.ListItems.Clear
   cmdSend.Enabled = False
   cmbStatus.Enabled = False
   
   StaBar.Panels(1).Text = "Disconnected..."
   StaBar.Panels(2).Text = "Topic: "
   
   'Creates a Connection
   Winsock1.Connect cmbAddress.Text, "15009"
   
   'Updates StatusBar
   StaBar.Panels(1).Text = "Connecting..."
   
   'Gets Settings 'Profile etc
   Dim varsplit As Variant
   varsplit = Split(frmMain.MyInfo, "/Ö/")
   
   frmLogin.txtID = varsplit(0)
   frmLogin.txtPass = varsplit(1)
   frmLogin.txtProfile(1) = varsplit(2)
   frmLogin.txtProfile(2) = varsplit(3)
   frmLogin.txtProfile(3) = varsplit(4)
   frmLogin.txtProfile(4) = varsplit(5)
   frmLogin.txtProfile(5) = varsplit(6)
   
   'Enables EveryThing
   cmdSend.Enabled = True
   cmbStatus.Enabled = True
   
   'Shows the Login Form
   frmLogin.Show
   
End Sub

Private Sub mnuLogOff_Click()
   On Error Resume Next
   Winsock1.Close
   
   RTBRecive.Text = "Disconnected..." & vbNewLine
   RTBRecive.SelStart = Len(RTBRecive.Text)
   
   lvUsers.ListItems.Clear
   cmdSend.Enabled = False
   cmbStatus.Enabled = False
   
   StaBar.Panels(1).Text = "Disconnected..."
   StaBar.Panels(2).Text = "Topic: "
   
End Sub

Private Sub mnuMuzzle_Click()
   On Error Resume Next
   
   Dim lngTime As Long
   lngTime = InputBox("Please Enter how long you would like to muzzle " & lvUsers.SelectedItem.Text & ". (1 - 20 Minutes Max)", "How Long?")
      
   If lngTime <> "" Then
     'Muzzles A User
     If Winsock1.State = 7 Then
        Winsock1.SendData "Muzzle/Ñ/" & lvUsers.SelectedItem.Text & "/Ñ/" & lngTime
     End If
   End If
   
End Sub

Private Sub mnuNewAccount_Click()
   On Error Resume Next
   
   If Winsock1.State <> 7 Then
      'Closes Connection
      Winsock1.Close
      'Creates a Connection
      Winsock1.Connect cmbAddress.Text, "15009"
   End If
   
   'Updates StatusBar
   StaBar.Panels(1).Text = "Connecting..."
         
   'Shows the Login Form
   frmNewAccount.Show
End Sub

Private Sub mnuNotifyReg_Click()
    On Error Resume Next
   'Tells Server To Notify You when a New User Registrates
   If Winsock1.State = 7 Then
      If mnuNotifyReg.Checked = False Then
          mnuNotifyReg.Checked = True
          Winsock1.SendData "NotifyRegistrations/Ñ/"
          Exit Sub
        Else
          mnuNotifyReg.Checked = False
          Winsock1.SendData "DeNotifyRegistrations/Ñ/"
          Exit Sub
      End If
     Else
      MsgBox "Must Be Connected to Server", vbOKOnly + vbExclamation, "PSC Chat 4"
   End If
   
End Sub

Private Sub mnuNotiMini_Click()
   If mnuNotiMini.Checked = True Then
     mnuNotiMini.Checked = False
     Exit Sub
   End If
   
   If mnuNotiMini.Checked = False Then mnuNotiMini.Checked = True
End Sub

Private Sub mnuPBans_Click()
   On Error Resume Next
   
   Dim strResone As String
   strResone = InputBox("Please Enter a Resone.")
   
   'If REMOVED Server will not ban!
   If strResone = "" Then Exit Sub
   
   'Bans User for a Perment Ban
   If Winsock1.State = 7 Then
      Winsock1.SendData "BanP/Ñ/" & lvUsers.SelectedItem.Text & "/Ñ/" & strResone
   End If
End Sub



Private Sub mnuProfile_Click()
    On Error Resume Next
   'Sends Server Profile Request
   If Winsock1.State = 7 Then
      Winsock1.SendData "Profile/Ñ/" & lvUsers.SelectedItem.Text
   End If
End Sub

Private Sub mnuStats_Click()
   On Error Resume Next
   'Sends Status Request to Server
   If Winsock1.State = 7 Then
      Winsock1.SendData "StatusRequest/Ñ/" & lvUsers.SelectedItem.Text
      frmUserStat.lblName = lvUsers.SelectedItem.Text & " Status..."
      frmUserStat.Show
   End If
End Sub

Private Sub mnuUnIgnore_Click()
  On Error Resume Next
   'Removes Blocked User
  Dim varsplit As Variant
  varsplit = Split(Ignored, "ÖÖ")
  
  Dim lngA As Long
  For lngA = 1 To UBound(varsplit)
    DoEvents
    If lvUsers.SelectedItem.Text = varsplit(lngA) Then
       varsplit(lngA) = ""
       GoTo OK
    End If
  Next lngA
       
OK:
   Ignored = ""
   
   For lngA = 1 To UBound(varsplit)
      DoEvents
      If varsplit(lngA) <> "" Then
        Ignored = Ignored & "ÖÖ" & varsplit(lngA)
      End If
   Next lngA
   
   'Now We Change the Users Icon
   lvUsers.SelectedItem.SmallIcon = 3
End Sub

Private Sub mnuViewIgn_Click()

  Dim varsplit As Variant
  varsplit = Split(Ignored, "ÖÖ")
  
  Dim lngA As Long
  For lngA = 1 To UBound(varsplit)
    MsgBox varsplit(lngA)
  Next lngA
  

  
End Sub

Private Sub mnuWhiteBoard_Click()
    On Error Resume Next
   'Loads Whiteboard Form
    If frmWhiteBoard.Visible = False Then
      
      'Asks for a Port to use
      Dim intPort As Integer
      intPort = InputBox("Please Enter a Open Port")
               
      frmWhiteBoard.SckServer.LocalPort = intPort
      frmWhiteBoard.SckServer.Listen
      
      If Winsock1.State = 7 Then
        Winsock1.SendData "WhiteBoard/Ñ/" & lvUsers.SelectedItem.Text & "/Ñ/" & intPort
      End If
      
      frmWhiteBoard.Caption = "< " & lvUsers.SelectedItem.Text & " > WhiteBoard"
      frmWhiteBoard.Show
      
      frmWhiteBoard.intType = 1
      Exit Sub
     Else
      MsgBox "WhiteBoard is in use", vbOKCancel + vbCritical
      
    End If
  
End Sub

Private Sub RTBRecive_Change()
  On Error Resume Next
  RTBRecive.SelStart = Len(RTBRecive.Text)
End Sub

Private Sub RTBRecive_KeyPress(KeyAscii As Integer)
  If KeyAscii = 3 Then Exit Sub
  RTBSend.SetFocus
  RTBSend.Text = RTBSend.Text & Chr(KeyAscii)
  RTBSend.SelStart = Len(RTBSend.Text)
  KeyAscii = 0
  
  RTBRecive.SelStart = Len(RTBRecive.Text)
End Sub

Private Sub RTBSend_KeyPress(KeyAscii As Integer)
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
             RTBSend.SelFontName = cd1.FontName
             SaveSetting "PSC Chat 4", "Settings", "Font", cd1.FontName
          End If
          
          RTBSend.SelBold = cd1.FontBold
          RTBSend.SelItalic = cd1.FontItalic
          RTBSend.SelFontSize = cd1.FontSize
          
          SaveSetting "PSC Chat 4", "Settings", "FontBold", cd1.FontBold
          SaveSetting "PSC Chat 4", "Settings", "FontItalic", cd1.FontItalic
          SaveSetting "PSC Chat 4", "Settings", "FontSize", cd1.FontSize
          
       'Shows the Size Dialog
       Case "Size"
          
       'Shows the Color Dialog
       Case "color"
          cd1.ShowColor
          
          'Adds the Colour to the RTB
          If cd1.Color <> "" Then
             RTBSend.SelColor = cd1.Color
             SaveSetting "PSC Chat 4", "Settings", "FontColor", cd1.Color
          End If
       Case "faces"
          frmFaces.Show
          
   End Select
   
End Sub

Private Sub TBar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
      'Next are for the Size
   Select Case ButtonMenu.Key
       Case "s8"
         RTBSend.SelFontSize = 8
         SaveSetting "PSC Chat 4", "Settings", "FontSize", "8"
       Case "s9"
         RTBSend.SelFontSize = 9
         SaveSetting "PSC Chat 4", "Settings", "FontSize", "9"
       Case "s10"
         RTBSend.SelFontSize = 10
         SaveSetting "PSC Chat 4", "Settings", "FontSize", "10"
       Case "s11"
         RTBSend.SelFontSize = 11
         SaveSetting "PSC Chat 4", "Settings", "FontSize", "11"
       Case "s12"
         RTBSend.SelFontSize = 12
         SaveSetting "PSC Chat 4", "Settings", "FontSize", "12"
       Case "s13"
         RTBSend.SelFontSize = 13
         SaveSetting "PSC Chat 4", "Settings", "FontSize", "13"
       Case "s14"
         RTBSend.SelFontSize = 14
         SaveSetting "PSC Chat 4", "Settings", "FontSize", "14"
   End Select
End Sub

Private Sub tmrIdle_Timer()
   If SystemIdleCheck = True Then
     strCount = (CLng(strCount) + 1)
     
     If CLng(strCount) = 1000 Then
     
       If cmbStatus.SelectedItem.Index = 1 Then
          cmbStatus.ComboItems.Item(2).Selected = True
          blnChanged = True
          If Winsock1.State = 7 Then
            Winsock1.SendData "StatusChange/Ñ/" & cmbStatus.SelectedItem.Index
          End If
       End If
      End If
    Else
       If blnChanged = True Then
         blnChanged = False
         strCount = 0
         cmbStatus.ComboItems.Item(1).Selected = True
         If Winsock1.State = 7 Then
            Winsock1.SendData "StatusChange/Ñ/" & cmbStatus.SelectedItem.Index
         End If
       
       
      End If
   End If

   
End Sub

Private Sub Winsock1_Close()
   RTBRecive.Text = "Disconnected..."
   lvUsers.ListItems.Clear
   cmdSend.Enabled = False
   cmbStatus.Enabled = False
   
   StaBar.Panels(1).Text = "Disconnected..."
   StaBar.Panels(2).Text = "Topic: "
   
End Sub

Private Sub Winsock1_Connect()
   blnChanged = False
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
   Dim lngA As Long
   Dim lngB As Long
   Dim varsplit As Variant
   Dim VarVbEnter As Variant
   Dim strInBuffer As String
   
   Winsock1.GetData strInBuffer
   VarVbEnter = Split(strInBuffer, vbNewLine)
   
   For lngA = 0 To UBound(VarVbEnter) - 1
      
      varsplit = Split(VarVbEnter(lngA), "/Ñ/")
      RTBRecive.SelStart = Len(RTBRecive.Text)
      Select Case varsplit(0)
         'Updates Psc Chat that tells user that he/she has connected
         Case "Connected"
             StaBar.Panels(1).Text = "Connected"
             RTBRecive.SelBold = True
             RTBRecive.SelColor = vbBlack
             RTBRecive.SelText = "Connected..."
             
             RTBSend.Enabled = True
             cmdSend.Enabled = True
             
         'Updates the Topic Area
         Case "Topic"
             RTBRecive.SelBold = True
             RTBRecive.SelColor = vbBlue
             RTBRecive.SelText = vbNewLine & "SERVER >> Todays Topic is: " & varsplit(1) & vbNewLine
             StaBar.Panels(2).Text = "Topic: " & varsplit(1)
             
         'Adds the Backup IPs
         Case "BackupIPs"
            If blnChanged = False Then
               cmbAddress.Clear
               blnChanged = True
            End If
            
            cmbAddress.AddItem varsplit(1)
                     
         'Displays a MsgBox to User
         Case "Msgbox"
             MsgBox varsplit(1), vbExclamation, "PSC Chat 4.0"
             
         'Displays Server Message in Chatwindow
         Case "MsgRTB"
             RTBRecive.SelBold = True
             RTBRecive.SelColor = vbBlue
             RTBRecive.SelText = "SERVER >> " & varsplit(1) & vbNewLine
             RTBRecive.SelStart = Len(RTBRecive.Text)
             
             'Adds Faces
             Dim varIcons3 As Variant
             varIcons3 = Split(Similies, ",")
             
             For lngB = 0 To UBound(varIcons3)
NextIcon3:
                DoEvents
                If InStr(LCase(RTBRecive.SelText), varIcons3(lngB)) Then
                    Dim lngStart3 As Long
                    lngStart3 = InStr(LCase(RTBRecive.SelText), varIcons3(lngB))
                
                    RTBRecive.SelStart = lngStart3 - 1
                    RTBRecive.SelLength = 2
                
                    'Gets Data in clipboard
                    Dim OldClip3
                    OldClip3 = Clipboard.GetText
                
                    Clipboard.Clear
                    Clipboard.SetData faces.ListImages(lngB + 1).Picture
                               
                    RTBRecive.Locked = False
                    SendMessage RTBRecive.hWnd, WM_PASTE, 0, 0
                    RTBRecive.Locked = True
                
                    Clipboard.SetText OldClip3
                
                    RTBRecive.SelStart = Len(RTBRecive.SelText)
                    GoTo NextIcon3
                End If
             Next lngB
             'StaBar.Panels(2).Text = "Topic: " & varSplit(1)
             
         'Displays Warning Level
         Case "Warning"
             StaBar.Panels(3).Text = "Warning(s): " & varsplit(1)
             
         'Adds Recived Text into RTBRecived
         Case "ChatRTBText"
             'Checks if you ignored the user
             Dim varban As Variant
             varban = Split(Ignored, "ÖÖ")
             
             For lngB = 1 To UBound(varban)
                DoEvents
                If varban(lngB) = varsplit(1) Then
                   Exit Sub
                End If
             Next lngB
             
             Dim strFinalOutput As String
             strFinalOutput = Replace(varsplit(2), "|ÿ|?|Ö|?|ÿ|", vbNewLine)
             
             'Adds Name Where it came from and
             'Adds the Message to the RTBRecive
             'RTBRecive.SelText = RTBRecive.SelText & vbCrLf
             RTBRecive.SelColor = vbBlue
             RTBRecive.SelFontSize = 8
             RTBRecive.SelItalic = False
             RTBRecive.SelBold = False
             RTBRecive.SelFontName = "Arial"
             RTBRecive.SelText = "<" & varsplit(1) & "> "
             
             RTBRecive.SelRTF = strFinalOutput
             RTBRecive.SelStart = Len(RTBRecive.Text)
             
             StaBar.Panels(4).Text = "Last Message: " & Time
             
             'Adds Faces
             Dim varIcons As Variant
             varIcons = Split(Similies, ",")
             
             For lngB = 0 To UBound(varIcons)
NextIcon:
                DoEvents
                If InStr(LCase(RTBRecive.Text), varIcons(lngB)) Then
                    Dim lngStart As Long
                    lngStart = InStr(LCase(RTBRecive.Text), varIcons(lngB))
                
                    RTBRecive.SelStart = lngStart - 1
                    RTBRecive.SelLength = 2
                
                    'Gets Data in clipboard
                    Dim OldClip
                    OldClip = Clipboard.GetText
                
                    Clipboard.Clear
                    Clipboard.SetData faces.ListImages(lngB + 1).Picture
                               
                    RTBRecive.Locked = False
                    SendMessage RTBRecive.hWnd, WM_PASTE, 0, 0
                    RTBRecive.Locked = True
                
                    Clipboard.SetText OldClip
                
                    RTBRecive.SelStart = Len(RTBRecive.SelText)
                    GoTo NextIcon
                End If
             Next lngB
             
             'Checks for Vbcrlf
             Dim strVBcrlf As String
againo:
             strVBcrlf = Right(RTBRecive.Text, 1)
             If strVBcrlf = Chr(10) Then
                Else
                RTBRecive.SelStart = Len(RTBRecive.Text)
                RTBRecive.SelText = RTBRecive.SelText & vbCrLf
             End If
             
             
             'Plays the Notify Sound when enabled
             If mnuNotiMini.Checked = True Then
                sndPlaySound App.Path & "\effect.wav", 1
             End If
             
         'Tells User that he isnt able to
         'Send Data b/c of Flood Manager
         Case "Err10"
             RTBRecive.SelBold = False
             RTBRecive.SelColor = vbRed
             RTBRecive.SelStart = Len(RTBRecive.Text)
             RTBRecive.SelText = "ANTI-FLOOD SYSTEM >> Your Message Wasnt Sent! Please Do not Continue to Type unless a Server Message Appears; Other wise you will be Banned. " & vbNewLine
             
         'Tells User he/she is Allowed to Talk Agin
         Case "Err11"
             RTBRecive.SelBold = False
             RTBRecive.SelColor = &H8000&
             RTBRecive.SelStart = Len(RTBRecive.Text)
             RTBRecive.SelText = "ANTI-FLOOD SYSTEM >> You May Now Resume Talking" & vbNewLine
             
         'Adds All Online Users to LvUsers
         Case "OnlineList"
             Dim varban1 As Variant
             varban1 = Split(Ignored, "ÖÖ")
             
             'Checks if you Ignored the Person
             Dim lngW As Long
             For lngW = 1 To UBound(varban1)
                DoEvents
                If varban1(lngW) = varsplit(1) Then
                    lvUsers.ListItems.Add , , varsplit(1), , 4
                    
                    Exit Sub
                End If
             Next lngW
             
             
             'Changes User(s) Icons
             Select Case varsplit(2)
                      Case 1
                         'Normal
                         lvUsers.ListItems.Add , , varsplit(1), , 3
                      Case 2
                         'Away
                         lvUsers.ListItems.Add , , varsplit(1), , 5
                      Case 3
                         'Phone
                         lvUsers.ListItems.Add , , varsplit(1), , 6
                      Case 4
                         'Eating
                         lvUsers.ListItems.Add , , varsplit(1), , 7
                      Case 5
                         'Brb
                         lvUsers.ListItems.Add , , varsplit(1), , 8
              End Select
              
         '-------------------------------
         'Adds User Name that just Joined
         Case "Join"
             Dim varban3 As Variant
             varban13 = Split(Ignored, "ÖÖ")
             
             'Checks if you Ignored the Person
             'Dim lngW3 As Long
             For lngW3 = 1 To UBound(varban13)
                DoEvents
                If varban13(lngW3) = varsplit(1) Then
                    lvUsers.ListItems.Add , , varsplit(1), , 4
                    Exit Sub
                End If
             Next lngW3
             
             RTBRecive.SelStart = Len(RTBRecive.Text)
             RTBRecive.SelColor = &H8000&
             RTBRecive.SelBold = True
             RTBRecive.SelText = " * " & varsplit(1) & " Has Joined..." & vbNewLine
             
             
             'Changes User(s) Icons
             Select Case varsplit(2)
                      Case 1
                         'Normal
                         lvUsers.ListItems.Add , , varsplit(1), , 3
                      Case 2
                         'Away
                         lvUsers.ListItems.Add , , varsplit(1), , 5
                      Case 3
                         'Phone
                         lvUsers.ListItems.Add , , varsplit(1), , 6
                      Case 4
                         'Eating
                         lvUsers.ListItems.Add , , varsplit(1), , 7
                      Case 5
                         'Brb
                         lvUsers.ListItems.Add , , varsplit(1), , 8
              End Select
                           
              
         'Removes Name from lvUsers
         Case "Left"
             For lngB = 1 To lvUsers.ListItems.Count
                DoEvents
                If lvUsers.ListItems(lngB).Text = varsplit(1) Then
                   lvUsers.ListItems.Remove lngB
                   Exit Sub
                End If
             Next lngB
             
         'This Changes User(s) Status Icon
         Case "StatusChange"
             
             'First we Check if you blocked the User
             Dim varblocked As Variant
             varblocked = Split(Ignored, "ÖÖ")
  
             Dim lngY As Long
             For lngY = 1 To UBound(varblocked)
                If varblocked(lngY) = varsplit(1) Then GoTo blocked
             Next lngY
             
             'Changes Icon if he/shes not blocked
             For lngB = 1 To lvUsers.ListItems.Count
                DoEvents
                If lvUsers.ListItems(lngB).Text = varsplit(1) Then
                   
                   'Here the Icons are Change To Status
                   Select Case varsplit(2)
                      Case 1
                         lvUsers.ListItems(lngB).SmallIcon = 3
                      Case 2
                         lvUsers.ListItems(lngB).SmallIcon = 5
                      Case 3
                         lvUsers.ListItems(lngB).SmallIcon = 6
                      Case 4
                         lvUsers.ListItems(lngB).SmallIcon = 7
                      Case 5
                         lvUsers.ListItems(lngB).SmallIcon = 8
                   End Select
blocked:
                   Exit Sub
                End If
             Next lngB
        Case "SendFile1"
           On Error Resume Next
           
           frmMain.Hide
           frmProgress.Show
           
           frmProgress.lbOP = "Operation:  Saving Images"
           
           'Removes Exsiting PIC and Resaves
           Kill "C:\tmp335548.bmp"
           Kill "C:\tmp335549.jpg"
                                
           SavePicture Me.tmppic, "C:\tmp335548.bmp"
                     
           'Resaves as jpg
           BMPFileToJpg "C:\tmp335548.bmp", "C:\tmp335549.jpg", 40
                      
           'Checks File Size
           If FileLen("C:\tmp335549.jpg") <= "500" Then
               MsgBox "File To Small"
               
           End If
           
           'Opens the File for Transfer
           Open "C:\tmp335549.jpg" For Binary As #22
           
           
           'Dims the File Buffer
           Dim fileB As String
           fileB = Space(1024)
           
           'Puts Data into the FileB
           Get 22, , fileB
           
           'Checks if EOF, So we can send File Ender
           If EOF(22) = True Then
              fileB = fileB & "\\F"
              frmProgress.cmdOK.Enabled = True
              'Closes the File
              Close #22
           End If
           
           Dim FileLength As String
           FileLength = FileLen("C:\tmp335549.jpg")
           
           frmProgress.lbOP = "Operation:  Sending Image"
           frmProgress.lblStatus(0).Caption = "File Size: " & FileLength
           frmProgress.lblStatus(1).Caption = "Total Sent: " & Len(fileB)
           frmProgress.lblStatus(2).Caption = "Percentage Complete: " & Format(Len(fileB) / FileLength, "00.00%")
           
           frmProgress.prg1.Max = FileLength
           frmProgress.prg1.Value = Len(fileB)
           
           size = FileLength
           Sent = Len(fileB)
           
           'Sends Data
           Winsock1.SendData FileLength & "/||ÖÖ||\" & fileB
                      
        Case "NextFile"
           
           'Dims Variable for File Transfer
           Dim filebb As String
           filebb = Space(1024)
           
           'Puts File into Mem
           Get 22, , filebb
           
           'Checks if EOF is true
           If EOF(22) = True Then
              filebb = filebb & "\\F"
              frmProgress.cmdOK.Enabled = True
              
              'Closes Open File if EOF
              Close #22
           End If
           
           Sent = Sent + Len(filebb)
           frmProgress.lblStatus(1).Caption = "Total Sent: " & Sent
           frmProgress.lblStatus(2).Caption = "Percentage Complete: " & Format(Sent / size, "00.00%")
           
           If Sent > size Then
             frmProgress.prg1.Value = frmProgress.prg1.Max
             GoTo ToBig
           End If
           
           frmProgress.prg1.Value = Sent
ToBig:
           
           'Sends Data
           Winsock1.SendData filebb
                      
       'When a Private Message is Recived
       Case "PM"
       
           'Checks if a Chatwindow is already open
           For lngB = 1 To 50
              If InStr(1, frmPM(lngB).Caption, varsplit(1)) Then
                   Dim strFinalOutput2 As String
                   strFinalOutput2 = Replace(varsplit(2), "|ÿ|?|Ö|?|ÿ|", vbNewLine)
             
                   'Adds Name Where it came from and
                   'Adds the Message to the RTBRecive
                   frmPM(lngB).rtbR.SelColor = vbBlue
                   frmPM(lngB).rtbR.SelText = "<" & varsplit(1) & "> "
             
                   frmPM(lngB).rtbR.SelRTF = strFinalOutput2
                   frmPM(lngB).rtbR.SelStart = Len(frmPM(lngB).rtbR.Text)
                   
                   frmPM(lngB).Show
                   
                   'Adds Faces to PM
                   Dim lngBB2 As Long
                   Dim varIcons2 As Variant
                   varIcons2 = Split(Similies, ",")
             
                  For lngBB2 = 0 To UBound(varIcons2)
NextIcon2:
                      DoEvents
                      If InStr(LCase(frmPM(lngB).rtbR.Text), varIcons2(lngBB2)) Then
                         Dim lngStart2 As Long
                         lngStart2 = InStr(LCase(frmPM(lngB).rtbR.Text), varIcons2(lngBB2))
                
                         frmPM(lngB).rtbR.SelStart = lngStart2 - 1
                         frmPM(lngB).rtbR.SelLength = 2
                
                         'Gets Data in clipboard
                         Dim OldClip2
                         OldClip2 = Clipboard.GetText
                
                         Clipboard.Clear
                         Clipboard.SetData faces.ListImages(lngBB2 + 1).Picture
                               
                         frmPM(lngB).rtbR.Locked = False
                         SendMessage frmPM(lngB).rtbR.hWnd, WM_PASTE, 0, 0
                         frmPM(lngB).rtbR.Locked = True
                
                         Clipboard.SetText OldClip2
                
                         frmPM(lngB).rtbR.SelStart = Len(frmPM(lngB).rtbR.Text)
                        GoTo NextIcon2
                      End If
                  Next lngBB2
                  
                  Dim strVBcrlf1 As String

                  strVBcrlf1 = Right(frmPM(lngB).rtbR.Text, 1)
                         
                  If strVBcrlf1 = Chr(10) Then
                     Else
                       frmPM(lngB).rtbR.SelStart = Len(frmPM(lngB).rtbR.Text)
                       frmPM(lngB).rtbR.SelText = frmPM(lngB).rtbR.SelText & vbCrLf
                  End If
                   
                   Dim nReturnValue2 As Long
                   nReturnValue2 = FlashWindow(frmPM(lngB).hWnd, False)
                   Exit Sub
               End If
           Next lngB
           
           'Creates New ChatWindow
           For lngB = 1 To 50
              If frmPM(lngB).Visible = False Then
                  frmPM(lngB).Caption = "< " & varsplit(1) & " > Private Message"
                  Dim strFinalOutput3 As String
                  strFinalOutput3 = Replace(varsplit(2), "|ÿ|?|Ö|?|ÿ|", vbNewLine)
             
                  'Adds Name Where it came from and
                  'Adds the Message to the RTBRecive
                  frmPM(lngB).rtbR.SelColor = vbBlue
                  frmPM(lngB).rtbR.SelText = "<" & varsplit(1) & "> "
             
                  frmPM(lngB).rtbR.SelRTF = strFinalOutput3
                  frmPM(lngB).rtbR.SelStart = Len(frmPM(lngB).rtbR.Text)
                   
                  frmPM(lngB).Show
                  Dim nReturnValue As Long
                  nReturnValue = FlashWindow(frmPM(lngB).hWnd, False)
                  frmPM(lngB).Show
                  Exit Sub
              End If
           Next lngB
           
        'Asks for Camera Permission
        Case "CamRequest"
          'Checks if a Chatwindow is already open
           For lngB = 1 To 50
              If InStr(1, frmPM(lngB).Caption, varsplit(1)) Then
                   frmPM(lngB).strIP = varsplit(2)
                   frmPM(lngB).strPort = varsplit(3)
                                                       
                   frmPM(lngB).picAccept.Visible = True
                   frmPM(lngB).Show
                   
                   Dim nReturnValue3 As Long
                   nReturnValue3 = FlashWindow(frmPM(lngB).hWnd, False)
                   Exit Sub
               End If
           Next lngB
           
           
        'WhiteBoard Request
        Case "WBRequest"
          If mnuBWhiteBoard.Checked = True Then
             'TODO: Sends Blocked
             Exit Sub
          End If
          
          Dim strYesNO As String
          strYesNO = MsgBox("You Have Recived a WhiteBoard Request from " & varsplit(1) & vbNewLine & vbNewLine & "Press Yes to Accept and No to Cancel.", vbYesNo)
          
          If strYesNO = vbYes Then
             frmWhiteBoard.Caption = "<" & varsplit(1) & "> WhiteBoard"
             
             'Connects Winsocks
             frmWhiteBoard.SckClient.Connect varsplit(2), varsplit(3)
             
             frmWhiteBoard.Show
          End If
          
        'Tells User that the Cam Request has been canceled
        Case "CamDeclined"
           For lngB = 1 To 50
              If InStr(1, frmPM(lngB).Caption, varsplit(1)) Then
                   'Tells user Video has been Canceled
                   frmPM(lngB).rtbR.Text = frmPM(lngB).rtbR.Text & vbNewLine & "Video Request has been Canceled" & vbNewLine
                   frmPM(lngB).Winsock1.Close
                   frmPM(lngB).Width = 5805
                   frmPM(lngB).Show
                   
                   Dim nReturnValue4 As Long
                   nReturnValue4 = FlashWindow(frmPM(lngB).hWnd, False)
                   Exit Sub
               End If
           Next lngB
           
        'Recivevs the User status
        Case "UserStatus"
          frmUserStat.Label1(0).Caption = varsplit(1)
          frmUserStat.Label1(1).Caption = varsplit(2)
          frmUserStat.Label1(3).Caption = varsplit(3)
          frmUserStat.Label1(2).Caption = varsplit(4)
          
        'Show Profile
        Case "Profile"
          On Error Resume Next
          Dim intA As Integer
          Dim varPhasing As Variant
          varPhasing = Split(varsplit(2), "/Ö/")
          
          For intA = 1 To UBound(varPhasing)
             DoEvents
             frmProfile.lblItem(intA).Caption = varPhasing(intA)
          Next intA
          
          frmProfile.Caption = "<" & varsplit(1) & "> Profile - PSC Chat 4"
          'frmProfile.WebBrowser1.Navigate "http://pscchat.intradream.com/Images/" & varsplit(1) & ".jpg"
          
          Open App.Path & "\Pic\" & varsplit(1) & ".jpg" For Binary As #200
          
          frmProfile.Show
          
          Dim data() As Byte
          data() = frmProfile.Inet1.OpenURL("http://pscchat.intradream.com/Images/" & varsplit(1) & ".jpg", 1)
          
          Put #200, , data()
          Close #200
          frmProfile.Picture1.Picture = LoadPicture(App.Path & "\Pic\" & varsplit(1) & ".jpg")
          
          frmProfile.Show
          On Error GoTo 0
          
        
          
          
      End Select
   Next lngA
   
End Sub




