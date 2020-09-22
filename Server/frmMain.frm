VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PSC Chat 4 Server"
   ClientHeight    =   6465
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   9915
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock winsock1 
      Index           =   0
      Left            =   120
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   15009
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   1005
      ButtonWidth     =   1244
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgLToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Status"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "  Users  "
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Picture"
            Key             =   "Picture"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Flooding"
            Key             =   "Flooding"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Advert"
            Key             =   "Advert"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            Key             =   "About"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Restart"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   11033
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmMain.frx":39FA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTopic"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblBackup"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblOutIP"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "rtbLog"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtTopic"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "imgLToolbar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdChangeProfile"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lstBackup"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdAddBckup"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdRemoveBckup"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtIP"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Users"
      TabPicture(1)   =   "frmMain.frx":3A16
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvAccounts"
      Tab(1).Control(1)=   "lvOnline"
      Tab(1).Control(2)=   "Frame5"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Pictures Service"
      TabPicture(2)   =   "frmMain.frx":3A32
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblFTPstatus"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lstftp"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "txtinput"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame3"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "lvImageQuery"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Inet1"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "tmrFTPup"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "chkGenHTML"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "chkUseHTML"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "chkEnable"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "AntiFlood/Ban"
      TabPicture(3)   =   "frmMain.frx":3A4E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdRemoveSelBan"
      Tab(3).Control(1)=   "cmdClear60minban"
      Tab(3).Control(2)=   "lvMuzzled"
      Tab(3).Control(3)=   "Frame1"
      Tab(3).Control(4)=   "Frame2"
      Tab(3).Control(5)=   "cmdClearMuzzle"
      Tab(3).Control(6)=   "timerMuzzled"
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "Advertising"
      TabPicture(4)   =   "frmMain.frx":3A6A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblleft"
      Tab(4).Control(1)=   "lblIval"
      Tab(4).Control(2)=   "lvAdvert"
      Tab(4).Control(3)=   "tmrAdvert"
      Tab(4).Control(4)=   "txtinterval"
      Tab(4).Control(5)=   "cmdAdDelete"
      Tab(4).Control(6)=   "cmdAdd"
      Tab(4).ControlCount=   7
      TabCaption(5)   =   "About"
      TabPicture(5)   =   "frmMain.frx":3A86
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmdVisitdotcom"
      Tab(5).Control(1)=   "pic"
      Tab(5).Control(2)=   "lblAbout(4)"
      Tab(5).Control(3)=   "lblAbout(3)"
      Tab(5).Control(4)=   "lblAbout(2)"
      Tab(5).Control(5)=   "lblAbout(1)"
      Tab(5).Control(6)=   "lblAbout(0)"
      Tab(5).ControlCount=   7
      Begin VB.CommandButton cmdVisitdotcom 
         Caption         =   "Visit IntraDream"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70200
         TabIndex        =   56
         Top             =   3120
         Width           =   4935
      End
      Begin VB.TextBox txtIP 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   49
         Text            =   "24.158.66.181"
         Top             =   480
         Width           =   2295
      End
      Begin VB.CheckBox chkEnable 
         Caption         =   "Enable Picture Service"
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
         Left            =   -68760
         TabIndex        =   48
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton cmdRemoveBckup 
         Caption         =   "Remove"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   47
         Top             =   5640
         Width           =   2055
      End
      Begin VB.CommandButton cmdAddBckup 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   5640
         Width           =   1815
      End
      Begin VB.ListBox lstBackup 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3000
         ItemData        =   "frmMain.frx":3AA2
         Left            =   120
         List            =   "frmMain.frx":3AA4
         TabIndex        =   45
         Top             =   2520
         Width           =   3855
      End
      Begin VB.Timer timerMuzzled 
         Interval        =   1000
         Left            =   -65520
         Top             =   2880
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70320
         TabIndex        =   39
         Top             =   5760
         Width           =   1455
      End
      Begin VB.CommandButton cmdAdDelete 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71880
         TabIndex        =   38
         Top             =   5760
         Width           =   1455
      End
      Begin VB.TextBox txtinterval 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74190
         TabIndex        =   37
         Text            =   "100"
         Top             =   5760
         Width           =   975
      End
      Begin VB.Timer tmrAdvert 
         Interval        =   1000
         Left            =   -65760
         Top             =   480
      End
      Begin VB.CommandButton cmdClearMuzzle 
         Caption         =   "Clear All Muzzles"
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
         Left            =   -67920
         TabIndex        =   36
         Top             =   5520
         Width           =   2775
      End
      Begin VB.Frame Frame5 
         Caption         =   "Operations"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   -74880
         TabIndex        =   31
         Top             =   360
         Width           =   2175
         Begin VB.CommandButton cmdOP 
            Caption         =   "OP Account"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   34
            Top             =   960
            Width           =   1695
         End
         Begin VB.CommandButton cmdDeOp 
            Caption         =   "DEOP Account"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   33
            Top             =   600
            Width           =   1695
         End
         Begin VB.CommandButton cmdDeleteAc 
            Caption         =   "Delete Account"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   32
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "      Quick Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   27
         Top             =   750
         Width           =   3855
         Begin VB.Timer Timer1 
            Interval        =   5000
            Left            =   3360
            Top             =   120
         End
         Begin VB.Label totalCon 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Connections: 0"
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
            Left            =   240
            TabIndex        =   30
            Top             =   1080
            Width           =   3375
         End
         Begin VB.Label lblTotalData 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Data Recived: 0"
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
            Left            =   240
            TabIndex        =   29
            Top             =   720
            Width           =   3375
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   120
            Picture         =   "frmMain.frx":3AA6
            Top             =   0
            Width           =   240
         End
         Begin VB.Label lblTotalUsers 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Logged in: 0"
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
            Left            =   240
            TabIndex        =   28
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.CommandButton cmdChangeProfile 
         Caption         =   "Change"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8760
         TabIndex        =   26
         Top             =   480
         Width           =   1080
      End
      Begin MSComctlLib.ImageList imgLToolbar 
         Left            =   9360
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3E30
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":41CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4564
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":48FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4C98
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5032
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":7984
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":7D1E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":80B8
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox chkUseHTML 
         Caption         =   "Use CoreHtml.txt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -68760
         TabIndex        =   22
         Top             =   480
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkGenHTML 
         Caption         =   "Generate Index.html"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -68760
         TabIndex        =   20
         Top             =   720
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.Timer tmrFTPup 
         Interval        =   10000
         Left            =   -65640
         Top             =   360
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   -66240
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         Protocol        =   2
         RemotePort      =   21
         URL             =   "ftp://"
         RequestTimeout  =   10
      End
      Begin MSComctlLib.ListView lvImageQuery 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   17
         Top             =   1920
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   7011
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "FileName"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Downloaded"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Size"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "index"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "HTML?"
            Object.Width           =   776
         EndProperty
      End
      Begin VB.Frame Frame3 
         Caption         =   "FTP Account"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -74880
         TabIndex        =   10
         Top             =   360
         Width           =   6015
         Begin VB.TextBox txtDir 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3600
            TabIndex        =   19
            Text            =   "htdocs\Images\"
            Top             =   240
            Width           =   2295
         End
         Begin VB.TextBox txtPass 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1200
            PasswordChar    =   "*"
            TabIndex        =   13
            Top             =   960
            Width           =   4695
         End
         Begin VB.TextBox txtUserName 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            TabIndex        =   12
            Text            =   "idream"
            Top             =   600
            Width           =   4695
         End
         Begin VB.TextBox txtLink 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            TabIndex        =   11
            Text            =   "ftp.pscchat.intradream.com"
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Password:"
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
            Left            =   120
            TabIndex        =   16
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "UserName:"
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
            Left            =   120
            TabIndex        =   15
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "FTP:"
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
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Anti-Flood System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   -74880
         TabIndex        =   8
         Top             =   2760
         Width           =   6855
         Begin VB.Timer tmrFloodClear 
            Interval        =   5
            Left            =   6360
            Top             =   240
         End
         Begin MSComctlLib.ListView lvFlood 
            Height          =   2655
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   4683
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Index"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "IP"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Time"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Text"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Notify"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Banning System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -74880
         TabIndex        =   6
         Top             =   360
         Width           =   9735
         Begin VB.Timer tmrBanning 
            Interval        =   1000
            Left            =   9240
            Top             =   240
         End
         Begin MSComctlLib.ListView lvBanned 
            Height          =   1935
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   3413
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Time"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "By?"
               Object.Width           =   2893
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "IP"
               Object.Width           =   3775
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Resone"
               Object.Width           =   3775
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Type"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "TimeLeft 60min Ban only"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.TextBox txtTopic 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4680
         TabIndex        =   4
         Text            =   "Visual Basic"
         Top             =   480
         Width           =   3975
      End
      Begin MSComctlLib.ListView lvOnline 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   3
         Top             =   3240
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   5106
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Index"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "IP"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Level"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Warning"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "OPNoti"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Mode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "total Chatwindow txt"
            Object.Width           =   2540
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtbLog 
         Height          =   5295
         Left            =   4080
         TabIndex        =   2
         Top             =   840
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   9340
         _Version        =   393217
         TextRTF         =   $"frmMain.frx":8452
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
      Begin MSComctlLib.ListView lvAccounts 
         Height          =   2655
         Left            =   -72600
         TabIndex        =   1
         Top             =   480
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   4683
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Login"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Pass"
            Object.Width           =   1305
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Level"
            Object.Width           =   1305
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "IP"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Last"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Profile"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtinput 
         Height          =   765
         Left            =   -73800
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   23
         Text            =   "frmMain.frx":84C9
         Top             =   3240
         Width           =   4215
      End
      Begin VB.ListBox lstftp 
         Height          =   2595
         Left            =   -71520
         TabIndex        =   21
         Top             =   2160
         Width           =   2535
      End
      Begin MSComctlLib.ListView lvMuzzled 
         Height          =   2055
         Left            =   -67920
         TabIndex        =   35
         Top             =   2880
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   3625
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Time"
            Object.Width           =   2187
         EndProperty
      End
      Begin MSComctlLib.ListView lvAdvert 
         Height          =   5175
         Left            =   -74880
         TabIndex        =   40
         Top             =   480
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   9128
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "#"
            Object.Width           =   776
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Company"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Advert"
            Object.Width           =   14887
         EndProperty
      End
      Begin VB.PictureBox pic 
         AutoSize        =   -1  'True
         Height          =   3105
         Left            =   -74880
         Picture         =   "frmMain.frx":84CF
         ScaleHeight     =   3045
         ScaleWidth      =   4320
         TabIndex        =   43
         Top             =   480
         Width           =   4380
      End
      Begin VB.CommandButton cmdClear60minban 
         Caption         =   "Clear 60 Min Ban(s)"
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
         Left            =   -67920
         TabIndex        =   57
         Top             =   5280
         Width           =   2775
      End
      Begin VB.CommandButton cmdRemoveSelBan 
         Caption         =   "Remove Selected Ban"
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
         Left            =   -67920
         TabIndex        =   58
         Top             =   5040
         Width           =   2775
      End
      Begin VB.Label lblAbout 
         Caption         =   $"frmMain.frx":33231
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   -70200
         TabIndex        =   55
         Top             =   2400
         Width           =   5055
      End
      Begin VB.Label lblAbout 
         Caption         =   $"frmMain.frx":332B8
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   -70200
         TabIndex        =   54
         Top             =   1680
         Width           =   5055
      End
      Begin VB.Label lblAbout 
         Caption         =   "Beta Testers: Unevenratio, Jeremy Gauthier, Luis Caicedo, Jonathan Lee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   -70200
         TabIndex        =   53
         Top             =   1200
         Width           =   5175
      End
      Begin VB.Label lblAbout 
         Caption         =   "Programmers: Carsten Dressler, Timothy Marin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   -70200
         TabIndex        =   52
         Top             =   840
         Width           =   4815
      End
      Begin VB.Label lblAbout 
         Caption         =   "IntraDream PSC Chat Server"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -70440
         TabIndex        =   51
         Top             =   480
         Width           =   4815
      End
      Begin VB.Label lblOutIP 
         Caption         =   "Outside IP: Click Me"
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
         Left            =   120
         TabIndex        =   50
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblBackup 
         Caption         =   "BackUp Server:"
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
         Left            =   120
         TabIndex        =   44
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label lblIval 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Interval:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74880
         TabIndex        =   42
         Top             =   5805
         Width           =   570
      End
      Begin VB.Label lblleft 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -73170
         TabIndex        =   41
         Top             =   5805
         Width           =   450
      End
      Begin VB.Label lblFTPstatus 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74880
         TabIndex        =   18
         Top             =   5955
         Width           =   9705
      End
      Begin VB.Label lblTopic 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Topic:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4080
         TabIndex        =   5
         Top             =   525
         Width           =   435
      End
   End
   Begin RichTextLib.RichTextBox rtbphase 
      Height          =   615
      Left            =   9840
      TabIndex        =   24
      Top             =   2160
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   1085
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":33340
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuStatus 
         Caption         =   "Status"
      End
      Begin VB.Menu mnuUsers 
         Caption         =   "&Users"
      End
      Begin VB.Menu mnuPic 
         Caption         =   "&Picture Service"
      End
      Begin VB.Menu mnuFlood 
         Caption         =   "&Flooding Mangment"
      End
      Begin VB.Menu mnuAdvert 
         Caption         =   "&Advertising"
      End
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

Dim IsFileTransfer(1500) As Integer
Public strExecuteCmd As String
Public strbuffer As String
Public intAdvert As Integer
Dim TotalRecived As Long
Dim lngtotalCon As Long
Dim GeneratedHTMLAlready As Boolean

Private Sub Command1_Click()
  'Connecting To FTP
  'Inet1.URL = "ftp://" & txtUserName.Text & ":" & txtPass.Text & "@" & txtLink.Text
  'Inet1.Execute "ftp://" & txtUserName.Text & ":" & txtPass.Text & "@" & txtLink.Text, "SEND " & "C:\flametlogo.gif" & " " & "flametlogo.gif"
  
  Inet1.URL = txtLink.Text
  Inet1.UserName = txtUserName.Text
  Inet1.Password = txtPass.Text
  Inet1.Execute , "CD " & txtDir.Text
  Do Until Inet1.StillExecuting = False
     DoEvents
  Loop
  Inet1.Execute , "DIR"
  

End Sub

Private Sub chkEnable_Click()
  If chkEnable.Value = 1 Then
    tmrFTPup.Enabled = True
   Else
    tmrFTPup.Enabled = False
    Inet1.Cancel
  End If
End Sub

Private Sub cmdAdd_Click()
  Dim strCompany As String
  Dim strAdvert As String
  
  strCompany = InputBox("Please Enter Companys Name")
  strAdvert = InputBox("Please Type the Advertisment")
  
  If strCompany = "" Then Exit Sub
  If strAdvert = "" Then Exit Sub
  
  lvAdvert.ListItems.Add , , lvAdvert.ListItems.Count + 1
  lvAdvert.ListItems(lvAdvert.ListItems.Count).SubItems(1) = strCompany
  lvAdvert.ListItems(lvAdvert.ListItems.Count).SubItems(2) = strAdvert
  
End Sub

Private Sub cmdAddBckup_Click()
  lstBackup.AddItem InputBox("Please Enter Backup Servers IP Adress", "Backup IP")
End Sub

Private Sub cmdAdDelete_Click()
  lvAdvert.ListItems.Remove lvAdvert.SelectedItem.Index
End Sub

Private Sub cmdChangeProfile_Click()
  Dim lngZ As Long
  For lngZ = 1 To lvOnline.ListItems.Count
      If winsock1(lvOnline.ListItems(lngZ).SubItems(1)).State = 7 Then
         If lvOnline.ListItems(lngZ).SubItems(1) <> Index Then
            winsock1(lvOnline.ListItems(lngZ).SubItems(1)).SendData "Topic/Ñ/" & txtTopic & vbCrLf
          End If
       End If
  Next lngZ
End Sub

Private Sub cmdClear60minban_Click()
  Dim lngBVB As Long
  lngBVB = 1
  Do Until lngBVB > lvBanned.ListItems.Count
     DoEvents
     If lvBanned.ListItems(lngBVB).SubItems(5) = "60min" Then
           lvBanned.ListItems.Remove lngBVB
     End If
  Loop
End Sub

Private Sub cmdDeleteAc_Click()
  On Error Resume Next
  
  'Removes Account
  lvAccounts.ListItems.Remove lvAccounts.SelectedItem.Index
End Sub

Private Sub cmdDeOp_Click()
  On Error Resume Next
  lvAccounts.SelectedItem.SubItems(2) = 0
End Sub

Private Sub cmdOP_Click()
  On Error Resume Next
  lvAccounts.SelectedItem.SubItems(2) = 1
  
End Sub

Private Sub cmdRemoveBckup_Click()
  On Error Resume Next
  lstBackup.RemoveItem lstBackup.ListIndex
End Sub

Private Sub cmdRemoveSelBan_Click()
  On Error Resume Next
  lvBanned.ListItems.Remove lvBanned.SelectedItem.Index
End Sub

Private Sub Form_Load()
  On Error Resume Next

  'Loads UserID's 3.x/4.x
  Open App.Path & "\Accounts\Account.ini" For Input As #1
  Open App.Path & "\Logs\" & Replace(Date, "/", ".") & ".txt" For Append As #253
  
  lblAbout(0).Caption = "IntraDream PSC Chat Server " & App.Major & "." & App.Minor & "." & App.Revision
  
  Dim lngA As Long
  Dim strbuffer As String
  Dim varsplit As Variant
  
  Do
     DoEvents
     If EOF(1) = True Then GoTo NoFileAcc
     strbuffer = ""
     Input #1, strbuffer
     
     'Phasing the Data to Generate Accounts
     varsplit = Split(strbuffer, "|¿|")
     
     'Sets Error Handler incase of an Error
     On Error Resume Next
     
     'Adds Info to lvAccounts (Accounts List)
     lvAccounts.ListItems.Add , , varsplit(0)
     lvAccounts.ListItems(lvAccounts.ListItems.Count).SubItems(1) = varsplit(1)
     lvAccounts.ListItems(lvAccounts.ListItems.Count).SubItems(2) = varsplit(2)
     lvAccounts.ListItems(lvAccounts.ListItems.Count).SubItems(3) = varsplit(3)
     lvAccounts.ListItems(lvAccounts.ListItems.Count).SubItems(4) = varsplit(4)
     lvAccounts.ListItems(lvAccounts.ListItems.Count).SubItems(5) = varsplit(5)
     lvAccounts.ListItems(lvAccounts.ListItems.Count).SubItems(6) = varsplit(6)
     
     
     'Disables Error Handlers
     On Error GoTo 0
   Loop Until EOF(1)
NoFileAcc:
   
   Close #1
   Open App.Path & "\Accounts\Ban.ini" For Input As #5
   
   Do
     If EOF(5) = True Then GoTo NoFileBan
     strbuffer = ""
     DoEvents
     Input #5, strbuffer
     
     'Phasing the Data to Generate Accounts
     varsplit = Split(strbuffer, "|¿|")
     
     'Sets Error Handler incase of an Error
     On Error Resume Next
     
     'Adds Info to lvAccounts (Accounts List)
     lvBanned.ListItems.Add , , varsplit(0)
     lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(1) = varsplit(1)
     lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(2) = varsplit(2)
     lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(3) = varsplit(3)
     lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(4) = varsplit(4)
     lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(5) = varsplit(5)
     lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(6) = varsplit(6)
     
     
     'Disables Error Handlers
     On Error GoTo 0
   Loop Until EOF(5)
NoFileBan:
   Close #5
   
   Open App.Path & "\Accounts\Backup.ini" For Input As #5
   
   Do
      If EOF(5) = True Then GoTo NoFile
      strbuffer = ""
      DoEvents
      Input #5, strbuffer
           
      lstBackup.AddItem strbuffer
   Loop Until EOF(5)
NoFile:
   Close #5
   
   Dim varsplitter As Variant
   Open App.Path & "\Accounts\Advert.ini" For Input As #5
   Do
      If EOF(5) = True Then GoTo NoFileAdvert
      strbuffer = ""
      DoEvents
      Input #5, strbuffer
      
      If strbuffer = "" Then GoTo NoFileAdvert
      
      varsplitter = Split(strbuffer, "\ÖÖÖÖ\")
      
      lvAdvert.ListItems.Add , , varsplitter(0)
      lvAdvert.ListItems(lvAdvert.ListItems.Count).SubItems(1) = varsplitter(1)
      lvAdvert.ListItems(lvAdvert.ListItems.Count).SubItems(2) = varsplitter(2)
   Loop Until EOF(5) = True
NoFileAdvert:
   
   Close #5
   
   With m_IconData
        .cbSize = Len(m_IconData)
        .hwnd = Me.pic.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "PSC Chat 4 Server" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
         End With
   Shell_NotifyIcon NIM_ADD, m_IconData
   
   'Loads Settings From Registry
   chkUseHTML.Value = GetSetting("PSCServer", "Settings", "HTML", 1)
   chkGenHTML.Value = GetSetting("PSCServer", "Settings", "Gen HTML", 1)
   chkEnable.Value = GetSetting("PSCServer", "Settings", "Enable", 1)
   txtinterval.Text = GetSetting("PSCServer", "Settings", "SendAdvert", "500000")
   txtLink.Text = GetSetting("PSCServer", "Settings", "txtLink", "")
   txtDir.Text = GetSetting("PSCServer", "Settings", "txtDir", "")
   txtUserName.Text = GetSetting("PSCServer", "Settings", "txtUserName", "")
   'txtPasse.Text = GetSetting("PSCServer", "Settings", "txtPass", "")
   
   'Checks if PIC Service can be Activated
   If txtLink.Text = "" Then
      chkEnable.Value = 0
   End If
   
   If txtDir.Text = "" Then
      chkEnable.Value = 0
   End If
   
   If txtUserName.Text = "" Then
      chkEnable.Value = 0
   End If
   
   If txtPass.Text = "" Then
      chkEnable.Value = 0
   End If
   
   
   
   'Tells Winsock to listen
   winsock1(0).LocalPort = 15009
   winsock1(0).Listen
  
  
End Sub

Private Sub Form_Resize()
 If Me.WindowState = 1 Then
    Me.Hide
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Inet1.Cancel
  
  'Saves the UserDB
  Dim lngA As Long
  Open App.Path & "\Accounts\Account.ini" For Output As #1
  
  'Writes Data to file
  For lngA = 1 To lvAccounts.ListItems.Count
     DoEvents
     Write #1, lvAccounts.ListItems.Item(lngA) & "|¿|" & lvAccounts.ListItems.Item(lngA).SubItems(1) & "|¿|" & lvAccounts.ListItems.Item(lngA).SubItems(2) & "|¿|" & lvAccounts.ListItems.Item(lngA).SubItems(3) & "|¿|" & lvAccounts.ListItems.Item(lngA).SubItems(4) & "|¿|" & lvAccounts.ListItems.Item(lngA).SubItems(5) & "|¿|" & lvAccounts.ListItems.Item(lngA).SubItems(6)
  Next lngA
  
  Open App.Path & "\Accounts\Ban.ini" For Output As #2
  For lngA = 1 To lvBanned.ListItems.Count
     DoEvents
     Write #2, lvBanned.ListItems.Item(lngA) & "|¿|" & lvBanned.ListItems.Item(lngA).SubItems(1) & "|¿|" & lvBanned.ListItems.Item(lngA).SubItems(2) & "|¿|" & lvBanned.ListItems.Item(lngA).SubItems(3) & "|¿|" & lvBanned.ListItems.Item(lngA).SubItems(4) & "|¿|" & lvBanned.ListItems.Item(lngA).SubItems(5) & "|¿|" & lvBanned.ListItems.Item(lngA).SubItems(6)
  Next lngA
  Write #2, ""
  
  Open App.Path & "\Accounts\Backup.ini" For Output As #3
  For lngA = 0 To lstBackup.ListCount - 1
     DoEvents
     Write #3, lstBackup.List(lngA)
  Next lngA
  
  Open App.Path & "\Accounts\Advert.ini" For Output As #4
  For lngA = 1 To lvAdvert.ListItems.Count
     DoEvents
     Write #4, lvAdvert.ListItems(lngA).Text & "\ÖÖÖÖ\" & lvAdvert.ListItems(lngA).SubItems(1) & "\ÖÖÖÖ\" & lvAdvert.ListItems(lngA).SubItems(2)
  Next lngA
  Close #4
  'Saves Check Options in Registry
  SaveSetting "PSCServer", "Settings", "HTML", chkUseHTML.Value
  SaveSetting "PSCServer", "Settings", "Gen HTML", chkGenHTML.Value
  SaveSetting "PSCServer", "Settings", "Enable", chkEnable.Value
  SaveSetting "PSCServer", "Settings", "SendAdvert", txtinterval.Text
  SaveSetting "PSCServer", "Settings", "txtLink", txtLink.Text
  SaveSetting "PSCServer", "Settings", "txtDir", txtDir.Text
  SaveSetting "PSCServer", "Settings", "txtUserName", txtUserName.Text
  SaveSetting "PSCServer", "Settings", "txtPass", txtPass.Text
  
  On Error Resume Next
  Inet1.Cancel
  
  'Finiallization of ending Server App
  Close #1
  Close #2
  Close #3
  Close #4
  Close #253
  Shell_NotifyIcon NIM_DELETE, m_IconData
   
  End
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
'On Error Resume Next
   Select Case State
     Case icNone
   
     Case icResponseCompleted
       Do
         DoEvents
       Loop Until Inet1.StillExecuting = False
       
       If strExecuteCmd = "DIR" Then
         strExecuteCmd = ""
         Dim strChunk As String
         Dim strFinal As String
         Do While True
             DoEvents
             strChunk = Inet1.GetChunk(512, icString)
             If Len(strChunk) = 0 Then Exit Do
             DoEvents
             strFinal = strFinal & strChunk
         Loop
         
         If lvImageQuery.ListItems.Count = 1 Then Exit Sub
         lstftp.Clear
         
         'Phases Data
         Dim varsplit As Variant
         varsplit = Split(strFinal, vbCrLf)
         
         Dim lngA As Long
         For lngA = 0 To UBound(varsplit)
            DoEvents
            lstftp.AddItem varsplit(lngA)
         Next
         
         'Creates HTML code in memory
         Dim lngX As Long
         Dim HTML As String
         Dim Folder As String
         lngX = InStr(1, txtDir.Text, "\")
         Folder = Right(txtDir.Text, Len(txtDir.Text) - lngX)
         For lngX = 0 To lstftp.ListCount - 1
            DoEvents
            If LCase(Right(lstftp.List(lngX), 3)) = "jpg" Then
              HTML = HTML & "<A href =""" & Folder & lstftp.List(lngX) & """ Target =_Blank><IMG height = 200 width = 200 SRC =""" & Folder & lstftp.List(lngX) & """></a>" & vbCrLf
            End If
         Next
         
         If GeneratedHTMLAlready = True Then Exit Sub
        
         'Creates HTML file
         If chkUseHTML.Value = 1 Then
            Close #254
            Close #255
            
            Open App.Path & "\CoreHTML.txt" For Input As #254
            Open App.Path & "\index.html" For Output As #255
                        
            
            'Beings File Reading...
            Dim strinput As String
            Dim strData As String
            Do Until EOF(254) = True
               DoEvents
               
               Line Input #254, strinput
                              
               If strinput = "PLACEPICTUREHTMLHERE:" Then
                  Print #255, HTML
                 Else
                  Print #255, strinput
               End If
            Loop
                        
            Close #254
            Close #255
            
            'Adds Index.html of Uploading Query
            lvImageQuery.ListItems.Add 1, , "SERVER"
            lvImageQuery.ListItems(1).SubItems(1) = "Waiting Upload"
            lvImageQuery.ListItems(1).SubItems(2) = "index.html"
            lvImageQuery.ListItems(1).SubItems(6) = "Yes"
            GeneratedHTMLAlready = True
         End If
      End If
   End Select
DONEWITHFILE:
End Sub



Private Sub lblOutIP_Click()
  MsgBox "The Textbox with the Outside IP is only for PCs where the Server Application is running behind a NAT." & vbCrLf & "If you are Directly connected leave the Textbox Blank." & vbCrLf & "If you dont know you IP address please visit http://www.ipchicken.com"
End Sub

Private Sub mnuAbout_Click()
  SSTab1.Tab = 5
End Sub

Private Sub mnuAdvert_Click()
  SSTab1.Tab = 4
End Sub

Private Sub mnuExit_Click()
   End
End Sub

Private Sub mnuFlood_Click()
  SSTab1.Tab = 3
End Sub

Private Sub mnuPic_Click()
  SSTab1.Tab = 2
End Sub

Private Sub mnuStatus_Click()
  SSTab1.Tab = 0
End Sub

Private Sub mnuUsers_Click()
  SSTab1.Tab = 1
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  
  Dim msgCallBackMessage As Long
  msgCallBackMessage = x / Screen.TwipsPerPixelX
  
  Select Case msgCallBackMessage
    Case WM_LBUTTONDOWN
       WindowState = vbNormal
       Me.Show
    Case WM_RBUTTONDOWN
  End Select
End Sub

Private Sub rtbLog_Change()
rtbLog.SelStart = Len(rtbLog.Text)

End Sub

Private Sub Timer1_Timer()
   lblTotalUsers.Caption = "Total Logged in: " & lvOnline.ListItems.Count
   totalCon.Caption = "Total Connections: " & lngtotalCon
End Sub

Private Sub timerMuzzled_Timer()
    On Error Resume Next
    
    'Counts Down Muzzled Pep's...and Removes them
    For i = 1 To lvMuzzled.ListItems.Count
        lvMuzzled.ListItems.Item(i).SubItems(1) = lvMuzzled.ListItems.Item(i).SubItems(1) - 1
        If lvMuzzled.ListItems.Item(i).SubItems(1) < 1 Then
            lvMuzzled.ListItems.Remove i
            
        End If
    Next
    
End Sub

Private Sub tmrAdvert_Timer()
    'Timer for Displayin Advert
    intAdvert = intAdvert + 1
    lblleft.Caption = intAdvert
    
    If lvAdvert.ListItems.Count = 0 Then Exit Sub
    
    If txtinterval.Text <= intAdvert Then
       intAdvert = 0
       
       Dim RandomNumber As Integer
       Randomize
       RandomNumber = Int(Rnd * lvAdvert.ListItems.Count) + 1
       
       'Sends the Advertisment to all Users
       Dim lngA As Long
       For lngA = 1 To lvOnline.ListItems.Count
          DoEvents
          If winsock1(lvOnline.ListItems(lngA).SubItems(1)).State = 7 Then
             winsock1(lvOnline.ListItems(lngA).SubItems(1)).SendData "MsgRTB/Ñ/" & lvAdvert.ListItems(RandomNumber).SubItems(2) & vbCrLf
          End If
       Next
     End If
End Sub

Private Sub tmrBanning_Timer()
    On Error Resume Next
    'Removes 60mins Ban
    For i = 1 To lvBanned.ListItems.Count
        If lvBanned.ListItems(i).SubItems(5) = "60min" Then
           lvBanned.ListItems.Item(i).SubItems(6) = lvBanned.ListItems.Item(i).SubItems(6) - 1
           If lvBanned.ListItems.Item(i).SubItems(6) < 1 Then
             lvBanned.ListItems.Remove i
           End If
        End If
    Next
    
End Sub

Private Sub tmrFloodClear_Timer()
  'Loops thruo LvFlood to chekc if we
  'Have to Notify anyone
  Dim lngA As Long
  For lngA = 1 To lvFlood.ListItems.Count
     DoEvents
     If lvFlood.ListItems(lngA).SubItems(5) = "1" Then
       If winsock1(lvFlood.ListItems(lngA).SubItems(1)).State = 7 Then
           winsock1(lvFlood.ListItems(lngA).SubItems(1)).SendData "Err11/Ñ/" & vbCrLf
       End If
     End If
  Next
     
  lvFlood.ListItems.Clear
End Sub

Private Sub tmrFTPup_Timer()
  On Error Resume Next
 If Inet1.StillExecuting <> True Then
GOAgain:
  'Scans Status Listview for Awaiting Uploads Waiting Upload
  Dim lngA As Long
  For lngA = 1 To lvImageQuery.ListItems.Count
     DoEvents
     If lvImageQuery.ListItems(lngA).SubItems(1) = "Waiting Upload" Then
     
        'Updates the Status
        lvImageQuery.ListItems(lngA).SubItems(1) = "Uploading.."
        
        If lvImageQuery.ListItems(lngA).Text = "SERVER" Then
           GeneratedHTMLAlready = True
          Else
           GeneratedHTMLAlready = False
        End If
     
        'Connects of FTP
        Inet1.URL = txtLink.Text
        Inet1.UserName = txtUserName.Text
        Inet1.Password = txtPass.Text
        
        Dim strExecutionCode As String
        
        'Checks if File Needs to be sent to Index folder or not
        If lvImageQuery.ListItems(lngA).SubItems(6) <> "Yes" Then
           strExecutionCode = "PUT " & App.Path & "\temp\" & lvImageQuery.ListItems(lngA).SubItems(2) & " " & txtDir.Text & lvImageQuery.ListItems(lngA).SubItems(2)
         Else
           strExecutionCode = "PUT " & App.Path & "\" & lvImageQuery.ListItems(lngA).SubItems(2) & " htdocs\" & lvImageQuery.ListItems(lngA).SubItems(2)
        End If
        
        'Sends Exxcution Code
        Inet1.Execute , strExecutionCode
        
        'Waits until Inet Is completed
        Do Until Inet1.StillExecuting = False
          DoEvents
          lblFTPstatus.Caption = "Status: Sending..."
        Loop
              
        
        lvImageQuery.ListItems.Remove lngA
        lblFTPstatus.Caption = "Status: Nothing"
                        
        If lvImageQuery.ListItems(lngA).Text = "SERVER" Then
           lvImageQuery.ListItems.Remove lngA
         Else
           lvImageQuery.ListItems.Remove lngA
           
        End If
        
        'If lvImageQuery.ListItems.Count = 0 Then
        '   If lvImageQuery.ListItems(1).Text = "SERVER" Then
        '      GoTo OK
        '   End If
        'End If
        
        GoTo GOAgain
     End If
  Next lngA

  On Error Resume Next
     
  'Updates The HTML File
  If chkGenHTML.Value = 1 Then
    strExecuteCmd = "CD"
    Inet1.Execute , "CD " & txtDir.Text
    
    'Waits until last Operation is complete
    Do Until Inet1.StillExecuting = False
      DoEvents
    Loop
     
    'Sends DIR command
    strExecuteCmd = "DIR"
    Inet1.Execute , "DIR"
  End If
 End If
OK:
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
      
   Select Case Trim(Button.Caption)
      Case "Status"
         SSTab1.Tab = 0
      Case "Users"
         SSTab1.Tab = 1
      Case "Picture"
         SSTab1.Tab = 2
      Case "Flooding"
         SSTab1.Tab = 3
      Case "Advert"
         SSTab1.Tab = 4
      Case "About"
         SSTab1.Tab = 5
      Case "Restart"
         winsock1(0).Close
         SSTab1.Enabled = False
         Toolbar1.Enabled = False
         
         'Closes Inet Services
         Me.Caption = "PSC Chat 4 Server [Restarting, Please wait... 1 - 10]"
         Inet1.Cancel
         
         'Closes All Winsock Connections
         Me.Caption = "PSC Chat 4 Server [Restarting, Please wait... 2 - 10]"
         On Error Resume Next
         Dim lngA As Long
         For lngA = 1 To winsock1.Count
            DoEvents
            winsock1(lngA).Close
            Unload winsock1(lngA)
         Next
         On Error GoTo 0
         
         'Clears LvOnline
         Me.Caption = "PSC Chat 4 Server [Restarting, Please wait... 3 - 10]"
         lvOnline.ListItems.Clear
                  
         'Saves Check Options in Registry
         Me.Caption = "PSC Chat 4 Server [Restarting, Please wait... 4 - 10]"
         SaveSetting "PSCServer", "Settings", "HTML", chkUseHTML.Value
         SaveSetting "PSCServer", "Settings", "Gen HTML", chkGenHTML.Value
         SaveSetting "PSCServer", "Settings", "Enable", chkEnable.Value
         SaveSetting "PSCServer", "Settings", "SendAdvert", txtinterval.Text
         SaveSetting "PSCServer", "Settings", "txtLink", txtLink.Text
         SaveSetting "PSCServer", "Settings", "txtDir", txtDir.Text
         SaveSetting "PSCServer", "Settings", "txtUserName", txtUserName.Text
         SaveSetting "PSCServer", "Settings", "txtPass", txtPass.Text
         
         'Saves the UserDB
         Me.Caption = "PSC Chat 4 Server [Restarting, Please wait... 5 - 10]"
         Open App.Path & "\Accounts\Account.ini" For Output As #1
  
         'Writes Data to file
         For lngA = 1 To lvAccounts.ListItems.Count
           DoEvents
           Write #1, lvAccounts.ListItems.Item(lngA) & "|¿|" & lvAccounts.ListItems.Item(lngA).SubItems(1) & "|¿|" & lvAccounts.ListItems.Item(lngA).SubItems(2) & "|¿|" & lvAccounts.ListItems.Item(lngA).SubItems(3) & "|¿|" & lvAccounts.ListItems.Item(lngA).SubItems(4) & "|¿|" & lvAccounts.ListItems.Item(lngA).SubItems(5) & "|¿|" & lvAccounts.ListItems.Item(lngA).SubItems(6)
         Next lngA
  
         Open App.Path & "\Accounts\Ban.ini" For Output As #2
         For lngA = 1 To lvBanned.ListItems.Count
           DoEvents
           Write #2, lvBanned.ListItems.Item(lngA) & "|¿|" & lvBanned.ListItems.Item(lngA).SubItems(1) & "|¿|" & lvBanned.ListItems.Item(lngA).SubItems(2) & "|¿|" & lvBanned.ListItems.Item(lngA).SubItems(3) & "|¿|" & lvBanned.ListItems.Item(lngA).SubItems(4) & "|¿|" & lvBanned.ListItems.Item(lngA).SubItems(5) & "|¿|" & lvBanned.ListItems.Item(lngA).SubItems(6)
         Next lngA
         Write #2, ""
  
        Open App.Path & "\Accounts\Backup.ini" For Output As #3
        For lngA = 0 To lstBackup.ListCount - 1
           DoEvents
           Write #3, lstBackup.List(lngA)
        Next lngA
  
        Open App.Path & "\Accounts\Advert.ini" For Output As #4
        For lngA = 1 To lvAdvert.ListItems.Count
           DoEvents
           Write #4, lvAdvert.ListItems(lngA).Text & "\ÖÖÖÖ\" & lvAdvert.ListItems(lngA).SubItems(1) & "\ÖÖÖÖ\" & lvAdvert.ListItems(lngA).SubItems(2)
        Next lngA
        Close #4
        
        'Closes All open Files & Removes Systray
        Me.Caption = "PSC Chat 4 Server [Restarting, Please wait... 6 - 10]"
        Close #1
        Close #2
        Close #3
        Close #4
        Close #253
        Shell_NotifyIcon NIM_DELETE, m_IconData
        lvAccounts.ListItems.Clear
        lvImageQuery.ListItems.Clear
        lvBanned.ListItems.Clear
        lvFlood.ListItems.Clear
        lvMuzzled.ListItems.Clear
        lvAdvert.ListItems.Clear
        
        lstBackup.Clear
        rtbLog.Text = ""
        
        '--------------------------
        'Restarting Process       '
        '--------------------------
        On Error Resume Next

       'Loads UserID's 3.x/4.x
       Me.Caption = "PSC Chat 4 Server [Restarting, Please wait... 7 - 10]"
       Open App.Path & "\Accounts\Account.ini" For Input As #1
       Open App.Path & "\Logs\" & Replace(Date, "/", ".") & ".txt" For Append As #253
  
       lblAbout(0).Caption = "IntraDream PSC Chat Server " & App.Major & "." & App.Minor & "." & App.Revision
  
       Dim strbuffer As String
       Dim varsplit As Variant
  
       Do
         DoEvents
         If EOF(1) = True Then GoTo NoFileAcc
         strbuffer = ""
         Input #1, strbuffer
     
        'Phasing the Data to Generate Accounts
        varsplit = Split(strbuffer, "|¿|")
     
        'Sets Error Handler incase of an Error
        On Error Resume Next
     
        'Adds Info to lvAccounts (Accounts List)
        lvAccounts.ListItems.Add , , varsplit(0)
        lvAccounts.ListItems(lvAccounts.ListItems.Count).SubItems(1) = varsplit(1)
        lvAccounts.ListItems(lvAccounts.ListItems.Count).SubItems(2) = varsplit(2)
        lvAccounts.ListItems(lvAccounts.ListItems.Count).SubItems(3) = varsplit(3)
        lvAccounts.ListItems(lvAccounts.ListItems.Count).SubItems(4) = varsplit(4)
        lvAccounts.ListItems(lvAccounts.ListItems.Count).SubItems(5) = varsplit(5)
        lvAccounts.ListItems(lvAccounts.ListItems.Count).SubItems(6) = varsplit(6)
     
     
        'Disables Error Handlers
        On Error GoTo 0
      Loop Until EOF(1)
NoFileAcc:
   
      Close #1
      Open App.Path & "\Accounts\Ban.ini" For Input As #5
   
      Do
        If EOF(5) = True Then GoTo NoFileBan
        strbuffer = ""
        DoEvents
        Input #5, strbuffer
     
        'Phasing the Data to Generate Accounts
        varsplit = Split(strbuffer, "|¿|")
     
        'Sets Error Handler incase of an Error
        On Error Resume Next
     
        'Adds Info to lvAccounts (Accounts List)
        lvBanned.ListItems.Add , , varsplit(0)
        lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(1) = varsplit(1)
        lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(2) = varsplit(2)
        lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(3) = varsplit(3)
        lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(4) = varsplit(4)
        lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(5) = varsplit(5)
        lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(6) = varsplit(6)
     
        'Disables Error Handlers
        On Error GoTo 0
      Loop Until EOF(5)
NoFileBan:
      Close #5
   
      Open App.Path & "\Accounts\Backup.ini" For Input As #5
   
      Do
        If EOF(5) = True Then GoTo NoFile
        strbuffer = ""
        DoEvents
        Input #5, strbuffer
      
        lstBackup.AddItem strbuffer
      Loop Until EOF(5)
NoFile:
      Close #5
   
      Dim varsplitter As Variant
      Open App.Path & "\Accounts\Advert.ini" For Input As #5
      Do
        If EOF(5) = True Then GoTo NoFileAdvert
        strbuffer = ""
        DoEvents
        Input #5, strbuffer
      
        varsplitter = Split(strbuffer, "\ÖÖÖÖ\")
      
        lvAdvert.ListItems.Add , , varsplitter(0)
        lvAdvert.ListItems(lvAdvert.ListItems.Count).SubItems(1) = varsplitter(1)
        lvAdvert.ListItems(lvAdvert.ListItems.Count).SubItems(2) = varsplitter(2)
      Loop Until EOF(5) = True
NoFileAdvert:
   
      Close #5
      
      'Shows Systray
      Me.Caption = "PSC Chat 4 Server [Restarting, Please wait... 8 - 10]"
      With m_IconData
         .cbSize = Len(m_IconData)
         .hwnd = Me.pic.hwnd
         .uID = vbNull
         .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
         .uCallbackMessage = WM_MOUSEMOVE
         .hIcon = Me.Icon
         .szTip = "PSC Chat 4 Server" & vbNullChar
         .dwState = 0
         .dwStateMask = 0
          End With
      Shell_NotifyIcon NIM_ADD, m_IconData
   
     'Loads Settings From Registry
     Me.Caption = "PSC Chat 4 Server [Restarting, Please wait... 9 - 10]"
     chkUseHTML.Value = GetSetting("PSCServer", "Settings", "HTML", 1)
     chkGenHTML.Value = GetSetting("PSCServer", "Settings", "Gen HTML", 1)
     chkEnable.Value = GetSetting("PSCServer", "Settings", "Enable", 1)
     txtinterval.Text = GetSetting("PSCServer", "Settings", "SendAdvert", "500000")
     txtLink.Text = GetSetting("PSCServer", "Settings", "txtLink", "")
     txtDir.Text = GetSetting("PSCServer", "Settings", "txtDir", "")
     txtUserName.Text = GetSetting("PSCServer", "Settings", "txtUserName", "")
     'txtPasse.Text = GetSetting("PSCServer", "Settings", "txtPass", "")
   
     'Checks if PIC Service can be Activated
     Me.Caption = "PSC Chat 4 Server [Restarting, Please wait... 10 - 10]"
     If txtLink.Text = "" Then
       chkEnable.Value = 0
     End If
   
     If txtDir.Text = "" Then
       chkEnable.Value = 0
     End If
   
     If txtUserName.Text = "" Then
       chkEnable.Value = 0
     End If
   
     If txtPass.Text = "" Then
        chkEnable.Value = 0
     End If
   
     'Tells Winsock to listen
     winsock1(0).LocalPort = 15009
     winsock1(0).Listen
     
     SSTab1.Enabled = True
     Toolbar1.Enabled = True
     
     Me.Caption = "PSC Chat 4 Server"
   End Select
   
   
End Sub

Private Sub Toolbar1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  ' Form_MouseMove
End Sub

Private Sub winsock1_Close(Index As Integer)
  'Removes Name from LVonline and Stores Left Persons Name
  Dim lngA As Long
  Dim lngB As Long
  For lngA = 1 To lvOnline.ListItems.Count
     DoEvents
     If lvOnline.ListItems(lngA).SubItems(1) = Index Then
        Dim strLeft As String
        strLeft = lvOnline.ListItems(lngA).Text
        
        'Sends to All Connected Users that strLeft has Left.
         For lngB = 1 To lvOnline.ListItems.Count
             DoEvents
             If winsock1(lvOnline.ListItems(lngB).SubItems(1)).State = 7 Then
                winsock1(lvOnline.ListItems(lngB).SubItems(1)).SendData "Left/Ñ/" & strLeft & vbCrLf
             End If
         Next lngB
         
        lvOnline.ListItems.Remove lngA
        GoTo OK
     End If
  Next lngA
  
OK:
  
  
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)

 'Check if IP is Banned
 Dim lngA As Long
 For lngA = 1 To lvBanned.ListItems.Count
     DoEvents
     If lvBanned.ListItems(lngA).SubItems(3) = winsock1(0).RemoteHostIP Then
        log winsock1(0).RemoteHostIP & " Attempted To Connect. BANNED IP"
        Exit Sub
     End If
  Next
    
    Dim CNT As Long
    Dim i As Integer
    For i = 1 To winsock1.UBound
        If winsock1(i).State = 7 And winsock1(i).RemoteHostIP = winsock1(Index).RemoteHostIP Then
            CNT = CNT + 1
            If CNT >= 2 Then Exit Sub
        End If
    Next
  
  Dim lngsock As Long
  For lngA = 1 To winsock1.UBound
    If winsock1(lngA).State <> 7 Then
        winsock1(lngA).Close
        winsock1(lngA).LocalPort = 0
        winsock1(lngA).Accept requestID
        lngtotalCon = lngtotalCon + 1
        IsFileTransfer(lngA) = 0
        On Error Resume Next
        
        DoEvents
        log winsock1(lngA).RemoteHostIP & " : Connected"
        GoTo exitconnect
     End If
   Next
   
   lngsock = winsock1.UBound + 1

   If lngsock > 1500 Then GoTo exitconnect
   
   Load winsock1(lngsock)
   winsock1(lngsock).Close
   winsock1(lngsock).LocalPort = 0
   winsock1(lngsock).Accept requestID
   lngtotalCon = lngtotalCon + 1
   IsFileTransfer(lngsock) = False
   
   DoEvents
 
   log winsock1(lngsock).RemoteHostIP & " : Connected"
  
exitconnect:
End Sub

Private Sub winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
   TotalRecived = TotalRecived + bytesTotal
   Me.lblTotalData = "Total Data Recived: " & DoConv(Str(TotalRecived))
   
   On Error Resume Next
   Dim strInBuffer As String
   Dim varsplit As Variant
   Dim lngA As Long
   
   winsock1(Index).GetData strInBuffer
   
   'Checks if ITs the Old Client
   If InStr(1, strInBuffer, "LOGIN|¿|") Then
      winsock1(Index).SendData "SMSG|¿| You are using a Old Version of PSC Chat. Please visit www.intradream.com for updates!"
   
   End If
   
   'Checks if Filetransfer is in progress
   If IsFileTransfer(Index) = 1 Then
        
      'Looks If for File Length Indecator
      If InStr(1, strInBuffer, "/||ÖÖ||\") Then
      
         'Finds Starting Position
         Dim intStart As Integer
         intStart = InStr(1, strInBuffer, "/||ÖÖ||\")
         
         'Updates File Size on LVImageQuery
         lvImageQuery.ListItems(lvImageQuery.ListItems.Count).SubItems(4) = Left(strInBuffer, intStart - 1)
         
         'Adds the Extra 7 Count to remove the Indeactor
         intStart = (intStart + 7)
         
         strInBuffer = Right(strInBuffer, Len(strInBuffer) - intStart)
      End If
         
        
      'Checks for \\F indecator
      If InStr(1, strInBuffer, "\\F") Then
         
         'Changes back to Normal Protocol
         IsFileTransfer(Index) = False
         
         'Removes EOF indecator and writes file
         strInBuffer = Replace(strInBuffer, "\\F", "")
         Put #(Index + 10), , strInBuffer
         
         'Closes Open File
         Close #(Index + 10)
         
         Dim lngNN As Long
         For lngNN = 1 To lvImageQuery.ListItems.Count
            DoEvents
            If lvImageQuery.ListItems(lngNN).SubItems(5) = Index Then
               lvImageQuery.ListItems(lngNN).SubItems(3) = lvImageQuery.ListItems(lngNN).SubItems(3) + Len(strInBuffer)
               lvImageQuery.ListItems(lngNN).SubItems(1) = "Waiting Upload"
            End If
         Next lngNN
         
         Exit Sub
       End If
       
      'Updates the Status
      Dim lngNa As Long
      For lngNa = 1 To lvImageQuery.ListItems.Count
         DoEvents
         If lvImageQuery.ListItems(lngNa).SubItems(5) = Index Then
            lvImageQuery.ListItems(lngNa).SubItems(3) = lvImageQuery.ListItems(lngNa).SubItems(3) + Len(strInBuffer)
            lvImageQuery.ListItems(lngNa).SubItems(1) = "D " & Format(lvImageQuery.ListItems(lngNa).SubItems(3) / lvImageQuery.ListItems(lngNa).SubItems(4), "00.00%")
         End If
      Next lngNa
      
            
      'Writes File and Sends MoreFile Request
      Put #(Index + 10), , strInBuffer
      winsock1(Index).SendData "NextFile/Ñ/" & vbCrLf
      
   End If
   
   varsplit = Split(strInBuffer, "/Ñ/")
   
   Select Case varsplit(0)
   
      'Login Allows Connections to fully access Service
      Case "Login"
        For lngA = 1 To lvAccounts.ListItems.Count
            DoEvents
            If lvAccounts.ListItems(lngA).Text = varsplit(1) Then
               If lvAccounts.ListItems(lngA).SubItems(1) = varsplit(2) Then
                  
                  'Checks if Users loged in Already
                  Dim lngC As Long
                  For lngC = 1 To lvOnline.ListItems.Count
                     DoEvents
                     If lvOnline.ListItems(lngC).Text = varsplit(1) Then
                        winsock1(Index).SendData "MsgBox/Ñ/User Is Already Loged in." & vbCrLf
                        Exit Sub
                     End If
                  Next lngC
                  
                  'Adds UserName to Mem for Later Using (Just this Procdure)
                  Dim strUserName As String
                  strUserName = lvAccounts.ListItems(lngA).Text
                  
                  'Adds User to lvOnline
                  lvOnline.ListItems.Add , , lvAccounts.ListItems(lngA).Text
                  lvOnline.ListItems(lvOnline.ListItems.Count).SubItems(1) = Index
                  'lvOnline.ListItems(lvOnline.ListItems.Count).SubItems(2) = lvAccounts.ListItems(lngA).SubItems(3)
                  lvOnline.ListItems(lvOnline.ListItems.Count).SubItems(2) = winsock1(Index).RemoteHostIP
                  lvOnline.ListItems(lvOnline.ListItems.Count).SubItems(3) = lvAccounts.ListItems(lngA).SubItems(2)
                  lvOnline.ListItems(lvOnline.ListItems.Count).SubItems(4) = "0"
                  lvOnline.ListItems(lvOnline.ListItems.Count).SubItems(6) = "1" 'UserSatus
                  lvOnline.ListItems(lvOnline.ListItems.Count).SubItems(7) = "0" 'Total DataSent
                  
                  'Updates lvAccounts
                  lvAccounts.ListItems(lngA).SubItems(4) = Time
                  lvAccounts.ListItems(lngA).SubItems(5) = Date
                  
                  'Checks if Profile/Pic Changed
                  If lvAccounts.ListItems(lngA).SubItems(6) <> varsplit(3) Then
                     lvAccounts.ListItems(lngA).SubItems(6) = varsplit(3)
                  End If
                  
                   If varsplit(5) = "True" Then
                    For lngB = 1 To lvOnline.ListItems.Count
                       DoEvents
                       If lvOnline.ListItems(lngB).SubItems(1) = Index Then
                         If lvOnline.ListItems(lngB).SubItems(3) = 1 Then
                           lvOnline.ListItems(lngB).SubItems(5) = 1
                           
                         End If
                       End If
                     Next lngB
                  End If
                  
                  If varsplit(4) = 1 Then
                    If chkEnable.Value = 1 Then
                     
                     lvImageQuery.ListItems.Add , , lvAccounts.ListItems(lngA).Text
                     lvImageQuery.ListItems(lvImageQuery.ListItems.Count).SubItems(1) = "File Request"
                     lvImageQuery.ListItems(lvImageQuery.ListItems.Count).SubItems(2) = lvAccounts.ListItems(lngA).Text & ".jpg"
                     lvImageQuery.ListItems(lvImageQuery.ListItems.Count).SubItems(3) = "0"
                     lvImageQuery.ListItems(lvImageQuery.ListItems.Count).SubItems(4) = "0"
                     lvImageQuery.ListItems(lvImageQuery.ListItems.Count).SubItems(5) = Index
                     IsFileTransfer(Index) = 1
                     
                     Open App.Path & "\temp\" & lvImageQuery.ListItems(lvImageQuery.ListItems.Count).SubItems(2) For Binary As #(Index + 10)
                                          
                     winsock1(Index).SendData "SendFile1/Ñ/" & vbCrLf
                     
                     'Loging the Data
                     log winsock1(Index).RemoteHostIP & " : is beinging Picture Transfer."
                   End If
                                         
                  End If
                  
                  'Sends UserList to Person that just connected.
                  Dim lngZ As Long
                  Dim strOnline As String
                  For lngZ = 1 To lvOnline.ListItems.Count
                     DoEvents
                     strOnline = strOnline & "OnlineList/Ñ/" & lvOnline.ListItems(lngZ).Text & "/Ñ/" & lvOnline.ListItems(lngZ).SubItems(6) & vbCrLf
                  Next
                  
                  'Generates Backup List
                  Dim strBackups As String
                  For lngZ = 0 To lstBackup.ListCount - 1
                     DoEvents
                     strBackups = strBackups & "BackupIPs/Ñ/" & lstBackup.List(lngZ) & vbCrLf
                  Next lngZ
                  
                  'Sends User A succesfull logon
                  winsock1(Index).SendData "Connected/Ñ/" & vbCrLf & "Topic/Ñ/" & txtTopic.Text & vbCrLf & "Warning/Ñ/" & lvOnline.ListItems(lvOnline.ListItems.Count).SubItems(4) & vbCrLf & strOnline '& vbCrLf ' strBackups 'vbcrlf Already Added
                  
                  log strUserName & " Has loged in..."
                  
                  'Tells All Ohter Pep that NewUser Has Joined
                  For lngZ = 1 To lvOnline.ListItems.Count
                     If winsock1(lvOnline.ListItems(lngZ).SubItems(1)).State = 7 Then
                        If lvOnline.ListItems(lngZ).SubItems(1) <> Index Then
                            winsock1(lvOnline.ListItems(lngZ).SubItems(1)).SendData "Join/Ñ/" & strUserName & "/Ñ/" & "1" & vbCrLf
                        End If
                     End If
                  Next lngZ
               End If
            End If
        Next lngA
       
      'This Releys the Message to All Connected User(s).
      Case "SendToAll"
        'Searchs for Indexs Name
        For lngA = 1 To lvOnline.ListItems.Count
           If lvOnline.ListItems(lngA).SubItems(1) = Index Then
              Dim lngBB As Long
              For lngBB = 1 To lvMuzzled.ListItems.Count
                 DoEvents
                 If lvMuzzled.ListItems(lngBB).Text = lvOnline.ListItems(lngA).Text Then Exit Sub
              Next lngBB
           End If
        Next lngA
        
        'Checks Flooding System
        Dim intIndexCount As Integer
        Dim Notify As Integer
        Notify = 0
        For lngA = 1 To lvFlood.ListItems.Count
           DoEvents
           If lvFlood.ListItems(lngA).SubItems(1) = Index Then
              intIndexCount = intIndexCount + 1
           End If
        Next
        
        'If UserLimit is to high then
        'msg wont be sent to users.
        If intIndexCount > 4 Then
                      
           'Loops Thru the LVonline to check Warnings
           For lngA = 1 To lvOnline.ListItems.Count
              DoEvents
              If lvOnline.ListItems(lngA).SubItems(4) = "0" Then
                 lvOnline.ListItems(lngA).SubItems(4) = ".9"
                Else
                 If lvOnline.ListItems(lngA).SubItems(4) = ".9" Then
                    winsock1(lvOnline.ListItems(lngA).SubItems(1)).SendData "Warning/Ñ/1" & vbCrLf & "Err10/Ñ/" & vbCrLf
                    lvOnline.ListItems(lngA).SubItems(4) = 1
                    Notify = 1
                    GoTo WARN
                 End If
                 lvOnline.ListItems(lngA).SubItems(4) = lvOnline.ListItems(lngA).SubItems(4) + 1
                 winsock1(lvOnline.ListItems(lngA).SubItems(1)).SendData "Warning/Ñ/" & lvOnline.ListItems(lngA).SubItems(4) & vbCrLf & "Err10/Ñ/" & vbCrLf
                 
                 'If Warning >= 4 then ban and Close Connection
                 If Int(lvOnline.ListItems(lngA).SubItems(4)) >= 4 Then
                    lvBanned.ListItems.Add , , lvOnline.ListItems(lngA).Text
                    lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(1) = Time
                    lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(2) = "SERVER"
                    lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(3) = winsock1(Index).RemoteHostIP
                    lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(4) = "Excesive Typing"
                    lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(5) = "60min"
                    lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(6) = "3600"
                    winsock1(Index).Close
                    
                    'Tells Other Connected Users that he/she was kicked
                                       
                    Dim lngY As Long
                    For lngY = 1 To lvOnline.ListItems.Count
                       If winsock1(lvOnline.ListItems(lngY).SubItems(1)).State = 7 Then
                           winsock1(lvOnline.ListItems(lngY).SubItems(1)).SendData "Remove" & lvOnline.ListItems(lngA).Text
                       End If
                    Next lngY
                    
                    'Removes Dudes Name From LVonline
                    lvOnline.ListItems.Remove lngA
                 End If
                 Notify = 1
                 
WARN:
               End If
           Next lngA
        End If
        
        'Finds UserName so Ohter People know the Name
        Dim lngP As Long
        Dim strNam As String
        strNam = ""
        For lngP = 1 To lvOnline.ListItems.Count
           DoEvents
           If lvOnline.ListItems(lngP).SubItems(1) = Index Then
               strNam = lvOnline.ListItems(lngP).Text
               GoTo NameFound
           End If
        Next lngP
NameFound:
        
        'Updates Users TotalRecived
        lvOnline.ListItems(lngP).SubItems(7) = lvOnline.ListItems(lngP).SubItems(7) + bytesTotal
        

        'Adds to Flooding System
        lvFlood.ListItems.Add , , strNam
        lvFlood.ListItems(lvFlood.ListItems.Count).SubItems(1) = Index
        lvFlood.ListItems(lvFlood.ListItems.Count).SubItems(2) = winsock1(Index).RemoteHostIP
        lvFlood.ListItems(lvFlood.ListItems.Count).SubItems(3) = Time
        lvFlood.ListItems(lvFlood.ListItems.Count).SubItems(5) = Notify
        
        'Checks if its Allowed to Send Message
        If Notify >= 1 Then
          Exit Sub
        End If
                
        'Phases Data into Normal format for log
        Dim strRTBformat As String
        Dim strNormal As String
        strRTBformat = Replace(varsplit(1), "|ÿ|?|Ö|?|ÿ|", vbCrLf)
        
        rtbphase.TextRTF = strRTBformat
        strNormal = rtbphase.Text
        
        'Checks File Lenght
        If Len(rtbphase.Text) >= 100 Then
           log "< " & strNam & " > Sending has been Declined & Ban 60 min"
           
           lvBanned.ListItems.Add , , lvOnline.ListItems(lngA).Text
           lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(1) = Time
           lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(2) = "SERVER"
           lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(3) = winsock1(Index).RemoteHostIP
           lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(4) = "Excesive Typing"
           lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(5) = "60min"
           lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(6) = "3600"
           winsock1(Index).Close
           winsock1_Close Index
           
           Exit Sub
        End If
        
        'Checks Font SIZE
        If rtbphase.SelFontSize > 14 Then
          winsock1(Index).SendData "Msgbox/Ñ/Your Font is to big" & vbCrLf
          Exit Sub
        End If
        
        log "< " & strNam & " > " & strNormal
        
        'Sends Message to All Users
        For lngA = 1 To lvOnline.ListItems.Count
           DoEvents
           If winsock1(lvOnline.ListItems(lngA).SubItems(1)).State = 7 Then
               If strNam <> "" Then
                  winsock1(lvOnline.ListItems(lngA).SubItems(1)).SendData "ChatRTBText/Ñ/" & strNam & "/Ñ/" & varsplit(1) & vbCrLf
               End If
           End If
        Next lngA
        
      'Creates a New Account
      Case "NewAccount"
        log winsock1(Index).RemoteHostIP & " : has Registered Account"
        For lngA = 1 To lvAccounts.ListItems.Count
            DoEvents
            If LCase(lvAccounts.ListItems(lngA).Text) = LCase(varsplit(1)) Then
               winsock1(Index).SendData "Msgbox/Ñ/User Account Already Exisits Please Select Another Name" & vbCrLf
               Exit Sub
            End If
        Next lngA
           
           'Adds New Name
           lvAccounts.ListItems.Add , , varsplit(1)
           lvAccounts.ListItems(lvAccounts.ListItems.Count).SubItems(1) = varsplit(2)
           lvAccounts.ListItems(lvAccounts.ListItems.Count).SubItems(2) = "0"
           lvAccounts.ListItems(lvAccounts.ListItems.Count).SubItems(3) = winsock1(Index).RemoteHostIP
           lvAccounts.ListItems(lvAccounts.ListItems.Count).SubItems(4) = "Never"
           lvAccounts.ListItems(lvAccounts.ListItems.Count).SubItems(5) = Date
           lvAccounts.ListItems(lvAccounts.ListItems.Count).SubItems(6) = varsplit(3)
           
           'Sends User OK Mess
           winsock1(Index).SendData "Msgbox/Ñ/" & varsplit(1) & " Has been Created and Ready to Use!" & vbCrLf
           
           For lngB = 1 To lvOnline.ListItems.Count
              DoEvents
              If lvOnline.ListItems(lngB).SubItems(5) = "1" Then
                 If winsock1(lvOnline.ListItems(lngB).SubItems(1)).State = 7 Then
                     winsock1(lvOnline.ListItems(lngB).SubItems(1)).SendData "MsgRTB/Ñ/SERVER >> " & varsplit(1) & " Has Registered at " & Time & vbCrLf
                 End If
              End If
           Next lngB
           
      'Removes A Account if proper Password is Entered
      Case "RemoveAcc"
           log winsock1(Index).RemoteHostIP & " : has Deleted Account"
           Dim varsplit2 As Variant
           varsplit2 = Split(varsplit(1), "/Ö/")
           
           For lngA = 1 To lvAccounts.ListItems.Count
              If lvAccounts.ListItems(lngA).Text = varsplit2(0) Then
                 If lvAccounts.ListItems(lngA).SubItems(1) = varsplit2(1) Then
                    
                    'Removes Account and Notifys the User
                    lvAccounts.ListItems.Remove lngA
                    
                    'Sends User OK Mess
                    winsock1(Index).SendData "Msgbox/Ñ/" & varsplit2(0) & " Has been REMOVED!" & vbCrLf
                    Exit Sub
                 End If
              End If
           Next lngA
           'Sends User OK Mess
           winsock1(Index).SendData "Msgbox/Ñ/" & varsplit2(0) & " HAS NOT BEEN REMOVED!" & vbCrLf
           
      'Sends User a List of OPS
      Case "/OPS"
         Dim strOPS
         For lngB = 1 To lvOnline.ListItems.Count
            DoEvents
            If lvOnline.ListItems(lngB).SubItems(3) = 1 Then
                strOPS = strOPS & lvOnline.ListItems(lngB).Text & " - "
            End If
         Next
                 
         winsock1(Index).SendData "Msgbox/Ñ/Currently Logged in OPS: " & strOPS & vbCrLf
         
      'Sends User the Status of the USER DB
      Case "/Stats"
         Dim lngNoOPs As Integer
         lngNoOPs = 0
         
         For lngB = 1 To lvAccounts.ListItems.Count
            DoEvents
            If lvAccounts.ListItems(lngB).SubItems(2) = "1" Then
               lngNoOPs = lngNoOPs + 1
            End If
         Next lngB
         
         winsock1(Index).SendData "Msgbox/Ñ/There are " & lvAccounts.ListItems.Count & " Registered Users. And " & lngNoOPs & " Operators." & vbCrLf
      'This Sends to All Users the Status Change
      Case "StatusChange"
          Dim strNaming As String
          
          For lngA = 1 To lvOnline.ListItems.Count
             DoEvents
             If lvOnline.ListItems(lngA).SubItems(1) = Index Then
                strNameing = lvOnline.ListItems(lngA).Text
                lvOnline.ListItems(lngA).SubItems(6) = varsplit(1)
             End If
          Next lngA
          
          For lngA = 1 To lvOnline.ListItems.Count
             If winsock1(lvOnline.ListItems(lngA).SubItems(1)).State = 7 Then
                winsock1(lvOnline.ListItems(lngA).SubItems(1)).SendData "StatusChange/Ñ/" & strNameing & "/Ñ/" & varsplit(1) & vbCrLf
             End If
          Next lngA
      'Clears all 60 min bans
      Case "Clear60minBan"
          For lngB = 1 To lvOnline.ListItems.Count
              DoEvents
              If lvOnline.ListItems(lngB).SubItems(1) = Index Then
                 If lvOnline.ListItems(lngB).SubItems(3) = 1 Then
                    
                    Dim lngBVB As Long
                    lngBVB = 1
                    Do Until lngBVB > lvBanned.ListItems.Count
                       DoEvents
                       If lvBanned.ListItems(lngBVB).SubItems(5) = "60min" Then
                         lvBanned.ListItems.Remove lngBVB
                       End If
                    Loop
                    Exit Sub
                 End If
              End If
          Next lngB
          
      'This will Tell Ops if a New User Registrates /If Enabled
      Case "NotifyRegistrations"
          For lngB = 1 To lvOnline.ListItems.Count
              DoEvents
              If lvOnline.ListItems(lngB).SubItems(1) = Index Then
                 If lvOnline.ListItems(lngB).SubItems(3) = 1 Then
                    lvOnline.ListItems(lngB).SubItems(5) = 1
                    Exit Sub
                 End If
              End If
          Next lngB
         
      Case "DeNotifyRegistrations"
          For lngB = 1 To lvOnline.ListItems.Count
              DoEvents
              If lvOnline.ListItems(lngB).SubItems(1) = Index Then
                 If lvOnline.ListItems(lngB).SubItems(3) = 1 Then
                    lvOnline.ListItems(lngB).SubItems(5) = 0
                    Exit Sub
                 End If
              End If
          Next lngB
      
      'Sends User a Users Profile
      Case "Profile"
      
          For lngB = 1 To lvAccounts.ListItems.Count
             DoEvents
             If LCase(lvAccounts.ListItems(lngB).Text) = LCase(varsplit(1)) Then
                winsock1(Index).SendData "Profile/Ñ/" & varsplit(1) & "/Ñ/" & lvAccounts.ListItems(lngB).SubItems(6) & vbCrLf
                Exit Sub
             End If
          Next lngB
          
      'Relays the Private Message
      Case "PM"
           Dim Sendersname As String
           
           'Searches for User name that sent the messate
           For lngB = 1 To lvOnline.ListItems.Count
              DoEvents
              If lvOnline.ListItems(lngB).SubItems(1) = Index Then
                 Sendersname = lvOnline.ListItems(lngB).Text
                 GoTo Nexto
              End If
           Next
Nexto:
                            
           For lngB = 1 To lvOnline.ListItems.Count
              DoEvents
              If lvOnline.ListItems(lngB).Text = varsplit(1) Then
                 winsock1(lvOnline.ListItems(lngB).SubItems(1)).SendData "PM/Ñ/" & Sendersname & "/Ñ/" & varsplit(2) & vbCrLf
                 log Sendersname & " Has sent PM"
                 Exit Sub
              End If
           Next lngB
           
      'Muzzles Pep
      Case "Muzzle"
          For lngB = 1 To lvOnline.ListItems.Count
             If lvOnline.ListItems(lngB).SubItems(1) = Index Then
                 If lvOnline.ListItems(lngB).SubItems(3) = 1 Then
                 
                    'Checks if User Is Already Muzzled
                    Dim lngCV As Long
                    For lngCV = 1 To lvMuzzled.ListItems.Count
                       DoEvents
                       If lvMuzzled.ListItems(lngCV).Text = varsplit(1) Then Exit Sub
                    Next lngCV
                    
                    'Adds User to Muzzled List
                    lvMuzzled.ListItems.Add , , varsplit(1)
                    lvMuzzled.ListItems(lvMuzzled.ListItems.Count).SubItems(1) = (Int(varsplit(2) * 60))
                    
                    'Finds Muzzlers Name
                    Dim lngCC As Long
                    For lngCC = 1 To lvOnline.ListItems.Count
                       DoEvents
                       If lvOnline.ListItems(lngCC).SubItems(1) = Index Then
                          Dim strMuzzler As String
                          strMuzzler = lvOnline.ListItems(lngCC).Text
                          GoTo OKKK
                       End If
                    Next lngCC
OKKK:
                    
                    'Sends Muzzled Notifications to All
                    NotifyallMuzzled (Int(varsplit(2) * 60)), strMuzzler, CStr(varsplit(1))
                    Exit Sub
                 End If
              End If
          Next lngB
          
      'Kicks Peps
      Case "Kick"
      
          For lngB = 1 To lvOnline.ListItems.Count
             If lvOnline.ListItems(lngB).SubItems(1) = Index Then
                 If lvOnline.ListItems(lngB).SubItems(3) = "1" Then
                                                        
                    'Finds Kickers Name
                    Dim lngCC2 As Long
                    For lngCC2 = 1 To lvOnline.ListItems.Count
                       DoEvents
                       If lvOnline.ListItems(lngCC2).SubItems(1) = Index Then
                          Dim strKicker As String
                          strKicker = lvOnline.ListItems(lngCC2).Text
                          GoTo OKKK2
                       End If
                    Next lngCC2
OKKK2:
                    'Closes Connection
                    Dim lngBNN As Long
                    For lngBNN = 1 To lvOnline.ListItems.Count
                       DoEvents
                       If lvOnline.ListItems(lngBNN).Text = varsplit(1) Then
                          
                          winsock1(lvOnline.ListItems(lngBNN).SubItems(1)).Close
                          lvOnline.ListItems.Remove Int(lvOnline.ListItems(lngBNN).SubItems(1))
                          GoTo AKOK
                       End If
                    Next
                    
AKOK:
                    'Sends All Users Notifications that X has kicked X
                    Me.Refresh
                    lvOnline.Refresh
                    Dim lngA4 As Long
                    For lngA4 = 1 To lvOnline.ListItems.Count
                        DoEvents
                        If winsock1(lvOnline.ListItems(lngA4).SubItems(1)).State = 7 Then
                            winsock1(Int(lvOnline.ListItems(lngA4).SubItems(1))).SendData "MsgRTB/Ñ/" & varsplit(1) & " Has Been kicked by " & strKicker & vbCrLf & "Left/Ñ/" & varsplit(1) & vbCrLf
                        End If
                    Next lngA4
                       
                          
                    
                    Exit Sub
                 End If
              End If
          Next lngB
    'Starts the User(s) WhiteBoard
    Case "WhiteBoard"
       Dim strRequester As String
       
       'Finds Requesters Name
       For lngB = 1 To lvOnline.ListItems.Count
          DoEvents
          If lvOnline.ListItems(lngB).SubItems(1) = Index Then
             strRequester = lvOnline.ListItems(lngB).Text
             GoTo KOK2
          End If
       Next lngB
       
KOK2:
       'Finds Recivers Index
       For lngB = 1 To lvOnline.ListItems.Count
          DoEvents
          If lvOnline.ListItems(lngB).Text = varsplit(1) Then
             If txtIP.Text <> "" Then
                If InStr(1, winsock1(Index).RemoteHostIP, "192.168" Or "127.0.0.1") Then
                    winsock1(lvOnline.ListItems(lngB).SubItems(1)).SendData "WBRequest/Ñ/" & strRequester & "/Ñ/" & txtIP.Text & "/Ñ/" & varsplit(2) & vbCrLf
                  Else
                    winsock1(lvOnline.ListItems(lngB).SubItems(1)).SendData "WBRequest/Ñ/" & strRequester & "/Ñ/" & winsock1(Index).RemoteHostIP & "/Ñ/" & varsplit(2) & vbCrLf
                End If
               Else
                winsock1(lvOnline.ListItems(lngB).SubItems(1)).SendData "WBRequest/Ñ/" & strRequester & "/Ñ/" & winsock1(Index).RemoteHostIP & "/Ñ/" & varsplit(2) & vbCrLf
             End If
             Exit Sub
          End If
       Next lngB
       
    'Perment Bans
    Case "BanP"
       For lngB = 1 To lvOnline.ListItems.Count
          If lvOnline.ListItems(lngB).SubItems(1) = Index Then
             If lvOnline.ListItems(lngB).SubItems(3) = "1" Then
               
               'Finds Kickers Name
                Dim lngCC3 As Long
                For lngCC3 = 1 To lvOnline.ListItems.Count
                   DoEvents
                   If lvOnline.ListItems(lngCC3).SubItems(1) = Index Then
                      Dim strBanner As String
                      strBanner = lvOnline.ListItems(lngCC3).Text
                      GoTo OKKK4
                   End If
                Next lngCC3
OKKK4:
                'Checks iF a resone was given
                If varsplit(2) = "" Then Exit Sub
                 
                'Bans Users Name
                lvBanned.ListItems.Add , , varsplit(1)
                lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(1) = Time
                lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(2) = strBanner
                lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(3) = winsock1(Index).RemoteHostIP
                lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(4) = varsplit(2)
                lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(5) = "Perment"
                lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(6) = ""
                
                Dim lngA5 As Long
                For lngA5 = 1 To lvOnline.ListItems.Count
                   DoEvents
                   If winsock1(lvOnline.ListItems(lngA5).SubItems(1)).State = 7 Then
                      winsock1(Int(lvOnline.ListItems(lngA5).SubItems(1))).SendData "MsgRTB/Ñ/" & varsplit(1) & " Has Been Banned by " & strBanner & vbCrLf
                   End If
                Next lngA5
                
              End If
           End If
        Next
    'Tells Cam Requestor that it has been declined
    Case "CamDeclined"
       Dim strRequestorsName2 As String
       For lngB = 1 To lvOnline.ListItems.Count
          DoEvents
          If lvOnline.ListItems(lngB).SubItems(1) = Index Then
             strRequestorsName2 = lvOnline.ListItems(lngB).Text
             GoTo NextTStep2
          End If
       Next lngB
NextTStep2:

       For lngB = 1 To lvOnline.ListItems.Count
          DoEvents
          If lvOnline.ListItems(lngB).Text = varsplit(1) Then
              winsock1(lvOnline.ListItems(lngB).SubItems(1)).SendData "CamDeclined/Ñ/" & strRequestorsName2 & vbCrLf
              Exit Sub
          End If
       Next lngB
              
    'Sends Cam Request
    Case "CamRequest"
       Dim strRequestorsName As String
       For lngB = 1 To lvOnline.ListItems.Count
          DoEvents
          If lvOnline.ListItems(lngB).SubItems(1) = Index Then
             strRequestorsName = lvOnline.ListItems(lngB).Text
             GoTo NextTStep
          End If
       Next lngB
NextTStep:
       
       For lngB = 1 To lvOnline.ListItems.Count
          DoEvents
          If varsplit(1) = lvOnline.ListItems(lngB).Text Then
             If winsock1(lvOnline.ListItems(lngB).SubItems(1)).State = 7 Then
                If txtIP.Text <> "" Then
                  If InStr(1, winsock1(Index).RemoteHostIP, "192.168" Or "127.0.0.1") Then
                     winsock1(lvOnline.ListItems(lngB).SubItems(1)).SendData "CamRequest/Ñ/" & strRequestorsName & "/Ñ/" & txtIP.Text & "/Ñ/" & varsplit(2) & vbCrLf
                    Else
                     winsock1(lvOnline.ListItems(lngB).SubItems(1)).SendData "CamRequest/Ñ/" & strRequestorsName & "/Ñ/" & winsock1(Index).RemoteHostIP & "/Ñ/" & varsplit(2) & vbCrLf
                  End If
                 Else
                  winsock1(lvOnline.ListItems(lngB).SubItems(1)).SendData "CamRequest/Ñ/" & strRequestorsName & "/Ñ/" & winsock1(Index).RemoteHostIP & "/Ñ/" & varsplit(2) & vbCrLf
                End If
                
                'Direct Sending Methods(Just for your help)
                'winsock1(lvOnline.ListItems(lngb).SubItems(1)).SendData "CamRequest/Ñ/" & strRequestorsName & "/Ñ/" & "24.158.66.181" & "/Ñ/" & varsplit(2) & vbCrLf
                'winsock1(lvOnline.ListItems(lngb).SubItems(1)).SendData "CamRequest/Ñ/" & strRequestorsName & "/Ñ/" & "127.0.0.1" & "/Ñ/" & varsplit(2) & vbCrLf
                Exit Sub
             End If
             
          End If
       Next lngB
       
    'User Status
    Case "StatusRequest"
       For lngB = 1 To lvOnline.ListItems.Count
             If lvOnline.ListItems(lngB).SubItems(1) = Index Then
                 If lvOnline.ListItems(lngB).SubItems(3) = "1" Then
                   
                   'For Simpler Indexing
                   Dim strUserIdx As String
                   Dim lngUserSel As Long
                   
                   'Finds Users Index & IP
                   Dim lngVB As Long
                   For lngVB = 1 To lvOnline.ListItems.Count
                      DoEvents
                      If lvOnline.ListItems(lngVB).Text = varsplit(1) Then
                          Dim strIP As String
                          strIP = winsock1(lvOnline.ListItems(lngVB).SubItems(3)).RemoteHostIP
                          
                          strUserIdx = lvOnline.ListItems(lngVB).SubItems(3)
                          lngUserSel = lngVB
                      End If
                   Next lngVB
                   
                   'Searches for all accounts w/ same IP & Last Acc login
                   Dim strAllAcc As String
                   Dim strLastlogin As String
                   For lngVB = 1 To lvAccounts.ListItems.Count
                      DoEvents
                      If lvAccounts.ListItems(lngVB).SubItems(3) = strIP Then
                         strAllAcc = strAllAcc & lvAccounts.ListItems(lngVB).Text & ", "
                      End If
                      
                      If lvAccounts.ListItems(lngVB).Text = varsplit(1) Then
                         strLastlogin = lvAccounts.ListItems(lngVB).SubItems(4) & " " & lvAccounts.ListItems(lngVB).SubItems(5)
                      End If
                   Next lngVB
                   
                   'Gets Total amount of Data in chatwindow
                   Dim strTotalData As String
                   strTotalData = lvOnline.ListItems(lngUserSel).SubItems(7)
                   
                   'Sends Databack
                   If winsock1(Index).State = 7 Then
                     winsock1(Index).SendData "UserStatus/Ñ/" & strIP & "/Ñ/" & strAllAcc & "/Ñ/" & strLastlogin & "/Ñ/" & strTotalData & vbCrLf
                   End If
                 End If
              End If
        Next
    'Bans A User for 60 minutes
    Case "Ban60"
       For lngB = 1 To lvOnline.ListItems.Count
          If lvOnline.ListItems(lngB).SubItems(1) = Index Then
             If lvOnline.ListItems(lngB).SubItems(3) = "1" Then
               
               'Finds Kickers Name
                Dim lngCC4 As Long
                For lngCC4 = 1 To lvOnline.ListItems.Count
                   DoEvents
                   If lvOnline.ListItems(lngCC4).SubItems(1) = Index Then
                      Dim strBanner1 As String
                      strBanner1 = lvOnline.ListItems(lngCC4).Text
                      GoTo OKKK3
                   End If
                Next lngCC4
OKKK3:
                'Checks iF a resone was given
                If varsplit(2) = "" Then Exit Sub
                 
                'Bans Users Name
                lvBanned.ListItems.Add , , varsplit(1)
                lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(1) = Time
                lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(2) = strBanner1
                lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(3) = winsock1(Index).RemoteHostIP
                lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(4) = varsplit(2)
                lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(5) = "60min"
                lvBanned.ListItems(lvBanned.ListItems.Count).SubItems(6) = "3600"
                
                Dim lngA6 As Long
                For lngA6 = 1 To lvOnline.ListItems.Count
                   DoEvents
                   If winsock1(lvOnline.ListItems(lngA6).SubItems(1)).State = 7 Then
                      winsock1(Int(lvOnline.ListItems(lngA6).SubItems(1))).SendData "MsgRTB/Ñ/" & varsplit(1) & " Has Been Banned by " & strBanner & vbCrLf
                   End If
                Next lngA6
                
              End If
           End If
        Next
    End Select
    
End Sub

Public Function log(message As String)
  rtbLog.Text = rtbLog.Text & vbCrLf & "[ " & Date & " " & Time & "] " & message
  Print #253, "[ " & Date & " " & Time & "] " & message
End Function

Public Sub NotifyallMuzzled(MuzzledTime As Long, Muzzler As String, MuzzledWho As String)
   'Sends All Users Notifications that X has muzzled X
   For lngA = 1 To lvOnline.ListItems.Count
      DoEvents
      If winsock1(lvOnline.ListItems(lngA).SubItems(1)).State = 7 Then
          winsock1(lvOnline.ListItems(lngA).SubItems(1)).SendData "MsgRTB/Ñ/" & MuzzledWho & " Has Been Muzzled for " & MuzzledTime & ". By " & Muzzler & vbCrLf
      End If
   Next lngA
End Sub

Private Sub winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   winsock1_Close Index
   
End Sub
