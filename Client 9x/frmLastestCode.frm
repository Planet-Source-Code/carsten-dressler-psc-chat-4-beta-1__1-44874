VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmLastestCode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PSCode.com's Lastest"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3240
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLastestCode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   3240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh Code"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3255
   End
   Begin VB.TextBox txtURL 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Text            =   "Navigate Press Enter"
      Top             =   240
      Width           =   3255
   End
   Begin SHDocVwCtl.WebBrowser wbLastestCode 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   3255
      ExtentX         =   5741
      ExtentY         =   7858
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmLastestCode"
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

Private Sub cmdRefresh_Click()
   'Nevigates IE to PSCode.com
   wbLastestCode.Navigate "http://www.Planet-Source-Code.com/vb/linktous/ScrollingCode.asp?lngWId=1"
End Sub

Private Sub Form_Load()
   'Nevigates IE to PSCode.com
   wbLastestCode.Navigate "http://www.Planet-Source-Code.com/vb/linktous/ScrollingCode.asp?lngWId=1"
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   cmdRefresh.Width = Me.ScaleWidth - 20
   Text1.Width = Me.ScaleWidth - 20
   wbLastestCode.Height = Me.ScaleHeight - 550
   wbLastestCode.Width = Me.ScaleWidth
End Sub

Private Sub txtURL_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     KeyAscii = 0
     On Error Resume Next
     w.Navigate Text1.Text
   End If
End Sub
