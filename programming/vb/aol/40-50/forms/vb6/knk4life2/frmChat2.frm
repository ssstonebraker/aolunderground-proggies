VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{72134BA9-52CA-11D2-A11E-24AE06C10000}#2.0#0"; "CHATOCX2.OCX"
Begin VB.Form frmChat 
   Caption         =   "Form1"
   ClientHeight    =   1890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   1890
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB6Chat2.Chat Chat1 
      Left            =   1320
      Top             =   240
      _ExtentX        =   3969
      _ExtentY        =   2170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1440
      Width           =   5175
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   2355
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmChat2.frx":0000
   End
   Begin MSComDlg.CommonDialog cdgColor 
      Left            =   6360
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Chat1_ChatMsg(Screen_Name As String, What_Said As String)
If Screen_Name = GetUser Then
Call DoChatStuff2(Screen_Name, What_Said, False)
End If
If Screen_Name <> GetUser Then
Call DoChatStuff(Screen_Name, What_Said, False)
End If

End Sub

Private Sub Command1_Click()
    ChatSend Text1
End Sub

Private Sub Form_Load()
Chat1.ScanOn
FormOnTop Me
End Sub
