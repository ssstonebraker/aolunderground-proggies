VERSION 5.00
Begin VB.Form MailOp 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2808
   ClientLeft      =   3552
   ClientTop       =   1116
   ClientWidth     =   2112
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "MailOp.frx":0000
   ScaleHeight     =   2808
   ScaleWidth      =   2112
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      ItemData        =   "MailOp.frx":460C
      Left            =   840
      List            =   "MailOp.frx":4613
      TabIndex        =   5
      Top             =   2160
      Width           =   492
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00000080&
      Caption         =   "Flash Mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   600
      TabIndex        =   4
      Top             =   1680
      Width           =   1212
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00000080&
      Caption         =   "Sent Mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   600
      TabIndex        =   3
      Top             =   1320
      Width           =   1092
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00000080&
      Caption         =   "Old Mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Width           =   1092
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000080&
      Caption         =   "New Mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   1092
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000080&
      Caption         =   "Mail Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2412
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1692
   End
End
Attribute VB_Name = "MailOp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Create.List1.ListCount < 0 Then
DoEvents
List1.Clear
MailOpenNew
Do Until Create.List1.ListCount > 0
DoEvents
Call MailToListNew(Create.List1)
Loop
Else
DoEvents
MailOpenNew
DoEvents
Call MailToListNew(Create.List1)
End If
End Sub

Private Sub Check2_Click()
If Create.List1.ListCount < 0 Then
DoEvents
List1.Clear
MailOpenOld
Do Until Create.List1.ListCount > 0
DoEvents
Call MailToListOld(Create.List1)
Loop
Else
DoEvents
MailOpenOld
DoEvents
Call MailToListOld(Create.List1)
End If
End Sub

Private Sub Check3_Click()
If Create.List1.ListCount < 0 Then
DoEvents
List1.Clear
MailOpenSent
Do Until Create.List1.ListCount > 0
DoEvents
Call MailToListSent(Create.List1)
Loop
Else
DoEvents
MailOpenSent
DoEvents
Call MailToListSent(Create.List1)
End If
End Sub

Private Sub Check4_Click()
If Create.List1.ListCount < 0 Then
DoEvents
List1.Clear
MailOpenFlash
Do Until Create.List1.ListCount > 0
DoEvents
Call MailToListFlash(Create.List1)
Loop
Else
DoEvents
MailOpenFlash
DoEvents
Call MailToListFlash(Create.List1)
End If
End Sub

Private Sub Form_Load()
FormOnTop Me
End Sub

Private Sub List1_Click()
MailOp.Hide
End Sub
