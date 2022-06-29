VERSION 5.00
Object = "{063CC6D8-C221-11D0-AD30-00400516FF78}#1.0#0"; "FLOATBUTTON.OCX"
Begin VB.Form Form1 
   Caption         =   "Fiction's Requester Bot  (For KnK)"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7950
   BeginProperty Font 
      Name            =   "Matura MT Script Capitals"
      Size            =   9.75
      Charset         =   0
      Weight          =   600
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "FORM1.frx":0000
   ScaleHeight     =   4305
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
   Begin FloatButtonControl.FloatButton FloatButton3 
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BackColor       =   -2147483642
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Matura MT Script Capitals"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Exit "
   End
   Begin FloatButtonControl.FloatButton FloatButton2 
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BackColor       =   -2147483647
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Matura MT Script Capitals"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Shut Off Bot"
   End
   Begin FloatButtonControl.FloatButton FloatButton1 
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BackColor       =   -2147483647
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Matura MT Script Capitals"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Start Bot"
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3720
      Top             =   3600
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   1320
      TabIndex        =   2
      Top             =   3360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Text            =   "Can u Send  "
      Top             =   4080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000001&
      ForeColor       =   &H80000005&
      Height          =   735
      Left            =   2520
      TabIndex        =   0
      Text            =   "What u want?"
      Top             =   480
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = True
SendChat "<FONT COLOR=#0033FF FACE=""Arial Narrow""><B><I>Fictions Request Bot"
TimeOut 0.1
SendChat "<FONT COLOR=#0066FF FACE=""Arial Narrow""><B><I>Coded By : <U>Fiction"
TimeOut 0.1
SendChat "<FONT COLOR=#0077FF FACE=""Arial Narrow""><B><I>Requesting : " + Text1.Text + ""
TimeOut 0.1
SendChat "<FONT COLOR=#0099FF FACE=""Arial Narrow""><B><I>type : -i got it    if you have it"
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
SendChat "<FONT COLOR=#0033FF FACE=""Arial Narrow""><B><I>Fictions Magikal Request Bot"
TimeOut 0.1
SendChat "<FONT COLOR=#0066FF FACE=""Arial Narrow""><B><I>Coded By : <U>Fiction"
TimeOut 0.1
SendChat "<FONT COLOR=#0077FF FACE=""Arial Narrow""><B><I>Got What I Was Lookin 4"
TimeOut 0.1
SendChat "<FONT COLOR=#0099FF FACE=""Arial Narrow""><B><I>Bot Is No Longer Functional"
End Sub

Private Sub Command3_Click()
SendChat "<FONT COLOR=#0033FF FACE=""Arial Narrow""><B><I>Fictions Magikal Request Bot"
TimeOut 0.1
SendChat "<FONT COLOR=#0066FF FACE=""Arial Narrow""><B><I>Coded By : Fiction"
TimeOut 0.1
SendChat "<FONT COLOR=#0077FF FACE=""Arial Narrow""><B><I>Fiction Is No Longer Magikal"
End
End Sub

Private Sub FloatButton1_Click()
Timer1.Enabled = True
SendChat "<FONT COLOR=#0033FF FACE=""Arial Narrow""><B><I>Fictions Request Bot"
TimeOut 0.1
SendChat "<FONT COLOR=#0066FF FACE=""Arial Narrow""><B><I>Coded By : <U>Fiction"
TimeOut 0.1
SendChat "<FONT COLOR=#0077FF FACE=""Arial Narrow""><B><I>Requesting : " + Text1.Text + ""
TimeOut 0.1
SendChat "<FONT COLOR=#0099FF FACE=""Arial Narrow""><B><I>type : -i got it    if you have it"
End Sub

Private Sub FloatButton2_Click()
Timer1.Enabled = False
SendChat "<FONT COLOR=#0033FF FACE=""Arial Narrow""><B><I>Fictions Magikal Request Bot"
TimeOut 0.1
SendChat "<FONT COLOR=#0066FF FACE=""Arial Narrow""><B><I>Coded By : <U>Fiction"
TimeOut 0.1
SendChat "<FONT COLOR=#0077FF FACE=""Arial Narrow""><B><I>Got What I Was Lookin 4"
TimeOut 0.1
SendChat "<FONT COLOR=#0099FF FACE=""Arial Narrow""><B><I>Bot Is No Longer Functional"
End Sub

Private Sub FloatButton3_Click()
SendChat "<FONT COLOR=#0033FF FACE=""Arial Narrow""><B><I>Fictions Magikal Request Bot"
TimeOut 0.1
SendChat "<FONT COLOR=#0066FF FACE=""Arial Narrow""><B><I>Coded By : Fiction"
TimeOut 0.1
SendChat "<FONT COLOR=#0077FF FACE=""Arial Narrow""><B><I>Fiction Is No Longer Magikal"
End
End Sub

Private Sub Form_Load()
SendChat "<FONT COLOR=#0033FF FACE=""Arial Narrow""><B><I> Fictions Magikal Request Bot"
TimeOut 0.1
SendChat "<FONT COLOR=#0066FF FACE=""Arial Narrow""><B><I> Coded By : <U>Fiction"
TimeOut 0.1
SendChat "<FONT COLOR=#0077FF FACE=""Arial Narrow""><B><I> Fiction Is Magikal Once Again!"
End Sub

Private Sub Timer1_Timer()
If UCase(LastChatLine) = UCase("-i got it") Then
List1.AddItem SNFromLastChatLine
SendChat "" + SNFromLastChatLine + " Can you send!"
For X = 0 To List1.ListCount - 1
TimeOut 0.1
Call IMKeyword("" + List1.List(X) + "", "Sup " + List1.List(X) + "? " + Text2.Text + "" + Text1.Text + " to me?")
Next X
End If
End Sub


