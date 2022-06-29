VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "UpChat"
   ClientHeight    =   255
   ClientLeft      =   5355
   ClientTop       =   255
   ClientWidth     =   3855
   LinkTopic       =   "Form2"
   ScaleHeight     =   255
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000002&
      Caption         =   "Close"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000002&
      Caption         =   "UnUpChat"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000002&
      Caption         =   "UpChat"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(AOL%, 1)
Call EnableWindow(Upp%, 0)
End Sub

Private Sub Command2_Click()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(AOL%, 1)
Call EnableWindow(Upp%, 0)
End Sub

Private Sub Command3_Click()
Unload Form2
Form2.Hide
End Sub

Private Sub Form_Load()
StayOnTop Form2
End Sub
