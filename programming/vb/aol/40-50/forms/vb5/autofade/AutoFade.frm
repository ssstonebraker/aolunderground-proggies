VERSION 5.00
Object = "{B5089F43-6EDC-101C-B41C-00AA0036005A}#4.0#0"; "DWSBC32.OCX"
Begin VB.Form Fader 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AutoFade"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   2070
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Wavy"
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OFF"
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   1800
      Width           =   495
   End
   Begin DwsbcLib.SubClass SubClass1 
      Left            =   0
      Top             =   0
      _Version        =   262144
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      CtlParam        =   ""
      Persist         =   0
      RegMessage1     =   ""
      RegMessage2     =   ""
      RegMessage3     =   ""
      RegMessage4     =   ""
      RegMessage5     =   ""
      Type            =   0
      Messages        =   "AutoFade.frx":0000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "On"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   495
   End
   Begin VB.HScrollBar HScroll6 
      Height          =   135
      Left            =   960
      Max             =   255
      TabIndex        =   7
      Top             =   1440
      Width           =   975
   End
   Begin VB.HScrollBar HScroll5 
      Height          =   135
      Left            =   960
      Max             =   255
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin VB.HScrollBar HScroll4 
      Height          =   135
      Left            =   960
      Max             =   255
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   135
      Left            =   960
      Max             =   255
      TabIndex        =   3
      Top             =   600
      Value           =   1
      Width           =   975
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   135
      Left            =   960
      Max             =   255
      TabIndex        =   2
      Top             =   360
      Value           =   1
      Width           =   975
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   135
      Left            =   960
      Max             =   255
      TabIndex        =   1
      Top             =   120
      Value           =   1
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   2280
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "Fader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
SubClass1.HwndParam = ChatSendBox
End Sub

Private Sub Command2_Click()
On Error Resume Next
SubClass1.HwndParam = 0
End Sub

Private Sub Form_Load()
StayOnTop Me
Me.Show
End Sub

Private Sub HScroll1_Change()
Picture1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub HScroll1_Scroll()
Picture1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub HScroll2_Change()
Picture1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub HScroll2_Scroll()
Picture1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub HScroll3_Change()
Picture1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub HScroll3_Scroll()
Picture1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub HScroll4_Change()
Picture2.BackColor = RGB(HScroll4.Value, HScroll5.Value, HScroll6.Value)
End Sub

Private Sub HScroll4_Scroll()
Picture2.BackColor = RGB(HScroll4.Value, HScroll5.Value, HScroll6.Value)
End Sub

Private Sub HScroll5_Change()
Picture2.BackColor = RGB(HScroll4.Value, HScroll5.Value, HScroll6.Value)
End Sub

Private Sub HScroll5_Scroll()
Picture2.BackColor = RGB(HScroll4.Value, HScroll5.Value, HScroll6.Value)
End Sub

Private Sub HScroll6_Change()
Picture2.BackColor = RGB(HScroll4.Value, HScroll5.Value, HScroll6.Value)
End Sub

Private Sub HScroll6_Scroll()
Picture2.BackColor = RGB(HScroll4.Value, HScroll5.Value, HScroll6.Value)
End Sub

Private Sub SubClass1_WndMessageX(wnd As Stdole.OLE_HANDLE, msg As Stdole.OLE_HANDLE, wp As Stdole.OLE_HANDLE, lp As Long, retval As Long, nodef As Integer, Process As Stdole.OLE_HANDLE)
Dim wav As Boolean
If wp = 13 Then
If Check1.Value = 1 Then wav = True
If Check1.Value = 0 Then wav = False
said$ = GetText(ChatSendBox)
If InStr(LCase(said$), "<font color=") <> 0 Then Exit Sub
If TrimSpaces(said$) = "" Then Exit Sub
said$ = ChatFade(said$, HScroll1.Value, HScroll4.Value, HScroll2.Value, HScroll5.Value, HScroll3.Value, HScroll6.Value, wav)
SubClass1.HwndParam = 0
Call SetText(ChatSendBox, "")
Call SetText(ChatSendBox, said$)
SubClass1.HwndParam = ChatSendBox
End If
End Sub
