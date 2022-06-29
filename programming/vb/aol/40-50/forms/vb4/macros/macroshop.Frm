VERSION 4.00
Begin VB.Form Form15 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Macro Shop"
   ClientHeight    =   3315
   ClientLeft      =   1755
   ClientTop       =   2610
   ClientWidth     =   6435
   Height          =   4005
   Icon            =   "macroshop.frx":0000
   Left            =   1695
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   Top             =   1980
   Width           =   6555
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   6615
      TabIndex        =   3
      Text            =   "true"
      Top             =   1980
      Width           =   330
   End
   Begin VB.TextBox Text2 
      Height          =   1800
      Left            =   6825
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   360
      Width           =   450
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      Columns         =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   825
      ItemData        =   "macroshop.frx":030A
      Left            =   30
      List            =   "macroshop.frx":0470
      Sorted          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   60
      Width           =   6360
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2355
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   930
      Width           =   6375
   End
   Begin MSComDlg.CommonDialog CmDialog1 
      Left            =   4290
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
      DialogTitle     =   "Black Out Macro Files"
      FileName        =   "*.bom"
      Filter          =   "1"
      FontName        =   "Arial"
      FromPage        =   1
      InitDir         =   "c:\blackout\art\"
      Min             =   1
      ToPage          =   9
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu Save_as 
         Caption         =   "Save as"
      End
      Begin VB.Menu open 
         Caption         =   "Open"
      End
      Begin VB.Menu send 
         Caption         =   "Send "
      End
      Begin VB.Menu line 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu clear 
         Caption         =   "Clear"
      End
   End
End
Attribute VB_Name = "Form15"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
CMDialog1.Filter = "ARMeGeDDoN Macro Files (*.agm)|*.agm| Macro Files(*.mdr)|*.mdr|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
CMDialog1.FilterIndex = 1
CMDialog1.Action = 1
Open CMDialog1.filename For Input As 1
text1.Text = Input$(LOF(1), 1)
Close 1

End Sub


Private Sub Command2_Click()
On Error Resume Next
CMDialog1.Filter = "ARMeGeDDoN Macro Files (*.agm)|*.agm| Macro Files(*.mdr)|*.mdr|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
CMDialog1.FilterIndex = 1
CMDialog1.Action = 2
Open CMDialog1.filename For Output As #1
Print #1, text1.Text
Close #1

End Sub


Private Sub Clear_Click()
text1.Text = ""
End Sub

Private Sub Form_Load()
StayOnTop Me
End Sub





































'kaos
Private Sub form_Unload(Cancel As Integer)
Me.Hide
End Sub


Private Sub List1_DblClick()
Index = List1.ListIndex
text1.Text = text1.Text + List1.List(Index)
End Sub


Private Sub noadver_Click()
noadver.Checked = True
onadver.Checked = False
Let text3.Text = "false"
End Sub


Private Sub onadver_Click()
onadver.Checked = True
noadver.Checked = False
Let text3.Text = "True"
End Sub

Private Sub open_Click()
On Error Resume Next
CMDialog1.Filter = "Black Out Macro Files (*.bom)|*.bom| ARMeGeDDoN Macro Files (*.agm)|*.agm| Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
CMDialog1.FilterIndex = 1
CMDialog1.Action = 1
Open CMDialog1.filename For Input As 1
text1.Text = Input$(LOF(1), 1)
Close 1
End Sub

Private Sub Save_as_Click()
On Error Resume Next
CMDialog1.Filter = "Black Out Macro Files (*.bom)|*.bom| ARMeGeDDoN Macro Files (*.agm)|*.agm| Macro Files(*.mdr)|*.mdr|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
CMDialog1.FilterIndex = 1
CMDialog1.Action = 2
Open CMDialog1.filename For Output As #1
Print #1, text1.Text
Close #1

End Sub


Private Sub Send_Click()
If text3.Text = "true" Then Timeout 0.4
                           SendChat "··•··•÷•··•··   BLaCKouT  ··•··•÷•··•··"
                           Timeout 0.4
                           SendChat "··•··•÷•··•··  Macro Shop ··•··•÷•··•··"
                           Timeout 1.4
Let Text2.Text = text1.Text
If Mid(text1, Len(text1), 1) <> Chr$(10) Then
    text1 = text1 + Chr$(13) + Chr$(10)
End If
Do While (InStr(text1, Chr$(13)) <> 0)
    Counter = Counter + 1
    SendChat Mid(text1, 1, InStr(text1, Chr(13)) - 1)
    Timeout 0.4
    If Counter = 4 Then
        Timeout (1.1)
        Counter = 0
    End If
    text1 = Mid(text1, InStr(text1, Chr(13) + Chr(10)) + 2)
Loop
Let text1.Text = Text2.Text

End Sub




