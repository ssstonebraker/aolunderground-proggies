VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Zer0 Effect Fader"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4920
   Icon            =   "aol710fade.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "About"
      Height          =   255
      Left            =   3720
      TabIndex        =   29
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Remove"
      Height          =   255
      Left            =   3720
      TabIndex        =   28
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Height          =   255
      Left            =   3720
      TabIndex        =   27
      Top             =   1320
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   1200
      TabIndex        =   25
      Top             =   1200
      Width           =   2415
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   135
      Left            =   3000
      Max             =   255
      TabIndex        =   21
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   135
      Left            =   1560
      Max             =   255
      TabIndex        =   20
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   135
      Left            =   120
      Max             =   255
      TabIndex        =   19
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4800
      TabIndex        =   18
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1800
      Top             =   1800
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Fader"
      Height          =   255
      Left            =   3720
      TabIndex        =   17
      Top             =   960
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Strikethru"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1920
      Width           =   975
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Italic"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1680
      Width           =   975
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Underline"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1440
      Width           =   975
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Bold"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1200
      Width           =   975
   End
   Begin VB.PictureBox Picture10 
      BackColor       =   &H80000007&
      Height          =   375
      Left            =   4440
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   12
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture9 
      BackColor       =   &H80000007&
      Height          =   375
      Left            =   3960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   11
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H80000007&
      Height          =   375
      Left            =   3480
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   10
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H80000007&
      Height          =   375
      Left            =   3000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   9
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H80000007&
      Height          =   375
      Left            =   2520
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   8
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H80000007&
      Height          =   375
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   7
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H80000007&
      Height          =   375
      Left            =   1560
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000007&
      Height          =   375
      Left            =   1080
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000007&
      Height          =   375
      Left            =   600
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Wavy"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      MaxLength       =   83
      TabIndex        =   0
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Color Scheme:"
      Height          =   255
      Left            =   1200
      TabIndex        =   26
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Blue"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3000
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Green"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   1560
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Red"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check2_Click()
If Check2 = 1 Then
Text2 = "<b>" & Text2
Else
Text2 = ReplaceText(Text2, "<b>", "")
End If
End Sub

Private Sub Check3_Click()
If Check3 = 1 Then
Text2 = "<u>" & Text2
Else
Text2 = ReplaceText(Text2, "<u>", "")
End If
End Sub

Private Sub Check4_Click()
If Check4 = 1 Then
Text2 = "<i>" & Text2
Else
Text2 = ReplaceText(Text2, "<i>", "")
End If
End Sub

Private Sub Check5_Click()
If Check5 = 1 Then
Text2 = "<s>" & Text2
Else
Text2 = ReplaceText(Text2, "<s>", "")
End If
End Sub

Private Sub Command1_Click()
If Text1 = "" Then Exit Sub
If Check1 = 1 Then
Call WavyText(Text1)
Call SendChat(Text2 & "<Font Color=#" & RGBtoHEX(RGB(HScroll3, HScroll2, HScroll1)) & ">" & Text1)
Text1 = ""
ElseIf Check6 = 1 Then
Call SendChat(Text2 & FadeByColor10(Picture1.BackColor, Picture2.BackColor, Picture3.BackColor, Picture4.BackColor, Picture5.BackColor, Picture6.BackColor, Picture7.BackColor, Picture8.BackColor, Picture9.BackColor, Picture10.BackColor, Text1, False))
Text1 = ""
End If
End Sub

Private Sub Command2_Click()
Dim result As String
GoTo ready:
Start:
MsgBox ("Please choose a name")
ready:
result = InputBox("Choose your name", "Save Color Scheme")
If result = "" Then GoTo Start:
List1.AddItem (result)
Call WriteToINI(result, "Fade1", Picture1.BackColor, App.Path & "\color.dat")
Call WriteToINI(result, "Fade2", Picture2.BackColor, App.Path & "\color.dat")
Call WriteToINI(result, "Fade3", Picture3.BackColor, App.Path & "\color.dat")
Call WriteToINI(result, "Fade4", Picture4.BackColor, App.Path & "\color.dat")
Call WriteToINI(result, "Fade5", Picture5.BackColor, App.Path & "\color.dat")
Call WriteToINI(result, "Fade6", Picture6.BackColor, App.Path & "\color.dat")
Call WriteToINI(result, "Fade7", Picture7.BackColor, App.Path & "\color.dat")
Call WriteToINI(result, "Fade8", Picture8.BackColor, App.Path & "\color.dat")
Call WriteToINI(result, "Fade9", Picture9.BackColor, App.Path & "\color.dat")
Call WriteToINI(result, "Fade10", Picture10.BackColor, App.Path & "\color.dat")

Call Save_ListBox(App.Path & "\scheme.lst", List1)
End Sub

Private Sub Command3_Click()
If List1.List(List1.ListIndex) = "" Then
MsgBox ("please select a scheme to remove")
Exit Sub
End If
List1.RemoveItem (List1.ListIndex)
Call Save_ListBox(App.Path & "\scheme.lst", List1)
End Sub

Private Sub Command4_Click()
MsgBox ("It took me a day in total to make this just how i wanted, except the fact it isn't exactly how i or you would want it, the problem is i can't get the wavy text and fader to work at the same time. I will fix this on my next update. please send comments to shaggyze@aol.com...-Shaggy")
End Sub

Private Sub Form_Load()
If direxists(App.Path & "\scheme.lst") = True Then Else GoTo skip:
Call Load_ListBox(App.Path & "\scheme.lst", List1)
Picture1.BackColor = GetFromINI(List1.List(0), "Fade1", App.Path & "\color.dat")
Picture2.BackColor = GetFromINI(List1.List(0), "Fade2", App.Path & "\color.dat")
Picture3.BackColor = GetFromINI(List1.List(0), "Fade3", App.Path & "\color.dat")
Picture4.BackColor = GetFromINI(List1.List(0), "Fade4", App.Path & "\color.dat")
Picture5.BackColor = GetFromINI(List1.List(0), "Fade5", App.Path & "\color.dat")
Picture6.BackColor = GetFromINI(List1.List(0), "Fade6", App.Path & "\color.dat")
Picture7.BackColor = GetFromINI(List1.List(0), "Fade7", App.Path & "\color.dat")
Picture8.BackColor = GetFromINI(List1.List(0), "Fade8", App.Path & "\color.dat")
Picture9.BackColor = GetFromINI(List1.List(0), "Fade9", App.Path & "\color.dat")
Picture10.BackColor = GetFromINI(List1.List(0), "Fade10", App.Path & "\color.dat")
skip:
Call SendChat("<b>" & FadeByColor10(Picture1.BackColor, Picture2.BackColor, Picture3.BackColor, Picture4.BackColor, Picture5.BackColor, Picture6.BackColor, Picture7.BackColor, Picture8.BackColor, Picture9.BackColor, Picture10.BackColor, "¤´° Zer0 Effect Fader °`¤", False))
TimeOut 0.3
Call SendChat("<b>" & FadeByColor10(Picture1.BackColor, Picture2.BackColor, Picture3.BackColor, Picture4.BackColor, Picture5.BackColor, Picture6.BackColor, Picture7.BackColor, Picture8.BackColor, Picture9.BackColor, Picture10.BackColor, "¤´°     For AOL 7.0     °`¤", False))
TimeOut 0.3
Call SendChat("<b><a href=http://darcfx.com/~shaggy>" & FadeByColor10(Picture1.BackColor, Picture2.BackColor, Picture3.BackColor, Picture4.BackColor, Picture5.BackColor, Picture6.BackColor, Picture7.BackColor, Picture8.BackColor, Picture9.BackColor, Picture10.BackColor, "¤´°      Zer0 Effect      °`¤", False) & "</a>")
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call SendChat("<b>" & FadeByColor10(Picture1.BackColor, Picture2.BackColor, Picture3.BackColor, Picture4.BackColor, Picture5.BackColor, Picture6.BackColor, Picture7.BackColor, Picture8.BackColor, Picture9.BackColor, Picture10.BackColor, "¤´° Zer0 Effect Fader °`¤", False))
TimeOut 0.3
Call SendChat("<b>" & FadeByColor10(Picture1.BackColor, Picture2.BackColor, Picture3.BackColor, Picture4.BackColor, Picture5.BackColor, Picture6.BackColor, Picture7.BackColor, Picture8.BackColor, Picture9.BackColor, Picture10.BackColor, "¤´°     For AOL 7.0     °`¤", False))
TimeOut 0.3
Call SendChat("<b><a href=http://darcfx.com/~shaggy>" & FadeByColor10(Picture1.BackColor, Picture2.BackColor, Picture3.BackColor, Picture4.BackColor, Picture5.BackColor, Picture6.BackColor, Picture7.BackColor, Picture8.BackColor, Picture9.BackColor, Picture10.BackColor, "¤´°      Zer0 Effect      °`¤", False) & "</a>")
End Sub

Private Sub HScroll1_Change()
Picture10.BackColor = RGB(HScroll1, HScroll2, HScroll3)
End Sub

Private Sub HScroll1_Scroll()
Picture10.BackColor = RGB(HScroll1, HScroll2, HScroll3)
End Sub

Private Sub HScroll2_Change()
Picture10.BackColor = RGB(HScroll1, HScroll2, HScroll3)
End Sub

Private Sub HScroll2_Scroll()
Picture10.BackColor = RGB(HScroll1, HScroll2, HScroll3)
End Sub

Private Sub HScroll3_Change()
Picture10.BackColor = RGB(HScroll1, HScroll2, HScroll3)
End Sub

Private Sub HScroll3_Scroll()
Picture10.BackColor = RGB(HScroll1, HScroll2, HScroll3)
End Sub

Private Sub List1_Click()
Picture1.BackColor = GetFromINI(List1.List(List1.ListIndex), "Fade1", App.Path & "\color.dat")
Picture2.BackColor = GetFromINI(List1.List(List1.ListIndex), "Fade2", App.Path & "\color.dat")
Picture3.BackColor = GetFromINI(List1.List(List1.ListIndex), "Fade3", App.Path & "\color.dat")
Picture4.BackColor = GetFromINI(List1.List(List1.ListIndex), "Fade4", App.Path & "\color.dat")
Picture5.BackColor = GetFromINI(List1.List(List1.ListIndex), "Fade5", App.Path & "\color.dat")
Picture6.BackColor = GetFromINI(List1.List(List1.ListIndex), "Fade6", App.Path & "\color.dat")
Picture7.BackColor = GetFromINI(List1.List(List1.ListIndex), "Fade7", App.Path & "\color.dat")
Picture8.BackColor = GetFromINI(List1.List(List1.ListIndex), "Fade8", App.Path & "\color.dat")
Picture9.BackColor = GetFromINI(List1.List(List1.ListIndex), "Fade9", App.Path & "\color.dat")
Picture10.BackColor = GetFromINI(List1.List(List1.ListIndex), "Fade10", App.Path & "\color.dat")

End Sub

Private Sub Picture1_Click()
CommonDialog1.ShowColor
Picture1.BackColor = CommonDialog1.Color
End Sub

Private Sub Picture10_Click()
If Check1 = 1 Then Exit Sub
CommonDialog1.ShowColor
Picture10.BackColor = CommonDialog1.Color
End Sub

Private Sub Picture2_Click()
CommonDialog1.ShowColor
Picture2.BackColor = CommonDialog1.Color
End Sub

Private Sub Picture3_Click()
CommonDialog1.ShowColor
Picture3.BackColor = CommonDialog1.Color
End Sub

Private Sub Picture4_Click()
CommonDialog1.ShowColor
Picture4.BackColor = CommonDialog1.Color
End Sub

Private Sub Picture5_Click()
CommonDialog1.ShowColor
Picture5.BackColor = CommonDialog1.Color
End Sub

Private Sub Picture6_Click()
CommonDialog1.ShowColor
Picture6.BackColor = CommonDialog1.Color
End Sub

Private Sub Picture7_Click()
CommonDialog1.ShowColor
Picture7.BackColor = CommonDialog1.Color
End Sub

Private Sub Picture8_Click()
CommonDialog1.ShowColor
Picture8.BackColor = CommonDialog1.Color
End Sub

Private Sub Picture9_Click()
CommonDialog1.ShowColor
Picture9.BackColor = CommonDialog1.Color
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = "13" Then Command1_Click
End Sub

Private Sub Timer1_Timer()
If Check1 = 1 Then
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
HScroll1.Visible = True
HScroll2.Visible = True
HScroll3.Visible = True
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
Picture8.Visible = False
Picture9.Visible = False
Check6 = 0
Else
HScroll1.Visible = False
HScroll2.Visible = False
HScroll3.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Picture1.Visible = True
Picture2.Visible = True
Picture3.Visible = True
Picture4.Visible = True
Picture5.Visible = True
Picture6.Visible = True
Picture7.Visible = True
Picture8.Visible = True
Picture9.Visible = True
Check1 = 0
Check6 = 1
End If
End Sub
