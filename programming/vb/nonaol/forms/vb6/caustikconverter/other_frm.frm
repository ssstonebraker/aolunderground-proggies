VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form other_frm 
   BorderStyle     =   0  'None
   Caption         =   "caustik converter"
   ClientHeight    =   5280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   120
      TabIndex        =   19
      Top             =   240
      Visible         =   0   'False
      Width           =   4335
      Begin VB.HScrollBar blueset 
         Height          =   255
         LargeChange     =   20
         Left            =   1080
         Max             =   255
         TabIndex        =   27
         Top             =   1560
         Width           =   1815
      End
      Begin VB.HScrollBar greenset 
         Height          =   255
         LargeChange     =   20
         Left            =   1080
         Max             =   255
         TabIndex        =   26
         Top             =   1200
         Width           =   1815
      End
      Begin VB.HScrollBar redset 
         Height          =   255
         LargeChange     =   20
         Left            =   1080
         Max             =   255
         TabIndex        =   25
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1080
         TabIndex        =   23
         Text            =   "Arial"
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label22 
         BackColor       =   &H00000000&
         Height          =   735
         Left            =   720
         TabIndex        =   31
         Top             =   960
         Width           =   255
      End
      Begin VB.Line Line25 
         X1              =   3000
         X2              =   3240
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Blue"
         Height          =   195
         Left            =   3360
         TabIndex        =   30
         Top             =   1560
         Width           =   315
      End
      Begin VB.Line Line24 
         X1              =   3000
         X2              =   3240
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Green"
         Height          =   195
         Left            =   3360
         TabIndex        =   29
         Top             =   1200
         Width           =   435
      End
      Begin VB.Line Line23 
         X1              =   3000
         X2              =   3240
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Red"
         Height          =   195
         Left            =   3360
         TabIndex        =   28
         Top             =   840
         Width           =   300
      End
      Begin VB.Line Line22 
         X1              =   960
         X2              =   720
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line21 
         X1              =   960
         X2              =   720
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line20 
         X1              =   720
         X2              =   720
         Y1              =   960
         Y2              =   1680
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   1200
         Width           =   360
      End
      Begin VB.Line Line19 
         X1              =   120
         X2              =   3960
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line18 
         X1              =   720
         X2              =   960
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1680
         TabIndex        =   22
         Top             =   1920
         Width           =   525
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Font"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   21
         Top             =   480
         Width           =   390
      End
      Begin VB.Line Line17 
         X1              =   120
         X2              =   3840
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Drawing Tools"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1200
         TabIndex        =   20
         Top             =   0
         Width           =   1500
      End
   End
   Begin RichTextLib.RichTextBox watch 
      Height          =   2415
      Left            =   360
      TabIndex        =   16
      Top             =   2520
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4260
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      TextRTF         =   $"other_frm.frx":0000
   End
   Begin VB.TextBox sizey2 
      Height          =   285
      Left            =   3480
      TabIndex        =   15
      Text            =   "10"
      Top             =   2080
      Width           =   495
   End
   Begin VB.PictureBox stat 
      Height          =   200
      Left            =   2520
      ScaleHeight     =   9
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   12
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox sizey 
      Height          =   285
      Left            =   2520
      TabIndex        =   7
      Text            =   "2"
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox whoy 
      Height          =   285
      Left            =   720
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Draw"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1920
      TabIndex        =   18
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Send Altered"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   600
      TabIndex        =   17
      Top             =   2160
      Width           =   915
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PtSize"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2880
      TabIndex        =   14
      Top             =   2160
      Width           =   450
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   960
      TabIndex        =   13
      Top             =   1440
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   4440
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preview  conversion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   210
      Left            =   1500
      TabIndex        =   11
      Top             =   1800
      Width           =   1500
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00FFFFFF&
      X1              =   4320
      X2              =   4320
      Y1              =   1320
      Y2              =   1680
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   4440
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scroll"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   960
      TabIndex        =   10
      Top             =   1440
      Width           =   405
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   240
      Y1              =   1320
      Y2              =   1680
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   4440
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Send to top Instant Message"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   210
      Left            =   1250
      TabIndex        =   9
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FFFFFF&
      X1              =   4320
      X2              =   4320
      Y1              =   2040
      Y2              =   2400
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FFFFFF&
      X1              =   4440
      X2              =   120
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   240
      Y1              =   2040
      Y2              =   2400
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1850
      TabIndex        =   8
      Top             =   2160
      Width           =   600
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   4080
      Picture         =   "other_frm.frx":00C9
      Top             =   4800
      Width           =   480
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   4560
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   4440
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   4320
      X2              =   4320
      Y1              =   480
      Y2              =   960
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   240
      Y1              =   480
      Y2              =   960
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Send Email"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3240
      TabIndex        =   0
      Top             =   645
      Width           =   780
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PtSize"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1920
      TabIndex        =   1
      Top             =   650
      Width           =   450
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   360
      TabIndex        =   2
      Top             =   650
      Width           =   225
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   4440
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Send Email"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   210
      Left            =   1800
      TabIndex        =   4
      Top             =   240
      Width           =   780
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   4440
      X2              =   4320
      Y1              =   120
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   4320
      X2              =   4440
      Y1              =   120
      Y2              =   0
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "caustik converter - other"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   1785
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4200
      TabIndex        =   5
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   5100
      Left            =   0
      Stretch         =   -1  'True
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "other_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

End Sub

Private Sub blueset_Change()
Call updateColor
End Sub

Private Sub blueset_Scroll()
Call updateColor
End Sub
Private Sub DELETEME()
On Error Resume Next
Text1.text = watch.SelFontName
If watch.SelLength > 0 Then
    color1 = watch.SelColor
    Label22.BackColor = color1
    Red1 = val("&h" & Mid(color1, 1, 2))
    Green1 = val("&h" & Mid(color1, 3, 2))
    Blue1 = val("&h" & Mid(color1, 5, 2))
    redset.Value = Red1
    greenset.Value = Green1
    blueset.Value = Blue1
End If

End Sub

Private Sub updateColor()
Label22.BackColor = RGB(redset.Value, greenset.Value, blueset.Value)
DoEvents
watch.SelColor = Label22.BackColor
End Sub

Private Sub Form_Load()
Image2.Picture = mshop_frm.Image1.Picture
Image1.Picture = mshop_frm.Image2.Picture
StayOnTop Me
End Sub

Private Sub Form_Resize()
If other_frm.Width < 4575 Then other_frm.Width = 4575
If other_frm.Height < 4275 Then other_frm.Height = 4275
Label14.Top = other_frm.Height - (5280 - 4920)
Line7.X2 = other_frm.Width
Image1.Height = other_frm.Height
Image1.Width = other_frm.Width
Image2.Width = other_frm.Width
Image3.Top = other_frm.Height - 435
Image3.Left = other_frm.Width - 495
watch.Height = other_frm.Height - 2985
watch.Width = other_frm.Width - 480
Line1.X1 = other_frm.Width - 255
Line2.X2 = other_frm.Width - 255
Line1.X2 = other_frm.Width - 135
Line2.X1 = other_frm.Width - 135
Label1.Left = other_frm.Width - 375
End Sub

Private Sub HScroll1_Change()

End Sub

Private Sub greenset_Change()
Call updateColor
End Sub

Private Sub greenset_Scroll()
Call updateColor
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Me
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Resize Me
End Sub

Private Sub Label1_Click()
Me.Visible = False
End Sub



Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
scanning = 0
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Showactive Label11
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
AOLChatSend2 ("<font face=" + Chr(34) + "arial" + Chr(34) + "color=" + Chr(34) + "#000080" + Chr(34) + ">•| caustik converter incoming macro |•")
Pause scan_timeout
AOLChatSend2 ("<font face=" + Chr(34) + "arial" + Chr(34) + "color=" + Chr(34) + "#000080" + Chr(34) + ">•| macro name - " & macro_name & " | Char - " & "<font face=" & Chr(34) & scan_font & Chr(34) & ">" & scan_letter & "<font face=" + Chr(34) + "arial" + Chr(34) + ">" + " |•")
Pause scan_timeout * 2
watch.SelLength = 1
watch.SelStart = 0
texty = watch.text
dafont = watch.SelFontName
lastfont = dafont
dasize = watch.SelFontSize
liney$ = "<font face=" & Chr(34) & dafont & Chr(34) & " ptsize=" & Chr(34) & dasize & Chr(34) & ">"
For v = 0 To Len(watch.text)
watch.SelStart = v
dachar$ = Mid$(texty, v, 1)
dacolor$ = VbtoAol(watch.SelColor)
dafont = watch.SelFontName
If dafont <> lastfont Then
liney$ = liney$ & "<font color=" & Chr(34) & "#" & dacolor$ & Chr(34) & " face=" & Chr(34) & dafont & Chr(34) & ">" & dachar$
Else
liney$ = liney$ & "<font color=" & Chr(34) & "#" & dacolor$ & Chr(34) & ">" & dachar$
End If
lastfont = dafont
If Asc(dachar$) = 10 Then
    AOLChatSend liney$
    liney$ = "<font face=" & Chr(34) & dafont & Chr(34) & " ptsize=" & Chr(34) & dasize & Chr(34) & ">"
    Pause scan_timeout
End If
Next
Pause scan_timeout * 2
AOLChatSend2 ("<font face=" + Chr(34) + "arial" + Chr(34) + "color=" + Chr(34) + "#000080" + Chr(34) + ">•| caustik converter macro complete |•")
End Sub

Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Showactive Label13
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Frame1.Visible = True
Text1.text = watch.SelFontName
End Sub

Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Showactive Label14
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Frame1.Visible = False
End Sub

Private Sub Label17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Showactive Label17
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
watchedit = 0
mshop_frmleft = mshop_frm.Left
mshop_frmtop = mshop_frm.Top
mshop_frm.WindowState = 0
mshop_frm.Left = mshop_frmleft
mshop_frm.Top = mshop_frmtop
Me.Enabled = False
mshop_frm.scanme.AutoRedraw = True
mshop_frm.Enabled = False
status_frm.Show (0)
On Error Resume Next
watch.TextRTF = "Macro is being rendered..."
watch.Refresh
DoEvents
watch.SelStart = 0
watch.SelLength = Len(watch.text)
DoEvents
watch.SelFontSize = sizey2
If scan_bold Then watch.SelBold = True Else watch.SelBold = False
If scan_italic Then watch.SelItalic = True Else watch.SelItalic = False
If scan_underline Then watch.SelUnderline = True Else watch.SelUnderline = False
watch.text = ""
mshop_frm.Label4.Visible = False
mshop_frm.Label3.Visible = False
scanning = 1
found = 0
mshop_frm.Label5.Visible = True
If old_x > new_x Then
temp = old_x
old_x = new_x
new_x = temp
End If
If old_y > new_y Then
temp = old_y
old_y = new_y
new_y = temp
End If
step_x = (new_x - old_x) / (scan_resx)
status_frm.stat1.Tag = (100 / ((new_y - old_y) / (step_x * scan_offset)))
status_frm.stat1.Cls
status_frm.stat1.Line (0, 0)-(status_frm.stat1.Tag, 13), 0, BF
status_frm.stat2.Cls
status_frm.stat2.Tag = 0
lasty = ""

' ---
other_frm.watch.Font = scan_font
other_frm.watch.text = ""
DoEvents
For Y2 = old_y To new_y Step (step_x * scan_offset)
status_frm.stat1.Tag = status_frm.stat1.Tag + (100 / ((new_y - old_y) / (step_x * scan_offset)))
status_frm.stat1.Cls
status_frm.stat1.Line (0, 0)-(status_frm.stat1.Tag, 13), 0, BF
num = 0
For X2 = old_x To new_x Step step_x
num = num + 1
DoEvents
' ? ^
Next
other_frm.watch.text = other_frm.watch.text + String(num, scan_letter)

DoEvents
other_frm.watch.text = other_frm.watch.text + Chr(10)
Next
spot = 0
' ---
status_frm.stat2.Tag = (100 / ((new_y - old_y) / (step_x * scan_offset)))
status_frm.stat2.Cls
status_frm.stat2.Line (0, 0)-(status_frm.stat2.Tag, 13), 0, BF
For Y2 = old_y To new_y Step step_x * scan_offset
status_frm.stat2.Tag = status_frm.stat2.Tag + (100 / ((new_y - old_y) / (step_x * scan_offset)))
status_frm.stat2.Cls
status_frm.stat2.Line (0, 0)-(status_frm.stat2.Tag, 13), 0, BF

For X2 = old_x To new_x Step step_x
If scanning = 0 Then
scan_active = 1
mshop_frm.Label5.Visible = False
mshop_frm.Label3.Visible = True
mshop_frm.Label4.Visible = True
mshop_frm.Label8.Visible = True
watch.Visible = True
Me.Enabled = True
mshop_frm.scanme.AutoRedraw = False
mshop_frm.Enabled = True
watchedit = 1
Exit Sub
End If
If (VbtoAol(mshop_frm.scanme.Point(X2 + (step_x / 2), Y2 + (step_y / 2)))) <> "FEFEFE" And found = 0 Then
found = 1
End If

If scan_mix = 1 Then
' ---
other_frm.watch.SelStart = spot
other_frm.watch.SelLength = (1 * Len(scan_letter))
other_frm.watch.SelColor = cmix2(VbtoAol(mshop_frm.scanme.Point(X2 - step_x, Y2 - step_y)), VbtoAol(mshop_frm.scanme.Point(X2 + step_x, Y2 + step_y)))
spot = spot + 1
DoEvents
' ---
Else

' ---
other_frm.watch.SelStart = spot
other_frm.watch.SelLength = 1
If (X2 + step_x) > mshop_frm.scanme.Width Or (Y2 + step_y) > mshop_frm.scanme.Height Then
    other_frm.watch.SelColor = RGB(255, 255, 255)
    Else
    other_frm.watch.SelColor = cmix2(VbtoAol(mshop_frm.scanme.Point(X2, Y2)), VbtoAol(mshop_frm.scanme.Point(X2, Y2)))
End If
spot = spot + 1
DoEvents
' ---

End If

Next
spot = spot + 1
Next
scanning = 0
scan_active = 1
watch.SelStart = 0
watch.SelLength = Len(watch.text)
DoEvents
watch.SelFontSize = sizey2
If scan_bold Then watch.SelBold = True Else watch.SelBold = False
If scan_italic Then watch.SelItalic = True Else watch.SelItalic = False
If scan_underline Then watch.SelUnderline = True Else watch.SelUnderline = False

'watch.SelFontSize = sizey2
'If scan_bold Then watch.SelBold = True Else watch.SelBold = False
watch.SelLength = 0
mshop_frm.Label3.Visible = True
mshop_frm.Label4.Visible = True
mshop_frm.Label8.Visible = True
mshop_frm.Label5.Visible = False
status_frm.Visible = False
Me.Enabled = True
mshop_frm.scanme.AutoRedraw = False
mshop_frm.Enabled = True
watchedit = 1
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Showactive Label2
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Me
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If whoy.text = "" Then Exit Sub
mshop_frmtop = mshop_frm.Top
mshop_frmwidth = mshop_frm.Width
mshop_frm.WindowState = 0
mshop_frm.Top = mshop_frmtop
mshop_frm.Width = mshop_frmwidth
On Error Resume Next
scanning = 1
status_frm.Show (0)
mshop_frm.scanme.AutoRedraw = True
mshop_frm.Enabled = False
If old_x > new_x Then
temp = old_x
old_x = new_x
new_x = temp
End If
If old_y > new_y Then
temp = old_y
old_y = new_y
new_y = temp
End If

step_x = (new_x - old_x) / (scan_resx)
status_frm.stat1.Tag = (100 / ((new_y - old_y) / (step_x * scan_offset)))
status_frm.stat1.Cls
status_frm.stat1.Line (0, 0)-(status_frm.stat1.Tag, 13), 0, BF
status_frm.stat2.Cls
status_frm.stat2.Tag = 0
For Y2 = old_y To new_y Step step_x * scan_offset
status_frm.stat1.Tag = status_frm.stat1.Tag + (100 / ((new_y - old_y) / (step_x * scan_offset)))
status_frm.stat1.Cls
status_frm.stat1.Line (0, 0)-(status_frm.stat1.Tag, 13), 0, BF
DoEvents
For X2 = old_x To new_x Step step_x
DoEvents
If scan_mix = 1 Then
holdy$ = holdy$ & "<font ptsize=" & sizey.text & " color=#" & cmix(VbtoAol(mshop_frm.scanme.Point(X2 + (step_x / 2), Y2 + (step_y / 2))), VbtoAol(mshop_frm.scanme.Point(X2 - (step_x / 2), Y2 - (step_y / 2)))) & ">" & scan_letter
Else
holdy$ = holdy$ & "<font ptsize=" & sizey.text & " color=#" & VbtoAol(mshop_frm.scanme.Point(X2 + (step_x / 2), Y2 + (step_y / 2))) & ">" & scan_letter
End If
Next
holdy$ = holdy$ & Chr(13) & Chr(10)
Next
mshop_frm.scanme.AutoRedraw = False
mshop_frm.Enabled = True
status_frm.Visible = False
    If scan_bold Then addy$ = addy$ + "<b>"
    If scan_italic Then addy$ = addy$ + "<i>"
    If scan_underline Then addy$ = addy$ + "<u>"
Call SendEmail(whoy.text, "caustik converter - " & macro_name, "<font ptsize=" & sizey.text & " face=" & scan_font & ">" & addy$ & holdy$)
scanning = 0
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Showactive Label7
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
mshop_frmtop = mshop_frm.Top
mshop_frmwidth = mshop_frm.Width
mshop_frm.WindowState = 0
mshop_frm.Top = mshop_frmtop
mshop_frm.Width = mshop_frmwidth
On Error Resume Next
scanning = 1
mshop_frm.scanme.AutoRedraw = True
mshop_frm.Enabled = False
Label9.Visible = False
Label11.Visible = True
found = 0
If old_x > new_x Then
temp = old_x
old_x = new_x
new_x = temp
End If
If old_y > new_y Then
temp = old_y
old_y = new_y
new_y = temp
End If
mshop_frm.scanme.Refresh
step_x = (new_x - old_x) / (scan_resx)

stat.Tag = (100 / ((new_y - old_y) / (step_x * scan_offset)))
stat.Cls
stat.Line (0, 0)-(stat.Tag, 9), 0, BF
lasty = ""

For Y2 = old_y To new_y Step step_x * scan_offset
stat.Tag = stat.Tag + (100 / ((new_y - old_y) / (step_x * scan_offset)))
stat.Cls
stat.Line (0, 0)-(stat.Tag, 9), 0, BF
For X2 = old_x To new_x Step step_x
If scanning = 0 Then
scan_active = 1
Label9.Visible = True
Label11.Visible = False
mshop_frm.scanme.AutoRedraw = False
mshop_frm.Enabled = True
Exit Sub
End If
If (VbtoAol(mshop_frm.scanme.Point(X2 + (step_x / 2), Y2 + (step_y / 2)))) <> "FEFEFE" And found = 0 Then
found = 1
End If
'word$ = word$ + "<font color=#" & VbtoAol(scanme.Point(X2 + (step_x / 2), Y2 + (step_y / 2))) & ">" + scan_letter
'lasty = VbtoAol(scanme.Point(X2 + (step_x / 2), Y2 + (step_y / 2)))
If scan_mix = 1 Then
word$ = word$ + "<font color=#" & cmix(VbtoAol(mshop_frm.scanme.Point(X2 - step_x, Y2 - step_y)), VbtoAol(mshop_frm.scanme.Point(X2 + step_x, Y2 + step_y))) & ">" + scan_letter

lasty = cmix(VbtoAol(mshop_frm.scanme.Point(X2 - step_x, Y2 - step_y)), VbtoAol(mshop_frm.scanme.Point(X2 + step_x, Y2 + step_y)))
Else

word$ = word$ + "<font color=#" & VbtoAol(mshop_frm.scanme.Point(X2, Y2)) & ">" + scan_letter
lasty = VbtoAol(mshop_frm.scanme.Point(X2, Y2))
End If
Next
If found = 1 Then
    If scan_bold Then
        DoEvents
        InstantMessageBody "<font face=" + Chr(34) + scan_font + Chr(34) + "><b>" + word$
        ClickInstantMessageSend
        Else
        DoEvents
        InstantMessageBody "<font face=" + Chr(34) + scan_font + Chr(34) + ">" + word$
        ClickInstantMessageSend
    End If

Pause scan_timeout
End If
spot = spot + 1
word$ = ""
Next
scanning = 0
scan_active = 1
Label9.Visible = True
Label11.Visible = False
mshop_frm.scanme.AutoRedraw = False
mshop_frm.Enabled = True
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Showactive Label9
End Sub

Private Sub redset_Change()
Call updateColor
End Sub

Private Sub redset_Scroll()
Call updateColor
End Sub

Private Sub watch_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 37 Or KeyCode = 10 Or KeyCode = 13 Or watchedit = 0 Or KeyCode = 38 Or KeyCode = 39 Or KeyCode = 40 Or KeyCode = 46 Or KeyCode = 8 Or Shift = 2 Or lastlen = Len(watch.text) Then Exit Sub
lastlen = Len(watch.text)
If watchedit = 1 Then
watch.SelStart = watch.SelStart - 1
watch.SelLength = 1
watch.SelColor = VbtoAol(Label22.BackColor)
watch.SelFontName = Text1.text
watch.Refresh
watch.SelStart = watch.SelStart + 1
watch.SelLength = 0
End If

End Sub
