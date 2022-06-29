VERSION 5.00
Begin VB.Form options_frm 
   BorderStyle     =   0  'None
   Caption         =   "caustik converter"
   ClientHeight    =   3705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2640
      Top             =   3360
   End
   Begin VB.PictureBox detect 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   23
      Top             =   2640
      Visible         =   0   'False
      Width           =   855
      Begin VB.Label chary 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "g"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.TextBox Combo1 
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      Text            =   "Arial"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox timey 
      Height          =   285
      Left            =   1680
      TabIndex        =   9
      Text            =   ".8"
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox fonty 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   11
      Text            =   "|"
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox yy 
      Height          =   285
      Left            =   1680
      TabIndex        =   13
      Text            =   "6"
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox xy 
      Height          =   285
      Left            =   1680
      TabIndex        =   12
      Text            =   "80"
      Top             =   3000
      Width           =   375
   End
   Begin VB.Line Line29 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   1800
      X2              =   2040
      Y1              =   2400
      Y2              =   2160
   End
   Begin VB.Line Line28 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   1800
      X2              =   1680
      Y1              =   2400
      Y2              =   2280
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1680
      TabIndex        =   28
      Top             =   2280
      Width           =   375
   End
   Begin VB.Line Line27 
      BorderColor     =   &H00FFFFFF&
      X1              =   960
      X2              =   1560
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Underline"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   27
      Top             =   2280
      Width           =   675
   End
   Begin VB.Line Line26 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   1800
      X2              =   2040
      Y1              =   2040
      Y2              =   1800
   End
   Begin VB.Line Line25 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   1680
      X2              =   1800
      Y1              =   1920
      Y2              =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   720
      X2              =   1560
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Italics"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   26
      Top             =   1920
      Width           =   405
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1680
      TabIndex        =   25
      Top             =   1920
      Width           =   375
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FFFFFF&
      X1              =   1320
      X2              =   1560
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto - detect"
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
      Left            =   3000
      TabIndex        =   22
      Top             =   3000
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Presets"
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
      TabIndex        =   5
      Top             =   1440
      Width           =   555
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ok"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3360
      TabIndex        =   19
      Top             =   1920
      Width           =   210
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   2880
      TabIndex        =   20
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Line Line24 
      BorderColor     =   &H00FFFFFF&
      X1              =   2880
      X2              =   4080
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line23 
      BorderColor     =   &H00FFFFFF&
      X1              =   4080
      X2              =   4080
      Y1              =   1800
      Y2              =   2280
   End
   Begin VB.Line Line22 
      BorderColor     =   &H00FFFFFF&
      X1              =   2880
      X2              =   2880
      Y1              =   1800
      Y2              =   2280
   End
   Begin VB.Line Line21 
      BorderColor     =   &H00FFFFFF&
      X1              =   2880
      X2              =   4080
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line20 
      BorderColor     =   &H00FFFFFF&
      X1              =   4080
      X2              =   4080
      Y1              =   1320
      Y2              =   1800
   End
   Begin VB.Line Line19 
      BorderColor     =   &H00FFFFFF&
      X1              =   4080
      X2              =   2880
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00FFFFFF&
      X1              =   2880
      X2              =   2880
      Y1              =   1320
      Y2              =   1800
   End
   Begin VB.Line Line17 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   1800
      X2              =   2040
      Y1              =   1680
      Y2              =   1440
   End
   Begin VB.Line Line16 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   1680
      X2              =   1800
      Y1              =   1560
      Y2              =   1680
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1680
      TabIndex        =   18
      Top             =   1560
      Width           =   375
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00FFFFFF&
      X1              =   600
      X2              =   1560
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bold"
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
      TabIndex        =   17
      Top             =   1560
      Width           =   315
   End
   Begin VB.Line Line13 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   1680
      X2              =   1800
      Y1              =   1200
      Y2              =   1320
   End
   Begin VB.Line Line14 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   2040
      X2              =   1800
      Y1              =   1080
      Y2              =   1320
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1680
      TabIndex        =   16
      Top             =   1200
      Width           =   375
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00FFFFFF&
      X1              =   1200
      X2              =   1560
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color blending"
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
      TabIndex        =   15
      Top             =   1200
      Width           =   1020
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FFFFFF&
      X1              =   960
      X2              =   1560
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FFFFFF&
      X1              =   840
      X2              =   1560
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      X1              =   960
      X2              =   1200
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   4560
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Under .7 is dangerous"
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
      Left            =   2640
      TabIndex        =   6
      Top             =   760
      Width           =   1605
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   2160
      X2              =   2520
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Timeout"
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
      TabIndex        =   1
      Top             =   770
      Width           =   555
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Advanced Options"
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
      TabIndex        =   7
      Top             =   2520
      Width           =   1350
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   2160
      X2              =   2760
      Y1              =   3000
      Y2              =   2640
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   2160
      X2              =   2760
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Character"
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
      TabIndex        =   2
      Top             =   2640
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Char y/x Offset"
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
      TabIndex        =   4
      Top             =   3360
      Width           =   1125
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Characters per line"
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
      Top             =   3000
      Width           =   1380
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Font Face"
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
      TabIndex        =   0
      Top             =   405
      Width           =   720
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   4320
      TabIndex        =   14
      Top             =   0
      Width           =   255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   4320
      X2              =   4440
      Y1              =   120
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   4320
      X2              =   4440
      Y1              =   0
      Y2              =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "caustik converter - options"
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
      Left            =   240
      TabIndex        =   8
      Top             =   0
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   3735
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   2880
      TabIndex        =   21
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "options_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Change()
scan_font = Combo1.text
On Error Resume Next: fonty.Font = scan_font: fonty.Refresh
Call Label18_MouseDown(1, 0, 0, 0)
End Sub

Private Sub Combo1_Click()
scan_font = Combo1.text
On Error Resume Next: fonty.Font = scan_font: fonty.Refresh
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
scan_font = Combo1.text
On Error Resume Next: fonty.Font = scan_font: fonty.Refresh
End Sub

Private Sub Combo1_Scroll()
scan_font = Combo1.text
On Error Resume Next: fonty.Font = scan_font: fonty.Refresh
If (LCase(scan_font) = "webdings") Then
fonty.text = "g"
xy.text = 20
yy.text = 10
scan_resx = 40
End If
If (LCase(scan_font) = "arial black") Then
fonty.text = "@"
xy.text = 20
yy.text = 10
scan_resx = 40
End If
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
scan_font = Combo1.text
On Error Resume Next: fonty.Font = scan_font: fonty.Refresh
End Sub

Private Sub fonty_Change()
scan_letter = fonty.text
Call Label18_MouseDown(1, 0, 0, 0)
End Sub

Private Sub fonty_KeyPress(KeyAscii As Integer)
scan_letter = fonty.text
End Sub

Private Sub fonty_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
scan_letter = fonty.text
End Sub

Private Sub Form_Load()
Image2.Picture = mshop_frm.Image1.Picture
Image1.Picture = mshop_frm.Image2.Picture
StayOnTop Me
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Me
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Me
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PopupMenu(Popups.men_presets)
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Showactive Label11
End Sub

Private Sub Label12_Click()
If Line13.Visible = True Then
Line13.Visible = False
Line14.Visible = False
scan_mix = 0
Else
Line13.Visible = True
Line14.Visible = True
scan_mix = 1
End If
End Sub

Private Sub Label14_Click()
If Line16.Visible = True Then
Line16.Visible = False
Line17.Visible = False
scan_bold = 0
Else
Line16.Visible = True
Line17.Visible = True
scan_bold = 1
End If
Call Label18_MouseDown(1, 0, 0, 0)
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Visible = False
End Sub

Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Showactive Label15
End Sub

Private Sub Label16_Click()
Me.Visible = False
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PopupMenu(Popups.men_presets)
End Sub

Public Sub Label18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
chary.Font = Combo1.text
chary.Caption = fonty.text
chary.FontBold = fonty.FontBold
chary.FontItalic = fonty.FontItalic
chary.FontUnderline = fonty.FontUnderline
For X = chary.Width To 0 Step -1
For Y = chary.Height To 0 Step -1
If detect.Point(X, Y) <> RGB(256, 256, 256) Then GoTo found
DoEvents
Next
Next
found:
For Y = chary.Height To 0 Step -1
For X = chary.Width To 0 Step -1
If detect.Point(X, Y) <> RGB(256, 256, 256) Then GoTo found2
DoEvents
Next
Next
found2:
yy.text = (Y / X)
scan_offset = (Y / X)
xy.text = Int(4800 / X)
If xy.text > 80 Then xy.text = 80
scan_resx = Int(4800 / X)
' 4800
If scan_resx > 80 Then scan_resx = 80
End Sub

Private Sub Label18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Showactive Label18
End Sub

Private Sub Label19_Click()
If Line25.Visible = True Then
Line25.Visible = False
Line26.Visible = False
scan_italic = 0
Else
Line25.Visible = True
Line26.Visible = True
scan_italic = 1
End If
Call Label18_MouseDown(1, 0, 0, 0)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Me
End Sub

Private Sub Label22_Click()
If Line28.Visible = True Then
Line28.Visible = False
Line29.Visible = False
scan_underline = 0
Else
Line28.Visible = True
Line29.Visible = True
scan_underline = 1
End If
Call Label18_MouseDown(1, 0, 0, 0)
End Sub

Private Sub Label3_Click()
Me.Visible = False
End Sub

Private Sub Timer1_Timer()
If options_frm.Visible = True Then
If scan_bold = 1 Then fonty.FontBold = True Else fonty.FontBold = False
If scan_italic = 1 Then fonty.FontItalic = True Else fonty.FontItalic = False
If scan_underline = 1 Then fonty.FontUnderline = True Else fonty.FontUnderline = False
End If
End Sub

Private Sub timey_Change()
On Error Resume Next
scan_timeout = timey.text
If Err Then
scan_timeout = 0.8
timey.text = 0.8
End If
End Sub

Private Sub xy_Change()
On Error Resume Next
scan_resx = xy.text
If Err Then
scan_resx = 80
xy.text = 80
End If
End Sub


Private Sub yy_Change()
On Error Resume Next
scan_offset = yy.text
If Err Then
scan_resy = 7
yy.text = 7
End If
End Sub
