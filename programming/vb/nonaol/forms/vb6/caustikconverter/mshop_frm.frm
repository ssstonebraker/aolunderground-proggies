VERSION 5.00
Begin VB.Form mshop_frm 
   BorderStyle     =   0  'None
   Caption         =   "caustik converter"
   ClientHeight    =   3120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4620
   Icon            =   "mshop_frm.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "mshop_frm.frx":030A
   ScaleHeight     =   3120
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox stat 
      BackColor       =   &H00808080&
      Height          =   255
      Left            =   1320
      ScaleHeight     =   13
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   9
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Timer info 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   4080
      Top             =   2400
   End
   Begin VB.Timer intro 
      Interval        =   50
      Left            =   4080
      Top             =   2640
   End
   Begin VB.PictureBox scanme 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1785
      ScaleWidth      =   4305
      TabIndex        =   2
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3360
      TabIndex        =   13
      Top             =   2400
      Width           =   330
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   0
      Picture         =   "mshop_frm.frx":182A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   255
   End
   Begin VB.Label mname 
      BackStyle       =   0  'Transparent
      Caption         =   "(open a file)"
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
      TabIndex        =   12
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3840
      TabIndex        =   11
      Top             =   0
      Width           =   375
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   3960
      X2              =   4200
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "about"
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
      TabIndex        =   10
      Top             =   0
      Width           =   405
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   360
      X2              =   4680
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label warn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "select region!"
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
      Top             =   2760
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   1920
      X2              =   4440
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   4320
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Other"
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
      TabIndex        =   7
      Top             =   2400
      Width           =   405
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
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
      Left            =   1800
      TabIndex        =   6
      Top             =   2400
      Width           =   555
   End
   Begin VB.Label Label5 
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
      Left            =   1080
      TabIndex        =   5
      Top             =   2400
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label Label4 
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
      Left            =   1080
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   4600
      X2              =   4420
      Y1              =   180
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   4600
      X2              =   4420
      Y1              =   0
      Y2              =   180
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Open"
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
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   390
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4320
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "• caustik converter •"
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
      Left            =   1680
      TabIndex        =   0
      Top             =   0
      Width           =   1470
   End
   Begin VB.Image Image1 
      Height          =   210
      Left            =   0
      Picture         =   "mshop_frm.frx":1B34
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4680
   End
   Begin VB.Image Image2 
      Height          =   3135
      Left            =   0
      Picture         =   "mshop_frm.frx":3106
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4620
   End
End
Attribute VB_Name = "mshop_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
StayOnTop Me
filename = ""
watchedit = 1
scan_mode = 1
scan_active = 1
scan_active2 = 0
old_x = 0
old_y = 0
macro_name = "[unknown]"

' Access Data File
On Error Resume Next
Open App.Path + "\data.dat" For Binary As #1
Get #1, 1, currentsettings
Close #1
' Access Data File
' Setup Data
If currentsettings.imagedir = "" Then
currentsettings.imagedir = Mid(App.Path, 1, 3)
currentsettings.imagefile = "NONE"
currentsettings.imagedrive = "NONE"
End If
Load file_frm
If currentsettings.cscan_font = "" Then
currentsettings.cscan_font = "Arial"
currentsettings.cscan_letter = "|"
currentsettings.cscan_bold = 0
currentsettings.cscan_italics = 0
currentsettings.cscan_underline = 0
currentsettings.cscan_mix = 0
currentsettings.cscan_offset = 6
currentsettings.cscan_resx = 80
currentsettings.cscan_timeout = 0.8
End If
scan_font = currentsettings.cscan_font
scan_letter = currentsettings.cscan_letter
options_frm.fonty.text = scan_letter
options_frm.fonty.Font = scan_font
options_frm.Combo1.text = scan_font
scan_bold = currentsettings.cscan_bold
scan_italic = currentsettings.cscan_italics
scan_underline = currentsettings.cscan_underline
scan_mix = currentsettings.cscan_mix
scan_offset = currentsettings.cscan_offset
scan_resx = currentsettings.cscan_resx
scan_timeout = currentsettings.cscan_timeout
If scan_mix > 0 Then
    options_frm.Line13.Visible = True
    options_frm.Line14.Visible = True
End If
If scan_bold > 0 Then
    options_frm.Line16.Visible = True
    options_frm.Line17.Visible = True
End If
If scan_italic > 0 Then
    options_frm.Line25.Visible = True
    options_frm.Line26.Visible = True
End If
If scan_underline > 0 Then
    options_frm.Line28.Visible = True
    options_frm.Line29.Visible = True
End If
options_frm.timey = scan_timeout
options_frm.xy = scan_resx
options_frm.yy = scan_offset

' Setup Data

AOLChatSend2 ("<font face=" + Chr(34) + "arial" + Chr(34) + "color=" + Chr(34) + "#000080" + Chr(34) + ">•| caustik converter loaded | •")
Pause 0.1
AOLChatSend2 ("<font face=" + Chr(34) + "arial" + Chr(34) + "color=" + Chr(34) + "#000080" + Chr(34) + ">•| http://come.to/caustik    | •")
End Sub

Private Sub Form_Resize()
On Error Resume Next
If mshop_frm.Width < 4620 Then mshop_frm.Width = 4620
If mshop_frm.Height < 3120 Then mshop_frm.Height = 3120
If mshop_frm.Width <> 4620 Then mshop_frm.Width = 4620

Line6.X2 = mshop_frm.Width - 420
Line6.X1 = Line6.X2 - 240
Image2.Width = mshop_frm.Width
Image2.Height = mshop_frm.Height
scanme.Width = mshop_frm.Width - 285
scanme.Height = mshop_frm.Height - 1320
stat.Top = mshop_frm.Height - 360
Image1.Width = mshop_frm.Width
Line1.X1 = mshop_frm.Width - 20
Line1.X2 = mshop_frm.Width - 200
Line2.X1 = mshop_frm.Width - 20
Line2.X2 = mshop_frm.Width - 200
Label8.Top = mshop_frm.Height - 720
Label3.Top = mshop_frm.Height - 720
Label4.Top = mshop_frm.Height - 720
Label5.Top = mshop_frm.Height - 720
Label7.Top = mshop_frm.Height - 720
Label10.Top = mshop_frm.Height - 720
Label2.Left = mshop_frm.Width - 200
Line3.Y1 = mshop_frm.Height - 480
Line3.Y2 = mshop_frm.Height - 480
Line3.X2 = mshop_frm.Width - 180
Line4.X2 = mshop_frm.Width - 180
warn.Top = mshop_frm.Height - 360
Line5.X2 = mshop_frm.Width
Label11.Top = Me.Height - (3120 - 2760)
Label11.Left = Me.Width - (4620 - 3360)
old_x = 0
old_y = 0
new_x = 0
new_y = 0
Call mshop_frm.autoselect
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Me
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PopupMenu(Popups.men_about)
End Sub

Private Sub info_Timer()
If info_times > 10 Then
info_times = 0
warn.Visible = False
info.Enabled = False
Exit Sub
End If
If warn.Visible = False Then warn.Visible = True Else warn.Visible = False
info_times = info_times + 1
End Sub

Private Sub intro_Timer()
Label1.Left = Label1.Left - 50
If Label1.Left < 361 Then
Label1.Left = 360
intro.Enabled = False
End If
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Me
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
help_frm.helpmovie.Base = App.Path
help_frm.helpmovie.Movie = App.Path + "\help.swf"
help_frm.helpmovie.BackgroundColor = RGB(0, 0, 0)
help_frm.helpmovie.BGColor = RGB(0, 0, 0)
help_frm.Visible = True
help_frm.helpmovie.GotoFrame (1)
help_frm.helpmovie.Play
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Showactive Label10
End Sub

Private Sub Label2_Click()
AOLChatSend2 ("<font face=" + Chr(34) + "arial" + Chr(34) + "color=" + Chr(34) + "#000080" + Chr(34) + ">•| caustik converter closed |•")
Pause 0.1
AOLChatSend2 ("<font face=" + Chr(34) + "arial" + Chr(34) + "color=" + Chr(34) + "#000080" + Chr(34) + ">•| http://come.to/caustik    | •")
' Access Data File
On Error Resume Next
currentsettings.cscan_font = scan_font
currentsettings.cscan_letter = scan_letter
currentsettings.cscan_bold = scan_bold
currentsettings.cscan_italics = scan_italic
currentsettings.cscan_underline = scan_underline
currentsettings.cscan_mix = scan_mix
currentsettings.cscan_offset = scan_offset
currentsettings.cscan_resx = scan_resx
currentsettings.cscan_timeout = scan_timeout
Open App.Path + "\data.dat" For Binary As #1
Put #1, 1, currentsettings
Close #1
' Access Data File
End
End Sub

Private Sub Label3_Click()
Label4.Visible = False
file_frm.Visible = True
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Showactive Label3
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
mshop_frmleft = mshop_frm.Left
mshop_frmtop = mshop_frm.Top
mshop_frm.WindowState = 0
mshop_frm.Left = mshop_frmleft
mshop_frm.Top = mshop_frmtop
On Error Resume Next
scanning = 1
mshop_frm.scanme.AutoRedraw = True
Label9.Enabled = False
Label4.Visible = False
Label3.Visible = False
found = 0
Label5.Visible = True
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
scanme.Refresh
step_x = (new_x - old_x) / (scan_resx)

AOLChatSend2 ("<font face=" + Chr(34) + "arial" + Chr(34) + "color=" + Chr(34) + "#000080" + Chr(34) + ">•| caustik converter incoming macro |•")
Pause scan_timeout
AOLChatSend2 ("<font face=" + Chr(34) + "arial" + Chr(34) + "color=" + Chr(34) + "#000080" + Chr(34) + ">•| macro name - " & macro_name & " | Char - " & "<font face=" & Chr(34) & scan_font & Chr(34) & ">" & scan_letter & "<font face=" + Chr(34) + "arial" + Chr(34) + ">" + " |•")
Pause scan_timeout * 2
stat.Tag = (100 / ((new_y - old_y) / (step_x * scan_offset)))
stat.Cls
stat.Line (0, 0)-(stat.Tag, 13), 0, BF
lasty = ""

For Y2 = old_y To new_y Step step_x * scan_offset
stat.Tag = stat.Tag + (100 / ((new_y - old_y) / (step_x * scan_offset)))
stat.Cls
stat.Line (0, 0)-(stat.Tag, 13), 0, BF
For X2 = old_x To new_x Step step_x
If scanning = 0 Then
scan_active = 1
Label5.Visible = False
Label3.Visible = True
Label4.Visible = True
Label8.Visible = True
mshop_frm.scanme.AutoRedraw = False
Label9.Enabled = True
Exit Sub
End If
If (VbtoAol(scanme.Point(X2 + (step_x / 2), Y2 + (step_y / 2)))) <> "FEFEFE" And found = 0 Then
found = 1
End If
'word$ = word$ + "<font color=#" & VbtoAol(scanme.Point(X2 + (step_x / 2), Y2 + (step_y / 2))) & ">" + scan_letter
'lasty = VbtoAol(scanme.Point(X2 + (step_x / 2), Y2 + (step_y / 2)))
mshop_frm.scanme.AutoRedraw = True
If scan_mix = 1 Then
word$ = word$ + "<font color=#" & cmix(VbtoAol(scanme.Point(X2 - step_x, Y2 - step_y)), VbtoAol(scanme.Point(X2 + step_x, Y2 + step_y))) & ">" + scan_letter

lasty = cmix(VbtoAol(scanme.Point(X2 - step_x, Y2 - step_y)), VbtoAol(scanme.Point(X2 + step_x, Y2 + step_y)))
Else

word$ = word$ + "<font color=#" & VbtoAol(scanme.Point(X2, Y2)) & ">" + scan_letter
lasty = VbtoAol(scanme.Point(X2, Y2))
End If
Next
If found = 1 Then
    If scan_bold Then addy$ = addy$ + "<b>"
    If scan_italic Then addy$ = addy$ + "<i>"
    If scan_underline Then addy$ = addy$ + "<u>"
        AOLChatSend2 "<font face=" + Chr(34) + scan_font + Chr(34) + ">" + addy$ + word$

Pause scan_timeout
End If
spot = spot + 1
word$ = ""
Next
Pause scan_timeout * 2
AOLChatSend2 ("<font face=" + Chr(34) + "arial" + Chr(34) + "color=" + Chr(34) + "#000080" + Chr(34) + ">•| caustik converter macro complete |•")

scanning = 0
scan_active = 1
Label3.Visible = True
Label4.Visible = True
Label8.Visible = True
Label5.Visible = False
mshop_frm.scanme.AutoRedraw = False
Label9.Enabled = True
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Showactive Label4
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
scanning = 0
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Showactive Label5
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
about_frm.aboutmovie.Base = App.Path
about_frm.aboutmovie.Movie = App.Path + "\about.swf"
about_frm.aboutmovie.BackgroundColor = RGB(0, 0, 0)
about_frm.aboutmovie.BGColor = RGB(0, 0, 0)
about_frm.Show (1)
about_frm.aboutmovie.GotoFrame (1)
about_frm.aboutmovie.Play
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Showactive Label6
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
options_frm.Visible = True
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Showactive Label7
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
other_frm.Visible = True
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Showactive Label8
End Sub

Private Sub RichTextBox1_Change()

End Sub

Public Sub autoselect()
On Error Resume Next
scanme.AutoRedraw = True
scanme.Line (0, 0)-(scanme.Width, scanme.Height), RGB(256, 256, 256), BF
'Call scanme.PaintPicture(stdole.LoadPicture(filename), 0, 0, scanme.Width, scanme.Height)
Call scanme.PaintPicture(scanme.Picture, 0, 0, scanme.Width, scanme.Height)
scanme.AutoRedraw = False
scanme.Refresh
old_x = 0
old_y = 0
new_x = scanme.Width
new_y = scanme.Height
End Sub

Private Sub Label9_Click()
Me.WindowState = 1
End Sub

Private Sub ShockwaveFlash1_FSCommand(ByVal command As String, ByVal args As String)

End Sub

