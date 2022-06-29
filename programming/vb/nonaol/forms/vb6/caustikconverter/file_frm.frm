VERSION 5.00
Begin VB.Form file_frm 
   BorderStyle     =   0  'None
   Caption         =   "Open File"
   ClientHeight    =   3075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "<-"
      Height          =   225
      Left            =   4560
      TabIndex        =   13
      Top             =   1750
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.CommandButton Command1 
      Caption         =   "->"
      Height          =   225
      Left            =   5520
      TabIndex        =   12
      Top             =   1775
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox scanboy 
      Height          =   540
      Left            =   4560
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   2160
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2640
      Top             =   240
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "file_frm.frx":0000
      Left            =   2280
      List            =   "file_frm.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2700
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Top             =   2280
      Width           =   2055
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   1935
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   1935
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   2280
      Pattern         =   "*.bmp;*.dib;*.gif;*.jpg;*.wmf;*.emf;*.ico;*.cur;*.dll;*.exe"
      System          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1/1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5025
      TabIndex        =   14
      Top             =   1735
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Open File"
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
      Left            =   4800
      TabIndex        =   6
      Top             =   2760
      Width           =   675
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00FFFFFF&
      X1              =   5760
      X2              =   5760
      Y1              =   2640
      Y2              =   3000
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00FFFFFF&
      X1              =   4440
      X2              =   5760
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00FFFFFF&
      X1              =   4320
      X2              =   5760
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00FFFFFF&
      X1              =   2160
      X2              =   2160
      Y1              =   2640
      Y2              =   3000
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   4440
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   120
      Y1              =   2640
      Y2              =   3000
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Open File Type:"
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
      TabIndex        =   9
      Top             =   2760
      Width           =   1125
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      X1              =   4560
      X2              =   5760
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   5880
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FFFFFF&
      X1              =   4560
      X2              =   5760
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "preview"
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
      Left            =   4920
      TabIndex        =   7
      Top             =   2160
      Width           =   600
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   4440
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   4440
      X2              =   4440
      Y1              =   360
      Y2              =   3000
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   120
      Y1              =   360
      Y2              =   2640
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   4440
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Image preview 
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   5520
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   5840
      X2              =   5700
      Y1              =   40
      Y2              =   185
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   5700
      X2              =   5840
      Y1              =   40
      Y2              =   185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "caustik converter - open file"
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
      TabIndex        =   0
      Top             =   0
      Width           =   2025
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5940
   End
   Begin VB.Image Image2 
      Height          =   3090
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "file_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error Resume Next
If curicon < iconcount Then curicon = curicon + 1
scanboy.Line (0, 0)-(scanboy.Width, scanboy.Height), RGB(255, 255, 255), BF
If (Len(Dir1.Path) > 3) Then filename2 = Dir1.Path + "\" + Text1.text Else filename2 = Dir1.Path + Text1.text
ok = ExtractIcon(0, filename2, curicon)
Label7.Caption = Str(curicon + 1) + "/" + Str(iconcount + 1)
scanboy.AutoRedraw = True
Call DrawIcon(scanboy.hDC, 0, 0, ok)
preview.Picture = scanboy.Image
preview.Refresh
End Sub

Private Sub Command2_Click()
On Error Resume Next
If curicon > 0 Then curicon = curicon - 1
scanboy.Line (0, 0)-(scanboy.Width, scanboy.Height), RGB(255, 255, 255), BF
If (Len(Dir1.Path) > 3) Then filename2 = Dir1.Path + "\" + Text1.text Else filename2 = Dir1.Path + Text1.text
ok = ExtractIcon(0, filename2, curicon)
Label7.Caption = Str(curicon + 1) + "/" + Str(iconcount + 1)
scanboy.AutoRedraw = True
Call DrawIcon(scanboy.hDC, 0, 0, ok)
preview.Picture = scanboy.Image
preview.Refresh
End Sub

Private Sub Dir1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
If Err Then
Dir1.Path = Mid(App.Path, 1, 3)
Drive1.Drive = Mid(App.Path, 1, 1)
End If
End Sub

Private Sub File1_KeyUp(KeyCode As Integer, Shift As Integer)
Call File1_MouseUp(1, 0, 0, 0)
End Sub

Private Sub File1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1.text = File1.filename
preview.Picture = New stdole.StdPicture
On Error Resume Next
If (Len(Dir1.Path) > 3) Then filename2 = Dir1.Path + "\" + Text1.text Else filename2 = Dir1.Path + Text1.text
If Mid(Text1.text, Len(Text1.text) - 3) = ".exe" Or Mid(Text1.text, Len(Text1.text) - 3) = ".dll" Then
Label7.Visible = True
Command1.Visible = True
Command2.Visible = True
scanboy.Line (0, 0)-(scanboy.Width, scanboy.Height), RGB(255, 255, 255), BF
ok = ExtractIcon(0, filename2, 0)
iconcount = ExtractIcon(0, filename2, -1) - 1
curicon = 0
Label7.Caption = "1/" + Str(iconcount + 1)
scanboy.AutoRedraw = True
Call DrawIcon(scanboy.hDC, 0, 0, ok)
preview.Picture = scanboy.Image
preview.Refresh
Else
preview.Picture = LoadPicture(filename2)
Label7.Visible = False
Command1.Visible = False
Command2.Visible = False
End If
' exe or DLL
End Sub

Private Sub Form_Load()
On Error Resume Next
Combo1.text = "All Picture Files"
Image2.Picture = mshop_frm.Image2.Picture
Image1.Picture = mshop_frm.Image1.Picture
StayOnTop Me
Drive1.Drive = currentsettings.imagedrive
Drive1.Refresh
Dir1.Path = currentsettings.imagedir
Dir1.Refresh
File1.Path = currentsettings.imagedir
File1.Refresh
DoEvents
File1.filename = currentsettings.imagefile
File1.Path = Dir1.Path
Call File1_MouseUp(0, 0, 0, 0)
Call Label3_MouseDown(0, 0, 0, 0)
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Me
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Me
End Sub

Private Sub Label2_Click()
Me.Visible = False
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
currentsettings.imagedrive = Drive1.Drive
currentsettings.imagedir = Dir1.Path
currentsettings.imagefile = File1.filename
If (Len(Dir1.Path) > 3) Then filename = Dir1.Path + "\" + Text1.text Else filename = Dir1.Path + Text1.text
If Mid(Text1.text, Len(Text1.text) - 3) = ".exe" Or Mid(Text1.text, Len(Text1.text) - 3) = ".dll" Then
mshop_frm.scanme.Picture = preview.Picture
Else
mshop_frm.scanme.Picture = LoadPicture(filename)
End If


If Err Then
filename = ""
Else
macro_name = LCase(Text1.text)
mshop_frm.mname.Caption = macro_name
mshop_frm.Label4.Visible = True
mshop_frm.Label8.Visible = True
Me.Visible = False
orig_offsetH = mshop_frm.scanme.Picture.Height / mshop_frm.scanme.Picture.Width
orig_offsetW = mshop_frm.scanme.Picture.Width / mshop_frm.scanme.Picture.Height
mshop_frm.scanme.Height = mshop_frm.scanme.Width * orig_offsetH
mshop_frm.Height = mshop_frm.scanme.Height + 1305
old_x = 0
old_y = 0
new_x = mshop_frm.Width
new_y = mshop_frm.Height
Call mshop_frm.autoselect
End If
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Showactive Label3
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Me
End Sub

Private Sub Label6_Click()
Call Label3_MouseDown(1, 0, 0, 0)
End Sub

Private Sub Timer1_Timer()
If Combo1.text = "All Picture Files" Then File1.Pattern = "*.bmp;*.dib;*.gif;*.jpg;*.wmf;*.emf;*.ico;*.cur;*.dll;*.exe" Else File1.Pattern = "*.*"
End Sub
