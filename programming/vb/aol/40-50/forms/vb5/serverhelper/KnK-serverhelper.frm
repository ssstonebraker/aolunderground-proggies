VERSION 5.00
Object = "{576A0B60-AD7A-11CF-959F-0020AF557A1A}#1.31#0"; "VDGT.OCX"
Begin VB.Form Form7 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "KnK Server Aid                               By: KnK"
   ClientHeight    =   5400
   ClientLeft      =   2535
   ClientTop       =   255
   ClientWidth     =   5175
   ForeColor       =   &H00C0C0C0&
   Icon            =   "KnK-serverhelper.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "KnK-serverhelper.frx":030A
   Picture         =   "KnK-serverhelper.frx":0614
   ScaleHeight     =   5400
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer19 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4680
      Top             =   4080
   End
   Begin VB.Timer Timer18 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3720
      Top             =   4080
   End
   Begin VB.Timer Timer17 
      Interval        =   5
      Left            =   4800
      Top             =   2760
   End
   Begin VB.ListBox List4 
      Height          =   1425
      Left            =   3720
      TabIndex        =   33
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1200
      Top             =   4080
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1200
      Top             =   4440
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   4080
   End
   Begin VB.Timer Timer10 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   4440
   End
   Begin VB.Timer Timer11 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2400
      Top             =   4080
   End
   Begin VB.Timer Timer12 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2400
      Top             =   4440
   End
   Begin VB.Timer Timer15 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3000
      Top             =   4080
   End
   Begin VB.Timer Timer16 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3000
      Top             =   4440
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3000
      Top             =   3360
   End
   Begin VB.Timer Timer13 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3000
      Top             =   3000
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1005
      Left            =   3600
      TabIndex        =   25
      Top             =   480
      Width           =   1335
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2400
      Top             =   3360
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   3360
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1200
      Top             =   3360
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2400
      Top             =   3000
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   3000
   End
   Begin VDGT.VDPush VDPush6 
      Height          =   375
      Left            =   3540
      TabIndex        =   20
      Top             =   2040
      Width           =   1530
      _Version        =   65536
      _ExtentX        =   2708
      _ExtentY        =   661
      _StockProps     =   79
      ForeColor       =   8388608
      BackColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorDown   =   8388608
      BevelWidth      =   0
      CornerColor     =   8388608
      DarkColor       =   8388608
      ForeColorDisabled=   8388608
      LightColor      =   8388608
      Picture         =   "KnK-serverhelper.frx":28FCC
      PictureDown     =   "KnK-serverhelper.frx":2A9E8
      ShowFocus       =   0
      SinkDeep        =   -1  'True
      TransColor      =   8388608
   End
   Begin VDGT.VDPush VDPush5 
      Height          =   330
      Left            =   4140
      TabIndex        =   19
      Top             =   1740
      Width           =   945
      _Version        =   65536
      _ExtentX        =   1667
      _ExtentY        =   591
      _StockProps     =   79
      ForeColor       =   8388608
      BackColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorDown   =   8388608
      BevelWidth      =   0
      CornerColor     =   8388608
      DarkColor       =   8388608
      ForeColorDisabled=   8388608
      LightColor      =   8388608
      Picture         =   "KnK-serverhelper.frx":2C404
      PictureDown     =   "KnK-serverhelper.frx":2D308
      ShowFocus       =   0
      SinkDeep        =   -1  'True
      TransColor      =   8388608
   End
   Begin VDGT.VDPush VDPush4 
      Height          =   330
      Left            =   3555
      TabIndex        =   18
      Top             =   1740
      Width           =   585
      _Version        =   65536
      _ExtentX        =   1041
      _ExtentY        =   591
      _StockProps     =   79
      ForeColor       =   8388608
      BackColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorDown   =   8388608
      BevelWidth      =   0
      CornerColor     =   8388608
      DarkColor       =   8388608
      ForeColorDisabled=   8388608
      LightColor      =   8388608
      Picture         =   "KnK-serverhelper.frx":2E20C
      PictureDown     =   "KnK-serverhelper.frx":2EB90
      ShowFocus       =   0
      SinkDeep        =   -1  'True
      TransColor      =   8388608
   End
   Begin VDGT.VDPush VDPush3 
      Height          =   255
      Left            =   2520
      TabIndex        =   17
      Top             =   1800
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   79
      ForeColor       =   8388608
      BackColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorDown   =   8388608
      BevelWidth      =   0
      CornerColor     =   8388608
      DarkColor       =   8388608
      ForeColorDisabled=   8388608
      LightColor      =   8388608
      Picture         =   "KnK-serverhelper.frx":2F514
      PictureDown     =   "KnK-serverhelper.frx":30418
      ShowFocus       =   0
      SinkDeep        =   -1  'True
      TransColor      =   8388608
   End
   Begin VDGT.VDPush VDPush2 
      Height          =   375
      Left            =   1200
      TabIndex        =   16
      Top             =   2040
      Width           =   2265
      _Version        =   65536
      _ExtentX        =   3995
      _ExtentY        =   661
      _StockProps     =   79
      ForeColor       =   8388608
      BackColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorDown   =   8388608
      BevelWidth      =   0
      CornerColor     =   8388608
      DarkColor       =   8388608
      ForeColorDisabled=   8388608
      LightColor      =   8388608
      Picture         =   "KnK-serverhelper.frx":3131C
      PictureDisabled =   "KnK-serverhelper.frx":33B54
      PictureDown     =   "KnK-serverhelper.frx":3638C
      ShowFocus       =   0
      SinkDeep        =   -1  'True
   End
   Begin VDGT.VDPush VDPush1 
      Height          =   365
      Left            =   0
      TabIndex        =   15
      Top             =   2040
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   644
      _StockProps     =   79
      ForeColor       =   8388608
      BackColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorDown   =   8388608
      BevelWidth      =   0
      CornerColor     =   8388608
      DarkColor       =   8388608
      ForeColorDisabled=   8388608
      LightColor      =   8388608
      Picture         =   "KnK-serverhelper.frx":38BC4
      PictureDown     =   "KnK-serverhelper.frx":39F00
      ShowFocus       =   0
      SinkDeep        =   -1  'True
      TransColor      =   8388608
      Transparent     =   1
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00800000&
      Caption         =   "Option3"
      Height          =   240
      Left            =   50
      TabIndex        =   6
      Top             =   1440
      Width           =   215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1200
      Top             =   3000
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1380
      Left            =   1650
      MouseIcon       =   "KnK-serverhelper.frx":3B23C
      TabIndex        =   5
      Top             =   645
      Width           =   750
   End
   Begin VB.ListBox List3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   930
      ItemData        =   "KnK-serverhelper.frx":3B546
      Left            =   2520
      List            =   "KnK-serverhelper.frx":3B548
      TabIndex        =   4
      Top             =   600
      Width           =   800
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   50
      TabIndex        =   3
      Top             =   1200
      Width           =   215
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   50
      TabIndex        =   2
      Top             =   960
      Width           =   215
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   230
      Left            =   120
      TabIndex        =   1
      Text            =   "Item to find"
      Top             =   1755
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   250
      Left            =   120
      TabIndex        =   0
      Text            =   "Screen Name"
      Top             =   645
      Width           =   1335
   End
   Begin VB.Label Label18 
      BackColor       =   &H00000000&
      Caption         =   "                AOL95 Timers"
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   1080
      TabIndex        =   32
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label Label19 
      BackColor       =   &H00000000&
      Caption         =   "AO-NiN"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   31
      Top             =   4155
      Width           =   615
   End
   Begin VB.Label Label20 
      BackColor       =   &H00000000&
      Caption         =   "valkyrie"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   30
      Top             =   4155
      Width           =   615
   End
   Begin VB.Label Label21 
      BackColor       =   &H00000000&
      Caption         =   "IM"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   29
      Top             =   4155
      Width           =   375
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000000&
      Caption         =   "embrace"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   28
      Top             =   4155
      Width           =   615
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "chat"
      Height          =   255
      Left            =   0
      TabIndex        =   27
      Top             =   3600
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   2040
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   "embrace"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   26
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label17 
      BackColor       =   &H00000000&
      Caption         =   "IM"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   24
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label16 
      BackColor       =   &H00000000&
      Caption         =   "valkyrie"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   23
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label15 
      BackColor       =   &H00000000&
      Caption         =   "AO-NiN"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   22
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label14 
      BackColor       =   &H00000000&
      Caption         =   "              AOL 4.o Timers"
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   1080
      TabIndex        =   21
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3720
      TabIndex        =   14
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   2400
      TabIndex        =   13
      Top             =   255
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   1560
      TabIndex        =   12
      Top             =   255
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   11
      Top             =   255
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   0
      TabIndex        =   10
      Top             =   255
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label5 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MoveMe As Boolean

Private Sub abt_Click()
MsgBox "Not done, IM me for help"

End Sub

Private Sub ad1_Click()
End Sub

Private Sub ad2_Click()
End Sub

Private Sub ad3_Click()
MsgBox "Dont have this one yet", 16, "sorry"
End Sub

Private Sub afk_Click()

End Sub

Private Sub att_Click()


End Sub

Private Sub claa_Click()

End Sub

Private Sub Command1_Click()
Text1.text = "Xx KnK xX"
End Sub

Private Sub Command2_Click()


End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()




End Sub

Private Sub Command6_Click()

End Sub

Private Sub Command7_Click()

End Sub

Private Sub dclaa_Click()


End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()
StayOnTop Me
'Form Sizes
Form7.Width = 3495
Form7.Height = 2445
'Set the main preferences
Option1 = True
'Generate the list of 0 to 10000
For i = 0 To 10000
List1.AddItem i
List1.ListIndex = 0
Next i
'''''''''''

Dim a As Variant
Dim b As Variant
On Error GoTo kook
a = 1
List4.Clear
Open CStr(App.Path + "\filename.lst") For Input As a
While (EOF(a) = False)
Line Input #a, b
List4.AddItem b
Wend

Close a
'End If
kook:
'''''''''''''''''''''
'<!------Start of .INI stuff
TimeKnK$ = GetFromINI("Pauses", "TimeKnK", App.Path + "\KnK.ini")
Clearornot$ = GetFromINI("List", "clearornot", App.Path + "\KnK.ini")
Color$ = GetFromINI("ascii", "Color", App.Path + "\KnK.ini")
adver$ = GetFromINI("Scroll", "adver", App.Path + "\KnK.ini")
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")

'<!------End of .INI stuff


 
End Sub

Private Sub greetz_Click()


End Sub

Private Sub gtname_Click()

End Sub

Private Sub hlp_Click()


End Sub

Private Sub idle_Click()

End Sub

Private Sub knk_Click()

End Sub

Private Sub knk4o_Click()
End Sub

Private Sub knkhmp_Click()

End Sub

Private Sub Label1_Click()
Playwav (App.Path + "\blip.WAV")
Loads2$ = GetFromINI("Exit", "Loads2", App.Path + "\KnK.ini")
If Loads2$ = "no" Then
End
End If
If Loads2$ = "yes" Then
Unload Form7
Form11.Show
End If
End Sub

Private Sub Label10_Click()
MsgBox "Ya thats me, KnK,  hope you like the prog!", vbInformation, "Ya Dats me"
End Sub

Private Sub Label2_Click()
Playwav (App.Path + "\blip.WAV")
Form10.PopupMenu Form10.File, 1
End Sub

Private Sub Label22_Click()

End Sub

Private Sub Label3_Click()
Playwav (App.Path + "\blip.WAV")
Form10.PopupMenu Form10.Toolz, 2
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Me
End Sub

Private Sub Label6_Click()
Playwav (App.Path + "\blip.WAV")
Form10.PopupMenu Form10.Options, 3
End Sub

Private Sub Label7_Click()
Playwav (App.Path + "\blip.WAV")
If Form7.Width = 5070 Then
Exit Sub
End If
Form7.Width = 3495
Form7.Width = 3500
Form7.Width = 3550
Form7.Width = 3600
Form7.Width = 3650
Form7.Width = 3700
Form7.Width = 3750
Form7.Width = 3800
Form7.Width = 3850
Form7.Width = 3900
Form7.Width = 3950
Form7.Width = 4000
Form7.Width = 4050
Form7.Width = 4100
Form7.Width = 4150
Form7.Width = 4200
Form7.Width = 4250
Form7.Width = 4300
Form7.Width = 4350
Form7.Width = 4400
Form7.Width = 4450
Form7.Width = 4500
Form7.Width = 4550
Form7.Width = 4600
Form7.Width = 4650
Form7.Width = 4700
Form7.Width = 4750
Form7.Width = 4800
Form7.Width = 4850
Form7.Width = 4900
Form7.Width = 4950
Form7.Width = 5000
Form7.Width = 5050
Form7.Width = 5070
End Sub

Private Sub List1_DblClick()
Playwav (App.Path + "\blip.WAV")
If List3.ListCount = 0 Then List3.AddItem List1
For i = 0 To List3.ListCount - 1
num = List3.List(i)
If num = List1 Then Exit Sub
Next i
List3.AddItem List1
End Sub

Private Sub List2_DblClick()
Playwav (App.Path + "\blip.WAV")
Text1.text = List2
End Sub

Private Sub List3_DblClick()
Playwav (App.Path + "\blip.WAV")
List3.RemoveItem List3.ListIndex
End Sub

Private Sub mail_Click()

End Sub

Private Sub min_Click()


End Sub

Private Sub off_Click()

End Sub

Private Sub on_Click()

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1 = Picture2
TimeOut 0.4
Picture2 = Picture1
End Sub

Private Sub PooK_Click()
End Sub

Private Sub pw_Click()
End Sub

Private Sub rb_Click()


End Sub

Private Sub s2_Click()
sec0.Checked = False
sec1.Checked = False
sec2.Checked = True
sec3.Checked = False
sec4.Checked = False
Label4.Caption = "2"
End Sub

Private Sub sec0_Click()

End Sub

Private Sub sec1_Click()

Label4.Caption = "1"
End Sub

Private Sub sec2_Click()

End Sub

Private Sub sec3_Click()

sec1.Checked = False
sec2.Checked = False
sec3.Checked = True
sec4.Checked = False
Label4.Caption = "3"
End Sub

Private Sub sec4_Click()

sec1.Checked = False
sec2.Checked = False
sec3.Checked = False
sec4.Checked = True
Label4.Caption = "4"
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture3 = Picture2

End Sub

Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture3 = Picture3
End Sub

Private Sub List4_Click()
Text1.text = List2
End Sub

Private Sub Option1_Click()
Text2.Enabled = False
End Sub

Private Sub Option2_Click()
Text2.Enabled = True
Text2.text = "ReVeNgE"
End Sub

Private Sub Option3_Click()
Text2.Enabled = False
End Sub

Private Sub SystemTray1_Error(ByVal ErrorNumber As Integer)

End Sub

Private Sub Timer1_Timer()
TimeKnK$ = GetFromINI("Pauses", "TimeKnK", App.Path + "\KnK.ini")
Clearornot$ = GetFromINI("List", "clearornot", App.Path + "\KnK.ini")
Color$ = GetFromINI("ascii", "Color", App.Path + "\KnK.ini")
adver$ = GetFromINI("Scroll", "adver", App.Path + "\KnK.ini")
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")



If Timer1.Enabled = False Then Exit Sub
For i = 0 To List3.ListCount - 1
SendChat ("/" + Text1 + " Send " + List3.List(i))
If Timer1.Enabled = False Then Exit Sub
'Time
If TimeKnK$ = "1" Then
TimeOut (1)
End If
If TimeKnK$ = "2" Then
TimeOut (2)
End If
If TimeKnK$ = "3" Then
TimeOut (3)
End If
If TimeKnK$ = "4" Then
TimeOut (4)
End If
Next i
Timer1.Enabled = False
If adver$ = "yes" Then
'Black blue Black ascii
If Color$ = "bbb" Then
SendChat BlackBlueBlack("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
SendChat BlackBlueBlack("«-×´¯`°  Done Requesting °´¯`×-»")
TimeOut (0.5)
SendChat BlackBlueBlack("«-×´¯`° Get a copy at  °´¯`×-»")
TimeOut (0.9)
SendChat BlackBlueBlack("«-×´¯`° http://knk.tierranet.com/serv")
End If

'Black green Black ascii
If Color$ = "bgb" Then
SendChat BlackGreenBlack("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
SendChat BlackGreenBlack("«-×´¯`°  Done Requesting °´¯`×-»")
TimeOut (0.5)
SendChat BlackGreenBlack("«-×´¯`° Get a copy at  °´¯`×-»")
TimeOut (0.9)
SendChat BlackGreenBlack("«-×´¯`° http://knk.tierranet.com/serv")
End If
End If
If adver$ = "no" Then
End If
VDPush2.Enabled = True
If Clearornot$ = "Clear" Then
List3.Clear
End If
If Clearornot$ = "dontclear" Then
Exit Sub
End If

End Sub

Private Sub ver_Click()

End Sub

Private Sub wrb_Click()
MsgBox "not done yet"
End Sub

Private Sub WebBrowser1_BeforeNavigate(ByVal URL As String, ByVal FLAGS As Long, ByVal TargetFrameName As String, PostData As Variant, ByVal Headers As String, Cancel As Boolean)

End Sub

Private Sub VDTTip1_Popup(ctlName As String, ByVal ctlIndex As Integer, ByVal ctlHwnd As Long, strText As String, ByVal X As Long, ByVal Y As Long)

End Sub

Private Sub VDAMon1_WndProc(message As Long, wParam As Long, lParam As Long, lRetval As Long, bCancel As Boolean)

End Sub

Private Sub Timer10_Timer()
If Text1 = "Screen Name" Then
MsgBox "You need to get the Servers SN First!", vbExclamation, "Need to get a SN"
Timer10 = False
Exit Sub
End If
If Option1 = True Then
AOLChatSend ("-" + Text1 + " send list")
Timer10 = False
End If
If Option2 = True Then
AOLChatSend ("-" + Text1 + " find " + Text2)
Timer10 = False
End If
If Option3 = True Then
AOLChatSend ("-" + Text1 + " status")
Timer10 = False
End If
Timer10 = False

End Sub

Private Sub Timer11_Timer()
TimeKnK$ = GetFromINI("Pauses", "TimeKnK", App.Path + "\KnK.ini")
Clearornot$ = GetFromINI("List", "clearornot", App.Path + "\KnK.ini")
Color$ = GetFromINI("ascii", "Color", App.Path + "\KnK.ini")
adver$ = GetFromINI("Scroll", "adver", App.Path + "\KnK.ini")
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")


If Timer11.Enabled = False Then Exit Sub
For i = 0 To List3.ListCount - 1
Call AOLInstantMessage(Text1, "Send " + List3.List(i))
If Timer11.Enabled = False Then Exit Sub
If TimeKnK$ = "1" Then
TimeOut (1)
End If
If TimeKnK$ = "2" Then
TimeOut (2)
End If
If TimeKnK$ = "3" Then
TimeOut (3)
End If
If TimeKnK$ = "4" Then
TimeOut (4)
End If
Next i
Timer11.Enabled = False
If adver$ = "yes" Then
AOLChatSend ("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
AOLChatSend ("«-×´¯`°  Done Requesting °´¯`×-»")
TimeOut (0.5)
AOLChatSend ("«-×´¯`° Get a copy at  °´¯`×-»")
TimeOut (0.9)
AOLChatSend ("«-×´¯`° http://knk.tierranet.com/serv")
End If
If adver$ = "no" Then
End If
VDPush2.Enabled = True
If Clearornot$ = "Clear" Then
List3.Clear
End If
If Clearornot$ = "dontclear" Then
Exit Sub
End If

End Sub

Private Sub Timer12_Timer()
If Text1 = "Screen Name" Then
MsgBox "You need to get the Servers SN First!", vbExclamation, "Need to get a SN"
Timer12 = False
Exit Sub
End If
If Option1 = True Then
Call AOLInstantMessage(Text1, "Send List")
Timer12 = False
End If
If Option2 = True Then
Call AOLInstantMessage(Text1, "Find " + Text2)
Timer12 = False
End If
If Option3 = True Then
Call AOLInstantMessage(Text1, "Send Status")
Timer12 = False
End If

End Sub

Private Sub Timer13_Timer()
TimeKnK$ = GetFromINI("Pauses", "TimeKnK", App.Path + "\KnK.ini")
Clearornot$ = GetFromINI("List", "clearornot", App.Path + "\KnK.ini")
Color$ = GetFromINI("ascii", "Color", App.Path + "\KnK.ini")
adver$ = GetFromINI("Scroll", "adver", App.Path + "\KnK.ini")
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")


If Timer13.Enabled = False Then Exit Sub
For i = 0 To List3.ListCount - 1
SendChat ("!" + Text1 + " send " + List3.List(i))
If Timer13.Enabled = False Then Exit Sub
If TimeKnK$ = "1" Then
TimeOut (1)
End If
If TimeKnK$ = "2" Then
TimeOut (2)
End If
If TimeKnK$ = "3" Then
TimeOut (3)
End If
If TimeKnK$ = "4" Then
TimeOut (4)
End If
Next i
Timer13.Enabled = False

If adver$ = "yes" Then
'Black blue Black ascii
If Color$ = "bbb" Then
SendChat BlackBlueBlack("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
SendChat BlackBlueBlack("«-×´¯`°  Done Requesting °´¯`×-»")
TimeOut (0.5)
SendChat BlackBlueBlack("«-×´¯`° Get a copy at  °´¯`×-»")
TimeOut (0.9)
SendChat BlackBlueBlack("«-×´¯`° http://knk.tierranet.com/serv")
End If

'Black green Black ascii
If Color$ = "bgb" Then
SendChat BlackGreenBlack("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
SendChat BlackGreenBlack("«-×´¯`°  Done Requesting °´¯`×-»")
TimeOut (0.5)
SendChat BlackGreenBlack("«-×´¯`° Get a copy at  °´¯`×-»")
TimeOut (0.9)
SendChat BlackGreenBlack("«-×´¯`° http://knk.tierranet.com/serv")
End If
End If

If adver$ = "no" Then
End If
VDPush2.Enabled = True
If Clearornot$ = "Clear" Then
List3.Clear
End If
If Clearornot$ = "dontclear" Then
Exit Sub
End If

End Sub

Private Sub Timer14_Timer()
If Text1 = "Screen Name" Then
MsgBox "You need to get the Servers SN First!", vbExclamation, "Need to get a SN"
Timer14 = False
Exit Sub
End If
If Option1 = True Then
SendChat ("!" + Text1 + " send list")
Timer14 = False
End If
If Option2 = True Then
SendChat ("!" + Text1 + " find " + Text2)
Timer14 = False
End If
Timer14 = False
End Sub

Private Sub Timer15_Timer()
TimeKnK$ = GetFromINI("Pauses", "TimeKnK", App.Path + "\KnK.ini")
Clearornot$ = GetFromINI("List", "clearornot", App.Path + "\KnK.ini")
Color$ = GetFromINI("ascii", "Color", App.Path + "\KnK.ini")
adver$ = GetFromINI("Scroll", "adver", App.Path + "\KnK.ini")
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")


If Timer15.Enabled = False Then Exit Sub
For i = 0 To List3.ListCount - 1
AOLChatSend ("!" + Text1 + " send " + List3.List(i))
If Timer15.Enabled = False Then Exit Sub
If TimeKnK$ = "1" Then
TimeOut (1)
End If
If TimeKnK$ = "2" Then
TimeOut (2)
End If
If TimeKnK$ = "3" Then
TimeOut (3)
End If
If TimeKnK$ = "4" Then
TimeOut (4)
End If
Next i
Timer15.Enabled = False

If adver$ = "yes" Then
AOLChatSend ("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
AOLChatSend ("«-×´¯`°  Done Requesting °´¯`×-»")
TimeOut (0.5)
AOLChatSend ("«-×´¯`° Get a copy at  °´¯`×-»")
TimeOut (0.9)
AOLChatSend ("«-×´¯`° http://knk.tierranet.com/serv")
End If
If adver$ = "no" Then
End If
VDPush2.Enabled = True
If Clearornot$ = "Clear" Then
List3.Clear
End If
If Clearornot$ = "dontclear" Then
Exit Sub
End If

End Sub

Private Sub Timer16_Timer()
If Text1 = "Screen Name" Then
MsgBox "You need to get the Servers SN First!", vbExclamation, "Need to get a SN"
Timer16 = False
Exit Sub
End If
If Option1 = True Then
AOLChatSend ("!" + Text1 + " send list")
Timer16 = False
End If
If Option2 = True Then
AOLChatSend ("!" + Text1 + " find " + Text2)
Timer16 = False
End If
Timer16 = False

End Sub

Private Sub Timer18_Timer()
 Color$ = GetFromINI("ascii", "Color", App.Path + "\KnK.ini")
adver$ = GetFromINI("Scroll", "adver", App.Path + "\KnK.ini")
Timer17.Enabled = False
'<!----------------Normal------------------!>
If Label12.Caption = "chat" Then
For i = 0 To List4.ListCount - 1
SendChat ("/" + Text1 + " Find " + List4.List(i))
'If Timer17.Enabled = False Then Exit Sub
TimeOut (2.5)
Next i
Timer18.Enabled = False
End If
'<!----------------Valkerie------------------!>
If Label12.Caption = "valkyrie" Then
For i = 0 To List4.ListCount - 1
SendChat ("-" + Text1 + " find " + List4.List(i))
TimeOut (2.5)
Next i
Timer18.Enabled = False
End If
'<!-----------IM---------!>
If Label12.Caption = "im" Then
For i = 0 To List4.ListCount - 1
Call IMKeyword(Text1, "Find " + List4.List(i))
TimeOut (2.5)
Next i
Timer18.Enabled = False
End If
'<!-----------embrace-------!>
If Label12.Caption = "embrace" Then
For i = 0 To List4.ListCount - 1
SendChat ("!" + Text1 + " send " + List4.List(i))
TimeOut (2.5)
Next i
Timer18.Enabled = False
End If
'<!----------------Advertising

If adver$ = "yes" Then
'Black blue Black ascii
If Color$ = "bbb" Then
SendChat BlackBlueBlack("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
SendChat BlackBlueBlack("«-×´¯`°  Request Bin °´¯`×-»")
TimeOut (0.5)
SendChat BlackBlueBlack("«-×´¯`° Deactivated  °´¯`×-»")
End If

'Black green Black ascii
If Color$ = "bgb" Then
SendChat BlackGreenBlack("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
SendChat BlackGreenBlack("«-×´¯`°  Request Bin °´¯`×-»")
TimeOut (0.5)
SendChat BlackGreenBlack("«-×´¯`° Deactivated  °´¯`×-»")
End If
End If
If adver$ = "no" Then
End If
Timer18.Enabled = False
Timer17.Enabled = False
End Sub

Private Sub Timer19_Timer()
adver$ = GetFromINI("Scroll", "adver", App.Path + "\KnK.ini")

'<!----------------Normal------------------!>
If Label12.Caption = "chat" Then
For i = 0 To List4.ListCount - 1
AOLChatSend ("/" + Text1 + " Find " + List4.List(i))
'If Timer17.Enabled = False Then Exit Sub
TimeOut (2.5)
Next i
Timer19.Enabled = False
End If
'<!----------------Valkerie------------------!>
If Label12.Caption = "valkyrie" Then
For i = 0 To List4.ListCount - 1
AOLChatSend ("-" + Text1 + " find " + List4.List(i))
TimeOut (2.5)
Next i
Timer19.Enabled = False
End If
'<!-----------IM---------!>
If Label12.Caption = "im" Then
For i = 0 To List4.ListCount - 1
Call AOLInstantMessage(Text1, "Find " + List4.List(i))
TimeOut (2.5)
Next i
Timer19.Enabled = False
End If
'<!-----------embrace-------!>
If Label12.Caption = "embrace" Then
For i = 0 To List4.ListCount - 1
AOLChatSend ("!" + Text1 + " send " + List4.List(i))
TimeOut (2.5)
Next i
Timer19.Enabled = False
End If
'<!----------------Advertising

If adver$ = "yes" Then
AOLChatSend ("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
AOLChatSend ("«-×´¯`°  Request Bin °´¯`×-»")
TimeOut (0.5)
AOLChatSend ("«-×´¯`° Deactivated  °´¯`×-»")
End If

End If
If adver$ = "no" Then
End If
Timer19.Enabled = False

End Sub

Private Sub Timer2_Timer()
TimeKnK$ = GetFromINI("Pauses", "TimeKnK", App.Path + "\KnK.ini")
Clearornot$ = GetFromINI("List", "clearornot", App.Path + "\KnK.ini")
Color$ = GetFromINI("ascii", "Color", App.Path + "\KnK.ini")
adver$ = GetFromINI("Scroll", "adver", App.Path + "\KnK.ini")
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")



If Timer2.Enabled = False Then Exit Sub
For i = 0 To List3.ListCount - 1
SendChat ("-" + Text1 + " send " + List3.List(i))
If Timer2.Enabled = False Then Exit Sub
'Time
If TimeKnK$ = "1" Then
TimeOut (1)
End If
If TimeKnK$ = "2" Then
TimeOut (2)
End If
If TimeKnK$ = "3" Then
TimeOut (3)
End If
If TimeKnK$ = "4" Then
TimeOut (4)
End If
Next i
Timer2.Enabled = False

If adver$ = "yes" Then
'Black blue Black ascii
If Color$ = "bbb" Then
SendChat BlackBlueBlack("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
SendChat BlackBlueBlack("«-×´¯`°  Done Requesting °´¯`×-»")
TimeOut (0.5)
SendChat BlackBlueBlack("«-×´¯`° Get a copy at  °´¯`×-»")
TimeOut (0.9)
SendChat BlackBlueBlack("«-×´¯`° http://knk.tierranet.com/serv")
End If

'Black green Black ascii
If Color$ = "bgb" Then
SendChat BlackGreenBlack("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
SendChat BlackGreenBlack("«-×´¯`°  Done Requesting °´¯`×-»")
TimeOut (0.5)
SendChat BlackGreenBlack("«-×´¯`° Get a copy at  °´¯`×-»")
TimeOut (0.9)
SendChat BlackGreenBlack("«-×´¯`° http://knk.tierranet.com/serv")
End If
If adver$ = "no" Then
End If
End If
VDPush2.Enabled = True
If Clearornot$ = "Clear" Then
List3.Clear
End If
If Clearornot$ = "dontclear" Then
Exit Sub
End If

End Sub

Private Sub Timer3_Timer()
TimeKnK$ = GetFromINI("Pauses", "TimeKnK", App.Path + "\KnK.ini")
Clearornot$ = GetFromINI("List", "clearornot", App.Path + "\KnK.ini")
Color$ = GetFromINI("ascii", "Color", App.Path + "\KnK.ini")
adver$ = GetFromINI("Scroll", "adver", App.Path + "\KnK.ini")
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")



If Timer3.Enabled = False Then Exit Sub
For i = 0 To List3.ListCount - 1
Call IMKeyword(Text1, "Send " + List3.List(i))
If Timer3.Enabled = False Then Exit Sub
If TimeKnK$ = "1" Then
TimeOut (1)
End If
If TimeKnK$ = "2" Then
TimeOut (2)
End If
If TimeKnK$ = "3" Then
TimeOut (3)
End If
If TimeKnK$ = "4" Then
TimeOut (4)
End If
Next i
Timer3.Enabled = False
If adver$ = "yes" Then

'Black blue Black ascii
If Color$ = "bbb" Then
SendChat BlackBlueBlack("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
SendChat BlackBlueBlack("«-×´¯`°  Done Requesting °´¯`×-»")
TimeOut (0.5)
SendChat BlackBlueBlack("«-×´¯`° Get a copy at  °´¯`×-»")
TimeOut (0.9)
SendChat BlackBlueBlack("«-×´¯`° http://knk.tierranet.com/serv")
End If

'Black green Black ascii
If Color$ = "bgb" Then
SendChat BlackGreenBlack("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
SendChat BlackGreenBlack("«-×´¯`°  Done Requesting °´¯`×-»")
TimeOut (0.5)
SendChat BlackGreenBlack("«-×´¯`° Get a copy at  °´¯`×-»")
TimeOut (0.9)
SendChat BlackGreenBlack("«-×´¯`° http://knk.tierranet.com/serv")
End If
End If
If adver$ = "no" Then
End If
VDPush2.Enabled = True
If Clearornot$ = "Clear" Then
List3.Clear
End If
If Clearornot$ = "dontclear" Then
Exit Sub
End If
End Sub

Private Sub Timer4_Timer()
If Text1 = "Screen Name" Then
MsgBox "You need to get the Servers SN First!", vbExclamation, "Need to get a SN"
Timer4 = False
Exit Sub
End If
If Option1 = True Then
SendChat ("/" + Text1 + " Send List")
Timer4 = False
End If
If Option2 = True Then
SendChat ("/" + Text1 + " Find " + Text2)
Timer4 = False
End If
If Option3 = True Then
SendChat ("/" + Text1 + " Send Status")
Timer4 = False
End If
Timer4 = False

End Sub

Private Sub Timer5_Timer()
If Text1 = "Screen Name" Then
MsgBox "You need to get the Servers SN First!", vbExclamation, "Need to get a SN"
Timer5 = False
Exit Sub
End If
If Option1 = True Then
SendChat ("-" + Text1 + " send list")
Timer5 = False
End If
If Option2 = True Then
SendChat ("-" + Text1 + " find " + Text2)
Timer5 = False
End If
If Option3 = True Then
SendChat ("-" + Text1 + " status")
Timer5 = False
End If
Timer5 = False

End Sub

Private Sub Timer6_Timer()
If Text1 = "Screen Name" Then
MsgBox "You need to get the Servers SN First!", vbExclamation, "Need to get a SN"
Timer6 = False
Exit Sub
End If
If Option1 = True Then
Call IMKeyword(Text1, "Send List")
Timer6 = False
End If
If Option2 = True Then
Call IMKeyword(Text1, "Find " + Text2)
Timer6 = False
End If
If Option3 = True Then
Call IMKeyword(Text1, "Send Status")
Timer6 = False
End If
End Sub

Private Sub Timer7_Timer()
TimeKnK$ = GetFromINI("Pauses", "TimeKnK", App.Path + "\KnK.ini")
Clearornot$ = GetFromINI("List", "clearornot", App.Path + "\KnK.ini")
Color$ = GetFromINI("ascii", "Color", App.Path + "\KnK.ini")
adver$ = GetFromINI("Scroll", "adver", App.Path + "\KnK.ini")
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")


If Timer7.Enabled = False Then Exit Sub
For i = 0 To List3.ListCount - 1
AOLChatSend ("/" + Text1 + " Send " + List3.List(i))
If Timer7.Enabled = False Then Exit Sub
If TimeKnK$ = "1" Then
TimeOut (1)
End If
If TimeKnK$ = "2" Then
TimeOut (2)
End If
If TimeKnK$ = "3" Then
TimeOut (3)
End If
If TimeKnK$ = "4" Then
TimeOut (4)
End If
Next i
Timer7.Enabled = False
If adver$ = "yes" Then
AOLChatSend ("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
AOLChatSend ("«-×´¯`°  Done Requesting °´¯`×-»")
TimeOut (0.5)
AOLChatSend ("«-×´¯`° Get a copy at  °´¯`×-»")
TimeOut (0.9)
AOLChatSend ("«-×´¯`° http://knk.tierranet.com/serv")
End If
If adver$ = "no" Then
End If
VDPush2.Enabled = True
If Clearornot$ = "Clear" Then
List3.Clear
End If
If Clearornot$ = "dontclear" Then
Exit Sub
End If
End Sub

Private Sub Timer8_Timer()
If Text1 = "Screen Name" Then
MsgBox "You need to get the Servers SN First!", vbExclamation, "Need to get a SN"
Timer8 = False
Exit Sub
End If
If Option1 = True Then
AOLChatSend ("/" + Text1 + " Send List")
Timer8 = False
End If
If Option2 = True Then
AOLChatSend ("/" + Text1 + " Find " + Text2)
Timer8 = False
End If
If Option3 = True Then
AOLChatSend ("/" + Text1 + " Send Status")
Timer8 = False
End If
Timer8 = False

End Sub

Private Sub Timer9_Timer()
TimeKnK$ = GetFromINI("Pauses", "TimeKnK", App.Path + "\KnK.ini")
Clearornot$ = GetFromINI("List", "clearornot", App.Path + "\KnK.ini")
Color$ = GetFromINI("ascii", "Color", App.Path + "\KnK.ini")
adver$ = GetFromINI("Scroll", "adver", App.Path + "\KnK.ini")
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")


If Timer9.Enabled = False Then Exit Sub
For i = 0 To List3.ListCount - 1
AOLChatSend ("-" + Text1 + " send " + List3.List(i))
If Timer9.Enabled = False Then Exit Sub
If TimeKnK$ = "1" Then
TimeOut (1)
End If
If TimeKnK$ = "2" Then
TimeOut (2)
End If
If TimeKnK$ = "3" Then
TimeOut (3)
End If
If TimeKnK$ = "4" Then
TimeOut (4)
End If
Next i
Timer9.Enabled = False
If adver$ = "yes" Then
AOLChatSend ("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
AOLChatSend ("«-×´¯`°  Done Requesting °´¯`×-»")
TimeOut (0.5)
AOLChatSend ("«-×´¯`° Get a copy at  °´¯`×-»")
TimeOut (0.9)
AOLChatSend ("«-×´¯`° http://knk.tierranet.com/serv")
End If
If adver$ = "no" Then
End If
VDPush2.Enabled = True
If Clearornot$ = "Clear" Then
List3.Clear
End If
If Clearornot$ = "dontclear" Then
Exit Sub
End If

End Sub

Private Sub VDPush1_Click()
Playwav (App.Path + "\blip.WAV")
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")


'Checking if your on
If UserSN = "" Then
MsgBox "You need to be signed on for this to work!", vbInformation, "Please Sign On"
Exit Sub
End If

'AOL95 Functions
If aversion$ = "aol95" Then
If Label12.Caption = "im" Then
Timer12 = True
End If

'Chat Method
If Label12.Caption = "chat" Then
Timer8 = True
End If

'Valkyrie method
If Label12.Caption = "valkyrie" Then
Timer10 = True
End If

'Chat Method
If Label12.Caption = "embrace" Then
Timer16 = True
End If
End If

'AOL 4.o controls
If aversion$ = "aol4" Then

'IM method
If Label12.Caption = "im" Then
Timer6 = True
End If

'Chat Method
If Label12.Caption = "chat" Then
Timer4 = True
End If

'Valkyrie method
If Label12.Caption = "valkyrie" Then
Timer5 = True
End If

'embrace method
If Label12.Caption = "embrace" Then
Timer14 = True
End If
End If
End Sub

Private Sub VDPush2_Click()
Playwav (App.Path + "\blip.WAV")
TimeKnK$ = GetFromINI("Pauses", "TimeKnK", App.Path + "\KnK.ini")
Clearornot$ = GetFromINI("List", "clearornot", App.Path + "\KnK.ini")
Color$ = GetFromINI("ascii", "Color", App.Path + "\KnK.ini")
adver$ = GetFromINI("Scroll", "adver", App.Path + "\KnK.ini")
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")


'Finding if you have put in a SN
If Text1 = "Screen Name" Then
MsgBox "You need to get the Servers SN First!", vbExclamation, "Need to get a SN"
Exit Sub
End If
'Seeing if thers anything to request
If List3.ListCount = 0 Then
MsgBox "You have nothing to request", vbExclamation, "Need Somethingt to Request"
Exit Sub
End If

'AOL95 Commands
If aversion$ = "aol95" Then
VDPush2.Enabled = False

'IM method
If Label12.Caption = "im" Then
If adver$ = "yes" Then
AOLChatSend ("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
AOLChatSend ("«-×´¯`°  Now Requesting °´¯`×-»")
TimeOut (0.5)
AOLChatSend ("«-×´¯`°       IM Method     °´¯`×-»")
TimeOut (2)
End If
If adver$ = "no" Then
End If
Timer11 = True
End If

'Normal chat method
If Label12.Caption = "chat" Then
If adver$ = "yes" Then
AOLChatSend ("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
AOLChatSend ("«-×´¯`°  Now Requesting °´¯`×-»")
TimeOut (0.5)
AOLChatSend ("«-×´¯`°  Normal Method  °´¯`×-»")
End If
If adver$ = "no" Then
End If
TimeOut (2)
Timer7 = True
End If

'Valkyrie method
If Label12.Caption = "valkyrie" Then
If adver$ = "yes" Then
AOLChatSend ("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
AOLChatSend ("«-×´¯`°  Now Requesting °´¯`×-»")
TimeOut (0.5)
AOLChatSend ("«-×´¯`°  válkyrie Method  °´¯`×-»")
End If
If adver$ = "no" Then
End If
TimeOut (2)
Timer9 = True
End If

'embrace method
If Label12.Caption = "embrace" Then
If adver$ = "yes" Then
AOLChatSend ("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
AOLChatSend ("«-×´¯`°  Now Requesting °´¯`×-»")
TimeOut (0.5)
AOLChatSend ("«-×´¯`°  èmbràcè Method  °´¯`×-»")
End If
If adver$ = "no" Then
End If
TimeOut (2)
Timer15 = True
End If
End If

'AOL4.o Commands
If aversion$ = "aol4" Then
VDPush2.Enabled = False

'IM method
If Label12.Caption = "im" Then
If adver$ = "yes" Then
'Black Blue Black ascii
If Color$ = "bbb" Then
SendChat BlackBlueBlack("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
SendChat BlackBlueBlack("«-×´¯`°  Now Requesting °´¯`×-»")
TimeOut (0.5)
SendChat BlackBlueBlack("«-×´¯`°       IM Method     °´¯`×-»")
TimeOut (2)
Timer3 = True
End If

'Black Green Black ascii
If Color$ = "bgb" Then
SendChat BlackGreenBlack("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
SendChat BlackGreenBlack("«-×´¯`°  Now Requesting °´¯`×-»")
TimeOut (0.5)
SendChat BlackGreenBlack("«-×´¯`°       IM Method     °´¯`×-»")
TimeOut (2)
Timer3 = True
End If
End If
If adver$ = "no" Then
TimeOut (2)
Timer3 = True
End If
End If

'Normal chat method
If Label12.Caption = "chat" Then

If adver$ = "yes" Then

'Black blue Black ascii
If Color$ = "bbb" Then
SendChat BlackBlueBlack("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
SendChat BlackBlueBlack("«-×´¯`°  Now Requesting °´¯`×-»")
TimeOut (0.5)
SendChat BlackBlueBlack("«-×´¯`°  Normal Method  °´¯`×-»")
TimeOut (2)
Timer1 = True
End If

'Black green Black ascii
If Color$ = "bgb" Then
SendChat BlackGreenBlack("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
SendChat BlackGreenBlack("«-×´¯`°  Now Requesting °´¯`×-»")
TimeOut (0.5)
SendChat BlackGreenBlack("«-×´¯`°  Normal Method  °´¯`×-»")
TimeOut (2)
Timer1 = True
End If
End If

If adver$ = "no" Then
TimeOut (2)
Timer1 = True
End If
End If

'Valkyrie method
If Label12.Caption = "valkyrie" Then
If adver$ = "yes" Then

'Black blue Black ascii
If Color$ = "bbb" Then
SendChat BlackBlueBlack("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
SendChat BlackBlueBlack("«-×´¯`°  Now Requesting °´¯`×-»")
TimeOut (0.5)
SendChat BlackBlueBlack("«-×´¯`°  válkyrie Method  °´¯`×-»")
TimeOut (2)
Timer2 = True
End If

'Black green Black ascii
If Color$ = "bgb" Then
SendChat BlackGreenBlack("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
SendChat BlackGreenBlack("«-×´¯`°  Now Requesting °´¯`×-»")
TimeOut (0.5)
SendChat BlackGreenBlack("«-×´¯`°  válkyrie Method  °´¯`×-»")
TimeOut (2)
Timer2 = True
End If
End If

If adver$ = "no" Then
TimeOut (2)
Timer2 = True
End If
End If

'embrace method
If Label12.Caption = "embrace" Then
If adver$ = "yes" Then

'Black blue Black ascii
If Color$ = "bbb" Then
SendChat BlackBlueBlack("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
SendChat BlackBlueBlack("«-×´¯`°  Now Requesting °´¯`×-»")
TimeOut (0.5)
SendChat BlackBlueBlack("«-×´¯`°  èmbràcè Method  °´¯`×-»")
TimeOut (2)
Timer13 = True
End If

'Black green Black ascii
If Color$ = "bgb" Then
SendChat BlackGreenBlack("«-×´¯`°  KnK Founders °´¯`×-»")
TimeOut (0.5)
SendChat BlackGreenBlack("«-×´¯`°  Now Requesting °´¯`×-»")
TimeOut (0.5)
SendChat BlackGreenBlack("«-×´¯`°  èmbràcè Method  °´¯`×-»")
TimeOut (2)
Timer13 = True
End If
End If
If adver$ = "no" Then
TimeOut (2)
Timer13 = True
End If
End If
End If
End Sub

Private Sub VDPush3_Click()
Playwav (App.Path + "\blip.WAV")
List3.Clear
End Sub

Private Sub VDPush4_Click()
Playwav (App.Path + "\blip.WAV")
Form7.Width = 5070
Form7.Width = 5050
Form7.Width = 5000
Form7.Width = 4950
Form7.Width = 4900
Form7.Width = 4850
Form7.Width = 4800
Form7.Width = 4750
Form7.Width = 4700
Form7.Width = 4650
Form7.Width = 4600
Form7.Width = 4550
Form7.Width = 4500
Form7.Width = 4450
Form7.Width = 4400
Form7.Width = 4350
Form7.Width = 4300
Form7.Width = 4250
Form7.Width = 4200
Form7.Width = 4150
Form7.Width = 4100
Form7.Width = 4050
Form7.Width = 4000
Form7.Width = 3950
Form7.Width = 3900
Form7.Width = 3850
Form7.Width = 3800
Form7.Width = 3750
Form7.Width = 3700
Form7.Width = 3650
Form7.Width = 3600
Form7.Width = 3550
Form7.Width = 3500
Form7.Width = 3495

End Sub

Private Sub VDPush5_Click()
Playwav (App.Path + "\blip.WAV")
List2.Clear

End Sub

Private Sub VDPush6_Click()
Playwav (App.Path + "\blip.WAV")
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")

'Finding out if your on
If UserSN = "" Then
MsgBox "You need to be signed on for this to work!", vbInformation, "Please Sign On"
Exit Sub
End If
'AOL95 command
If aversion$ = "aol95" Then
List2.Clear
Call AddRoom(List2)
End If

'AOL4.o command
If aversion$ = "aol4" Then
List2.Clear
Call AddRoomToListBox(List2)
End If
End Sub

