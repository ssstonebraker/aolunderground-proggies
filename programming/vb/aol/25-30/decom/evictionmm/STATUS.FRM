VERSION 2.00
Begin Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MassMail Status"
   ClientHeight    =   4848
   ClientLeft      =   4116
   ClientTop       =   2412
   ClientWidth     =   5352
   Height          =   5184
   Icon            =   STATUS.FRX:0000
   Left            =   4068
   LinkTopic       =   "Form2"
   ScaleHeight     =   4848
   ScaleWidth      =   5352
   Top             =   2124
   Width           =   5448
   Begin VBMsg VBMsg1 
      Height          =   420
      Left            =   0
      MessageCount    =   STATUS.FRX:0302
      MessageList     =   STATUS.FRX:0304
      MessageTypes    =   0  'Selected Messages
      PostDefault     =   0   'False
      Top             =   3720
      Width           =   420
   End
   Begin ListBox List3 
      Height          =   420
      Left            =   1200
      TabIndex        =   19
      Top             =   5040
      Width           =   3375
   End
   Begin DBPush DBPush2 
      Alignment       =   2  'Center
      AutoPictureDisabled=   0  'Off
      BackColor       =   &H00C0C0C0&
      BackColorDown   =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "Exit"
      CornerColor     =   &H00000000&
      FocusRect       =   0   'False
      Font3D          =   0  'None
      Font3DDisabled  =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForceDown       =   0   'False
      ForceUp         =   0   'False
      ForeColor       =   &H00E0E0E0&
      ForeColorDisabled=   &H00000000&
      Height          =   255
      Left            =   2160
      Outline         =   -1  'True
      OutlineSize     =   1
      Picture         =   STATUS.FRX:0502
      PictureDown     =   STATUS.FRX:146E
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   17
      Top             =   3360
      TransColor      =   &H00000000&
      Transparent     =   0   'False
      VertAlignment   =   2  'Center
      Width           =   1095
   End
   Begin DBPush DBPush1 
      Alignment       =   2  'Center
      AutoPictureDisabled=   0  'Off
      BackColor       =   &H00C0C0C0&
      BackColorDown   =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "Stop MM"
      CornerColor     =   &H00000000&
      FocusRect       =   0   'False
      Font3D          =   0  'None
      Font3DDisabled  =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForceDown       =   0   'False
      ForceUp         =   0   'False
      ForeColor       =   &H00E0E0E0&
      ForeColorDisabled=   &H00000000&
      Height          =   255
      Left            =   3480
      Outline         =   -1  'True
      OutlineSize     =   1
      Picture         =   STATUS.FRX:23DA
      PictureDown     =   STATUS.FRX:3346
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   12
      Top             =   3360
      TransColor      =   &H00000000&
      Transparent     =   0   'False
      VertAlignment   =   2  'Center
      Width           =   1095
   End
   Begin SSPanel Panel3D1 
      BackColor       =   &H00000000&
      BevelInner      =   1  'Inset
      BevelWidth      =   2
      BorderWidth     =   2
      FloodColor      =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   600
      Outline         =   -1  'True
      TabIndex        =   8
      Top             =   3720
      Width           =   4215
      Begin Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Times New Roman"
         FontSize        =   24
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   4215
      End
      Begin Label Label9 
         BackColor       =   &H00000000&
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Times New Roman"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1200
         TabIndex        =   11
         Top             =   120
         Width           =   2895
      End
      Begin Label Label7 
         BackColor       =   &H00000000&
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Times New Roman"
         FontSize        =   12
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1560
         TabIndex        =   14
         Top             =   480
         Width           =   2535
      End
      Begin Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "Progress:"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Times New Roman"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1335
      End
      Begin Label Label22 
         BackColor       =   &H00000000&
         Caption         =   "Status:"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Times New Roman"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1095
      End
   End
   Begin DBPush DBPush9 
      Alignment       =   2  'Center
      AutoPictureDisabled=   0  'Off
      BackColor       =   &H00C0C0C0&
      BackColorDown   =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "Pause MM"
      CornerColor     =   &H00000000&
      FocusRect       =   0   'False
      Font3D          =   0  'None
      Font3DDisabled  =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForceDown       =   0   'False
      ForceUp         =   0   'False
      ForeColor       =   &H00E0E0E0&
      ForeColorDisabled=   &H00000000&
      Height          =   255
      Left            =   960
      Outline         =   -1  'True
      OutlineSize     =   1
      Picture         =   STATUS.FRX:42B2
      PictureDown     =   STATUS.FRX:521E
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   6
      Top             =   3360
      TransColor      =   &H00000000&
      Transparent     =   0   'False
      VertAlignment   =   2  'Center
      Width           =   1095
   End
   Begin PictureBox Picture2 
      Height          =   2725
      Left            =   2880
      Picture         =   STATUS.FRX:618A
      ScaleHeight     =   2690.11
      ScaleMode       =   0  'User
      ScaleWidth      =   2268
      TabIndex        =   2
      Top             =   360
      Width           =   2295
      Begin ListBox List2 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   2175
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
   End
   Begin PictureBox Picture1 
      Height          =   2725
      Left            =   240
      Picture         =   STATUS.FRX:1A266
      ScaleHeight     =   2690.11
      ScaleMode       =   0  'User
      ScaleWidth      =   2268
      TabIndex        =   0
      Top             =   360
      Width           =   2295
      Begin ListBox List1 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   2175
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
   Begin Label Label17 
      Caption         =   "Label17"
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   5640
      Width           =   495
   End
   Begin Label Label16 
      Caption         =   "Label16"
      Height          =   255
      Left            =   4440
      TabIndex        =   25
      Top             =   5040
      Width           =   735
   End
   Begin Label Label15 
      BackColor       =   &H00000000&
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   24
      Top             =   120
      Width           =   495
   End
   Begin Label Label14 
      Caption         =   "Label14"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   5040
      Width           =   735
   End
   Begin Label Label13 
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   5040
      Width           =   255
   End
   Begin Label Label12 
      Height          =   375
      Left            =   480
      TabIndex        =   21
      Top             =   4920
      Width           =   735
   End
   Begin Label Label10 
      Caption         =   "1"
      Height          =   135
      Left            =   4800
      TabIndex        =   18
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin Label Label2 
      Caption         =   "0"
      Height          =   135
      Left            =   4920
      TabIndex        =   10
      Top             =   3840
      Visible         =   0   'False
      Width           =   135
   End
   Begin Label Label11 
      Caption         =   "0"
      Height          =   135
      Left            =   5160
      TabIndex        =   16
      Top             =   3840
      Visible         =   0   'False
      Width           =   135
   End
   Begin Label Label8 
      Caption         =   "0"
      Height          =   255
      Left            =   4920
      TabIndex        =   15
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Removed:"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin Label Label6 
      BackColor       =   &H00000000&
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "PeePz on MM:"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
End
' STATUS.FRM
Option Explicit
Const mc0074 = 16 ' &H10%
Const mc0080 = 513 ' &H201%
Const mc0086 = 514 ' &H202%
Const mc008C = 9 ' &H9%
Const mc009A = 2 ' &H2%
Const mc00AE = 1036 ' &H40C%
Const mc00D6 = 1031 ' &H407%
Const mc00FC = 12 ' &HC%
Const mc013A = 14 ' &HE%
Const mc0142 = 13 ' &HD%

Sub DBPush1_Click ()
Label4.Caption = "Stopping MassMail"
DBPush1.Enabled = False
If  DBPush9.Caption = "Un-Pause" Then
DBPush9.Caption = "Pause MM"
Label4.Caption = "Un-Paused"
End If
End Sub

Sub DBPush2_Click ()
Unload Me
Form1.Show
Form1.WindowState = 0
End Sub

Sub DBPush9_Click ()
If  DBPush9.Caption = "Pause MM" Then
DBPush9.Caption = "Un-Pause"
Label4.Caption = "PAUSED"
Exit Sub
End If
If  DBPush9.Caption = "Un-Pause" Then
DBPush9.Caption = "Pause MM"
Label4.Caption = "Un-Paused"
Exit Sub
End If
End Sub

Sub Form_Activate ()
sub0193 Me
End Sub

Sub Form_Load ()
Dim l0032 As Integer
Dim l003A As Variant
Dim l0042 As String
Dim l0060 As Variant
Dim l0070 As Variant
' Const mc0074 = 16 ' &H10%
Dim l0078 As Integer
Dim l007A As Integer
Dim l007E As Integer
' Const mc0080 = 513 ' &H201%
Dim l0084 As Integer
' Const mc0086 = 514 ' &H202%
Dim l008A As Integer
' Const mc008C = 9 ' &H9%
Dim l0096 As Integer
Dim l0098 As Integer
' Const mc009A = 2 ' &H2%
Dim l009E As Integer
Dim l00A2 As Variant
Dim l00AC As Integer
' Const mc00AE = 1036 ' &H40C%
Dim l00B0 As Variant
Dim l00C0 As Variant
Dim l00C4 As String
Dim l00C6 As String
Dim l00CE As Integer
Dim l00D0 As Integer
Dim l00D2 As Integer
Dim l00D4 As Integer
' Const mc00D6 = 1031 ' &H407%
Dim l00D8 As Integer
Dim l00DA As Integer
Dim l00DE As Integer
Dim l00E8 As Integer
Dim l00EA As Variant
Dim l00EE As Integer
Dim l00F0 As Integer
Dim l00F2 As Integer
Dim l00F4 As Integer
Dim l00F6 As Integer
Dim l00F8 As Integer
Dim l00FA As String
' Const mc00FC = 12 ' &HC%
Dim l0100 As Integer
Dim l0104 As Integer
Dim l0106 As Integer
Dim l0108 As Integer
Dim l010A As Integer
Dim l010C As Integer
Dim l010E As Integer
Dim l0110 As Variant
Dim l011A As Variant
Dim l011E As Variant
Dim l0122 As Variant
Dim l0126 As Variant
Dim l012A As Variant
Dim l012E As Integer
Dim l0130 As Integer
Dim l0132 As Variant
Dim l0136 As Integer
Dim l0138 As Integer
' Const mc013A = 14 ' &HE%
Dim l013C As Variant
Dim l0140 As String
' Const mc0142 = 13 ' &HD%
Dim l0144 As String
Dim l0146 As Variant
Dim l014A As Variant
Dim l0160 As String
Dim l0162 As Variant
Dim l0166 As Variant
Dim l0170 As String
Dim l0172 As Variant
Dim l0176 As String
Dim l0178 As Integer
Dim l017A As Integer
l0032% = extfn04BA("AOL Frame25", 0&)
VBMsg1.SubClasshWnd = l0032%
l003A = Date
If  options.Check4.Value = 1 Then
l0042$ = options.Text1.Text
Label13.Caption = "Mail One"
End If
Label9.Caption = "Idle"
If  Form1.Option7.Value = True Then
If  Form3.Visible = True Then
Me.Hide
Do
DoEvents
Loop Until Form3.Visible = False
Me.Show
Label9.Caption = "Idle"
End If
End If
sub0193 Me
DBPush1.Visible = True
DBPush9.Visible = True
DBPush2.Visible = False
List1.Clear
If  Form1.List1.ListCount > 0 Then
For l0060 = 0 To Form1.List1.ListCount - 1
List1.AddItem Form1.List1.List(l0060)
Next l0060
GoTo L9E8A
End If
If  form10.List1.ListCount > 0 Then
For l0060 = 0 To form10.List1.ListCount - 1
List1.AddItem form10.List1.List(l0060)
Next l0060
GoTo L9E8A
End If
L9E8A:
Label6.Caption = List1.ListCount
If  options.Check3.Value = 1 Then
Label9.Caption = "Turning IM's Off"
sub0182
options.Hide
End If
If  options.Check1.Value = 1 Then
Label9.Caption = "Closing ChatRoom"
l0070 = fn0101()
l0078% = extfn0398(l0070, mc0074, 0, 0)
options.Hide
End If
If  options.Check2.Value = 1 Then
l007A% = extfn04BA("_AOL_Palette", "America Online Timer")
If  l007A% <> 0 Then
l007E% = extfn0524(l007A%, "_AOL_Button")
l0084% = extfn02DF(l007E%, mc0080, &HD, 0)
l0084% = extfn02DF(l007E%, mc0086, &HD, 0)
End If
options.Hide
End If
Label9.Caption = "Setting Prefs"
sub01E5
Label9.Caption = "Finding Box"
l0032% = extfn04BA("AOL Frame25", 0&)
If  Form1.Option1.Value = True Then
l008A% = extfn0510(l0032%, "New Mail")
gv00E2 = extfn0538(l008A%, mc008C)
End If
If  Form1.Option4.Value = True Then
l008A% = extfn0510(l0032%, "Old Mail")
gv00E2 = extfn0538(l008A%, mc008C)
End If
If  Form1.Option6.Value = True Then
l008A% = extfn0510(l0032%, "Outgoing Mail")
gv00E2 = extfn0538(l008A%, mc008C)
End If
If  Form1.Option5.Value = True Then
l008A% = extfn0510(l0032%, "Incoming FlashMail")
gv00E2 = extfn0538(l008A%, mc008C)
End If
If  Label4.Caption = "PAUSED" Then
Do
DoEvents
Loop Until Label4.Caption = "Un-Paused"
Label4.Caption = ""
End If
If  DBPush1.Enabled = False Then
Label4.Caption = "STOPPED"
DBPush1.Visible = False
DBPush9.Visible = False
DBPush2.Visible = True
Exit Sub: Unload Form2
End If
If  l008A% = 0 Then
Label9.Caption = "Loading Mailbox"
If  Form1.Option1.Value = True Then
Call sub020F(l0032%, "Read &new Mail")
End If
If  Form1.Option4.Value = True Then
Call sub020F(l0032%, "Check Mail You've &Read")
End If
If  Form1.Option6.Value = True Then
Call sub020F(l0032%, "Check Mail You've &Sent")
End If
If  Form1.Option5.Value = True Then
Call sub020F(l0032%, "Read &Incoming Mail")
End If
Do
DoEvents
l0096% = extfn04BA("#32770", "America Online")
If  l0096% <> 0 Then
l0098% = extfn0524(l0096%, "Static")
l009E% = extfn055D(l0098%, mc009A)
l00A2 = fn014A(l009E%)
Label16.Caption = "" + l00A2 + ""
If  Label16.Caption = "You have no unread mail." Or Label16.Caption = "You have no old mail" Or Label16.Caption = "You have no outgoing mail." Then
sub01AD
Form2.WindowState = 1
MsgBox "You have no mail!"
Label16.Caption = ""
Form1.Show
Form1.WindowState = 0
Me.Hide
Exit Sub
End If
End If
l0032% = extfn04BA("AOL Frame25", 0&)
If  Form1.Option1.Value = True Then
l008A% = extfn0510(l0032%, "New Mail")
End If
If  Form1.Option4.Value = True Then
l008A% = extfn0510(l0032%, "Old Mail")
End If
If  Form1.Option6.Value = True Then
l008A% = extfn0510(l0032%, "Outgoing Mail")
End If
If  Form1.Option5.Value = True Then
l008A% = extfn0510(l0032%, "Incoming FlashMail")
End If
Loop Until l008A% <> 0
Label9.Caption = "Waiting...."
sub0274
End If
Label9.Caption = "Counting the Mails"
l0032% = extfn04BA("AOL Frame25", 0&)
l00AC% = extfn0524(l0032%, "_AOL_TREE")
l00B0 = extfn02DF(l00AC%, mc00AE, 0, 0)
If  Form1.Option7.Value = False Then
If  Form1.Text2.Enabled = False Then
Label9.Caption = "Building List"
List3.Clear
For l0060 = 0 To l00B0 - 1
List3.AddItem l0060
Next l0060
Label11.Caption = List3.ListCount
End If
If  Form1.Text2.Enabled = True Then
Label9.Caption = "Building List"
Label11.Caption = "" + Form1.Text2.Text + ""
List3.Clear
For l0060 = 0 To Form1.Text2.Text - 1
List3.AddItem l0060
Next l0060
End If
End If
Label9.Caption = "Starting MM"
For l00C0 = 0 To List3.ListCount - 1
l00C4$ = l00C0 + 1
l00C6$ = List3.ListCount
Label7.Caption = "Mail #" + l00C4$ + " out of " + l00C6$ + ""
If  Form1.Option7.Value = True Then
Label11.Caption = List3.ListCount
End If
If  Label4.Caption = "PAUSED" Then
Do
DoEvents
Loop Until Label4.Caption = "Un-Paused"
Label4.Caption = ""
End If
If  DBPush1.Enabled = False Then
Label4.Caption = "STOPPED"
DBPush1.Visible = False
DBPush9.Visible = False
DBPush2.Visible = True
Exit Sub: Unload Form2
End If
Do
DoEvents
If  Form1.Option1.Value = True Then
l008A% = extfn0510(l0032%, "New Mail")
End If
If  Form1.Option4.Value = True Then
l008A% = extfn0510(l0032%, "Old Mail")
End If
If  Form1.Option6.Value = True Then
l008A% = extfn0510(l0032%, "Outgoing Mail")
End If
If  Form1.Option5.Value = True Then
l008A% = extfn0510(l0032%, "Incoming FlashMail")
End If
Loop Until l008A% <> 0
l00CE% = extfn0510(l0032%, "Read")
l00D0% = extfn0510(l0032%, "Open")
If  Form1.Option1.Value = True Then
l00D2% = extfn0510(l0032%, "Keep as new")
End If
l00D4% = extfn0524(l0032%, "_AOL_TREE")
Label9.Caption = "Opening Mail"
l00D8% = extfn0398(l00D4%, mc00D6, List3.List(l00C0), 0)
sub01B8
sub0269 (.0001#)
sub00A9 (l00CE%)
sub00A9 (l00D0%)
sub00A9 (l00CE%)
sub00A9 (l00D0%)
sub01B8
If  Label4.Caption = "PAUSED" Then
Do
DoEvents
Loop Until Label4.Caption = "Un-Paused"
Label4.Caption = ""
End If
If  DBPush1.Enabled = False Then
Label4.Caption = "STOPPED"
DBPush1.Visible = False
DBPush9.Visible = False
DBPush2.Visible = True
Exit Sub: Unload Form2
End If
If  Form1.Option1.Value = True Then
Label9.Caption = "Keeping as New"
sub00A9 (l00D2%)
End If
Label9.Caption = "Stopping Text"
Call sub020F(l0032%, "Stop incoming Text")
Do
DoEvents
l00DA% = extfn0510(l0032%, "Forward")
Loop Until l00DA% <> 0
sub01B8
Label9.Caption = "Clicking Forward"
l009E% = extfn055D(l00DA%, mc009A)
l00DE% = extfn0474(l009E%)
l0060 = fn0158(l00DE%)
Label12 = "" & l0060 & ""
sub00ED (.1#)
If  Label4.Caption = "PAUSED" Then
Do
DoEvents
Loop Until Label4.Caption = "Un-Paused"
Label4.Caption = ""
End If
If  DBPush1.Enabled = False Then
Label4.Caption = "STOPPED"
DBPush1.Visible = False
DBPush9.Visible = False
DBPush2.Visible = True
Exit Sub: Unload Form2
End If
sub00A9 (l009E%)
sub00ED (.08#)
Label9.Caption = "Closing"
l00E8% = extfn0474(l009E%)
l00EA = extfn0398(l00E8%, mc0074, 0, 0)
sub00ED (.04#)
Label9.Caption = "Setting Mail Text"
Do
DoEvents
l00EE% = extfn0510(l0032%, "Send Now")
l00AC% = extfn0474(l00EE%)
Loop Until l00AC% <> 0
l00F0% = extfn0524(l00AC%, "_AOL_Edit")
l00F4% = extfn0474(l00F2%)
l00F6% = extfn0524(l00AC%, "_AOL_Icon")
l00F8% = extfn0524(l00AC%, "_AOL_View")
l00F0% = extfn0524(l00AC%, "_AOL_Edit")
l00FA$ = ""
For gv00E2 = 0 To List1.ListCount - 1
l00FA$ = l00FA$ + "(" + List1.List(gv00E2) + "), "
Next gv00E2
If  Label4.Caption = "PAUSED" Then
Do
DoEvents
Loop Until Label4.Caption = "Un-Paused"
Label4.Caption = ""
End If
If  DBPush1.Enabled = False Then
Label4.Caption = "STOPPED"
DBPush1.Visible = False
DBPush9.Visible = False
DBPush2.Visible = True
Exit Sub: Unload Form2
End If
l0100% = extfn02F3(l00F0%, mc00FC, 0, l00FA$)
l0104% = extfn04C8(l00F0%, mc009A)
l0106% = extfn04C8(l0104%, mc009A)
l0108% = extfn04C8(l0106%, mc009A)
l010A% = extfn04C8(l0108%, mc009A)
l010C% = extfn04C8(l010A%, mc009A)
l0100% = extfn02F3(l010A%, mc00FC, 0, Label12)
l010E% = extfn0524(l00AC%, "RICHCNTL")
If  l010E% <> 0 Then
l0110 = Chr(13) + Chr(13) + Form1.Text1.Text
l011A = (Label10 / Label11) * 100
l011E = Chr(13) + "²º·‚¸‚•²º·‚¸‚• UPSS MMer by: EvicTioN " + Chr(13)
l0122 = "²º·‚¸‚•²º·‚¸‚• This Is mail " + Label10 + " out of " + Label11 + "" + Chr(13)
l0126 = "²º·‚¸‚•²º·‚¸‚• There is " + Label6.Caption + " People on this Mass Mail" + Chr(13)
l012A = "²º·‚¸‚•²º·‚¸‚• This Mass Mail is " + l011A + "% complete"
l0100% = extfn02F3(l010E%, mc00FC, 0, "" + l011E + "" + l0122 + "" + l0126 + "" + l012A + "" + l0110 + "")
End If
If  l010C% <> 0 Then
l0110 = Chr(13) + Chr(13) + Form1.Text1.Text
l011A = (Label10 / Label11) * 100
l011E = Chr(13) + "²º·‚¸‚•²º·‚¸‚• UPSS MMer by: EvicTioN" + Chr(13)
l0122 = "²º·‚¸‚•²º·‚¸‚• This Is mail " + Label10 + " out of " + Label11 + "" + Chr(13)
l0126 = "²º·‚¸‚•²º·‚¸‚• There is " + Label6.Caption + " People on this Mass Mail" + Chr(13)
l012A = "²º·‚¸‚•²º·‚¸‚• This Mass Mail is " + l011A + "% complete"
l0100% = extfn02F3(l010C%, mc00FC, 0, "" + l011E + "" + l0122 + "" + l0126 + "" + l012A + "" + l0110 + "")
End If
Label9.Caption = "Clicking Send"
sub01B8
sub0269 (.06#)
sub00A9 (l00F6%)
sub00A9 (l00F6%)
If  Label4.Caption = "PAUSED" Then
Do
DoEvents
Loop Until Label4.Caption = "Un-Paused"
Label4.Caption = ""
End If
If  DBPush1.Enabled = False Then
Label4.Caption = "STOPPED"
DBPush1.Visible = False
DBPush9.Visible = False
DBPush2.Visible = True
Exit Sub: Unload Form2
End If
Do
DoEvents
If  Label4.Caption = "PAUSED" Then
Do
DoEvents
Loop Until Label4.Caption = "Un-Paused"
Label4.Caption = ""
End If
If  DBPush1.Enabled = False Then
Label4.Caption = "STOPPED"
DBPush1.Visible = False
DBPush9.Visible = False
DBPush2.Visible = True
Exit Sub: Unload Form2
End If
l00AC% = extfn0474(l00EE%)
sub00ED (.2#)
l0096% = extfn04BA("#32770", "America Online")
If  l0096% <> 0 Then
Label9.Caption = "Closing Error"
l012E% = extfn0524(l0096%, "Static")
l0130% = extfn0524(l0096%, "Button")
l0132 = fn014A(l0130%)
l0132 = Label16.Caption
If  Label16.Caption = "Sorry, we are temporarily unable to send mail.  Please try again in a few minutes." Or l0132 = "Sorry, we are temporarily unable to send mail.  Please try again in a few minutes. " Or l0132 = "Sorry, we are temporarily unable to send mail.  Please try again later." Then
sub01AD
Label9.Caption = "Clicking Send"
sub01B8
sub0269 (.06#)
sub00A9 (l00F6%)
sub00A9 (l00F6%)
Label16.Caption = ""
End If
End If
l0136% = extfn0510(l0032%, "Error")
If  l0136% <> 0 Then
Label9.Caption = "Removing Names"
Do
DoEvents
l0138% = extfn0524(l0136%, "_AOL_View")
l013C = extfn0398(l0138%, mc013A, 0&, 0&)
l0140$ = Space$(l013C)
l0060 = extfn02F3(l0138%, mc0142, l013C + 1, l0140$)
l0140$ = Trim$(l0140$)
sub0269 (1.5#)
l0138% = extfn0524(l0136%, "_AOL_View")
l013C = extfn0398(l0138%, mc013A, 0&, 0&)
l0144$ = Space$(l013C)
l0060 = extfn02F3(l0138%, mc0142, l013C + 1, l0144$)
Loop Until l0140$ = Trim$(l0144$)
LC4A6:
For l0146 = 0 To List1.ListCount - 1
l014A = InStr(l0140$, List1.List(l0146))
If  l014A Then
List2.AddItem List1.List(l0146)
List1.RemoveItem l0146
Label6.Caption = List1.ListCount
GoTo LC4A6
End If
Next l0146
Do
DoEvents
l0136% = extfn0510(l0032%, "Error")
gv00E2% = extfn0398(l0136%, mc0074, 0, 0)
sub0269 (.4#)
Loop While l0136% <> 0
If  List1.ListCount = 0 Then
MsgBox "No More people left on MassMail"
DBPush1.Visible = False
DBPush9.Visible = False
DBPush2.Visible = True
Exit Sub
End If
l00FA$ = ""
For gv00E2 = 0 To List1.ListCount - 1
l00FA$ = l00FA$ + "(" + List1.List(gv00E2) + "), "
Next gv00E2
Label15.Caption = List2.ListCount
l0100% = extfn02F3(l00F0%, mc00FC, 0, l00FA$)
l00F6% = extfn0524(l00AC%, "_AOL_Icon")
Label9.Caption = "Clicking Send"
sub01B8
sub0269 (.07#)
sub00A9 (l00F6%)
sub00A9 (l00F6%)
End If
If  Label4.Caption = "PAUSED" Then
Do
DoEvents
Loop Until Label4.Caption = "Un-Paused"
Label4.Caption = ""
End If
If  DBPush1.Enabled = False Then
Label4.Caption = "STOPPED"
DBPush1.Visible = False
DBPush9.Visible = False
DBPush2.Visible = True
Exit Sub: Unload Form2
End If
Loop While l00AC% <> 0
Label9.Caption = "Mail Sent"
Label2 = Label2 + 1
Label10 = Label10 + 1
sub00ED (.1#)
sub01B8
If  Label4.Caption = "PAUSED" Then
Do
DoEvents
Loop Until Label4.Caption = "Un-Paused"
Label4.Caption = ""
End If
If  DBPush1.Enabled = False Then
Label4.Caption = "STOPPED"
DBPush1.Visible = False
DBPush9.Visible = False
DBPush2.Visible = True
Exit Sub: Unload Form2
End If
Label9.Caption = "Getting Next"
If  Label13.Caption = "Mail One" Then
If  options.Check4.Value = 1 Then
l0160$ = l0042$
l0162 = options.Text2.Text
l0166 = options.Text3.Text
Label14 = "" + l003A + ""
l0170$ = "" & l0162 & " " & l0166 & " " & l003A & " FirstMail " & Label11.Caption & " Mails"
l0172 = "This was done by the " + Chr(34) + "UPSS MMER by: EvicTioN" + Chr(34) + ""
l0176$ = Label12 + Chr(13) + Chr(13) + l00FA$ + Chr(13) + "" + l011E + "" + l0122 + "" + l0126 + "" + l012A + "" + l0110 + "" + Chr(13) + "" + Chr(13) + "" + l0172 + ""
sub0234 l0160$, l0170$, l0176$
Do
DoEvents
l0032% = extfn04BA("AOL FRAME25", 0&)
l0178% = extfn0524(l0032%, "MDIClient")
l017A% = extfn0510(l0178%, "Compose Mail")
Loop Until l017A% = 0
Label13.Caption = ""
End If
End If
Label13 = "" + l011A + ""
If  Label13.Caption = "100" Then
If  options.Check4.Value = 1 Then
l0160$ = l0042$
l0162 = options.Text2.Text
l0166 = options.Text3.Text
Label14 = "" + l003A + ""
l0170$ = "" & l0162 & " MM" & l0166 & " " & l003A & " LastMail " & Label11.Caption & " Mails"
l0172 = "This was done by the " + Chr(34) + "UPSS MMER by: EvicTioN" + Chr(34) + ""
l0176$ = Label12 + Chr(13) + Chr(13) + l00FA$ + Chr(13) + "" + l011E + "" + l0122 + "" + l0126 + "" + l012A + "" + l0110 + "" + Chr(13) + "" + Chr(13) + "" + l0172 + ""
sub0234 l0160$, l0170$, l0176$
Do
DoEvents
l0032% = extfn04BA("AOL FRAME25", 0&)
l0178% = extfn0524(l0032%, "MDIClient")
l017A% = extfn0510(l0178%, "Compose Mail")
Loop Until l017A% = 0
End If
Form3.List1.Clear
Unload Form3
Form1.List1.Clear
For l0060 = 0 To Form2.List1.ListCount - 1
Form1.List1.AddItem Form2.List1.List(l0060)
Next l0060
Unload form10
MsgBox "Massmail is complete!"
DBPush1.Visible = False
DBPush9.Visible = False
DBPush2.Visible = True
Exit Sub
End If
Next
Label4.Caption = "Mass Mail Complete!"
DBPush1.Visible = False
DBPush9.Visible = False
DBPush2.Visible = True
End Sub

Sub Form_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
sub0193 Me
End Sub

Sub Form_Unload (Cancel As Integer)
Form1.Show
Form1.WindowState = 0
End Sub

Sub List1_DblClick ()
List1.RemoveItem List1.ListIndex
Label6.Caption = List1.ListCount
End Sub

Sub VBMsg1_WindowMessage (hWindow As Integer, Msg As Integer, wParam As Integer, lParam As Long, RetVal As Long, CallDefProc As Integer)
Dim l0192 As Integer
l0192% = extfn04BA("AOL Frame25", 0&)
If  l0192% = 0 Then
Label4.Caption = "paused"
MsgBox "AOL has either been closed, or has errored. MM is paused"
VBMsg1.SubClasshWnd = 0
End If
End Sub
