VERSION 2.00
Begin Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UPSS Mass Mailer by: EvicTioN"
   ClientHeight    =   4092
   ClientLeft      =   3228
   ClientTop       =   1920
   ClientWidth     =   4704
   Height          =   4428
   Icon            =   MAINMMER.FRX:0000
   Left            =   3180
   LinkTopic       =   "Form1"
   ScaleHeight     =   4092
   ScaleWidth      =   4704
   Top             =   1632
   Width           =   4800
   Begin TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   28
      Text            =   "300"
      Top             =   1080
      Width           =   495
   End
   Begin Timer Timer1 
      Interval        =   500
      Left            =   1680
      Top             =   120
   End
   Begin DBPush DBPush4 
      Alignment       =   2  'Center
      AutoPictureDisabled=   0  'Off
      BackColor       =   &H00C0C0C0&
      BackColorDown   =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "Exit MMer"
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
      Left            =   480
      Outline         =   -1  'True
      OutlineSize     =   1
      Picture         =   MAINMMER.FRX:0302
      PictureDown     =   MAINMMER.FRX:126E
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   21
      Top             =   1680
      TransColor      =   &H00000000&
      Transparent     =   0   'False
      VertAlignment   =   2  'Center
      Width           =   1095
   End
   Begin SSPanel Panel3D1 
      BevelOuter      =   0  'None
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   600
      Width           =   1335
      Begin OptionButton Option7 
         BackColor       =   &H00000000&
         Caption         =   "Select Mails"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   27
         Top             =   240
         Width           =   1335
      End
      Begin OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "Select #:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   20
         Top             =   480
         Width           =   1335
      End
      Begin OptionButton Option3 
         BackColor       =   &H00000000&
         Caption         =   "All Mails"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin OptionButton Option6 
      BackColor       =   &H00000000&
      Caption         =   "Sent Mail"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   4920
      Width           =   1215
   End
   Begin OptionButton Option5 
      BackColor       =   &H00000000&
      Caption         =   "Flash Mail"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   360
      Width           =   1215
   End
   Begin OptionButton Option4 
      BackColor       =   &H00000000&
      Caption         =   "Old Mail"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   4680
      Width           =   1215
   End
   Begin OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "New Mail"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   120
      Value           =   -1  'True
      Width           =   1215
   End
   Begin DBPush DBPush9 
      Alignment       =   2  'Center
      AutoPictureDisabled=   0  'Off
      BackColor       =   &H00C0C0C0&
      BackColorDown   =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "Start MM"
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
      Picture         =   MAINMMER.FRX:21DA
      PictureDown     =   MAINMMER.FRX:3146
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   13
      Top             =   2160
      TransColor      =   &H00000000&
      Transparent     =   0   'False
      VertAlignment   =   2  'Center
      Width           =   1095
   End
   Begin DBPush DBPush5 
      Alignment       =   2  'Center
      AutoPictureDisabled=   0  'Off
      BackColor       =   &H00C0C0C0&
      BackColorDown   =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "Options"
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
      Left            =   720
      Outline         =   -1  'True
      OutlineSize     =   1
      Picture         =   MAINMMER.FRX:40B2
      PictureDown     =   MAINMMER.FRX:501E
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   12
      Top             =   1920
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
      Caption         =   "MM Bot"
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
      Left            =   240
      Outline         =   -1  'True
      OutlineSize     =   1
      Picture         =   MAINMMER.FRX:5F8A
      PictureDown     =   MAINMMER.FRX:6EF6
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   11
      Top             =   1440
      TransColor      =   &H00000000&
      Transparent     =   0   'False
      VertAlignment   =   2  'Center
      Width           =   1095
   End
   Begin TextBox Text1 
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   9.6
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   2880
      Width           =   1935
   End
   Begin CommonDialog CMDialog1 
      Left            =   1800
      Top             =   4080
   End
   Begin DBPush DBPush2 
      Alignment       =   2  'Center
      AutoPictureDisabled=   0  'Off
      BackColor       =   &H00C0C0C0&
      BackColorDown   =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "Add Room"
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
      Left            =   3360
      Outline         =   -1  'True
      OutlineSize     =   1
      Picture         =   MAINMMER.FRX:7E62
      PictureDown     =   MAINMMER.FRX:8DCE
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   6
      Top             =   3240
      TransColor      =   &H00000000&
      Transparent     =   0   'False
      VertAlignment   =   2  'Center
      Width           =   1095
   End
   Begin DBPush DBPush3 
      Alignment       =   2  'Center
      AutoPictureDisabled=   0  'Off
      BackColor       =   &H00C0C0C0&
      BackColorDown   =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "Clear List"
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
      Left            =   2760
      Outline         =   -1  'True
      OutlineSize     =   1
      Picture         =   MAINMMER.FRX:9D3A
      PictureDown     =   MAINMMER.FRX:ACA6
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   5
      Top             =   3480
      TransColor      =   &H00000000&
      Transparent     =   0   'False
      VertAlignment   =   2  'Center
      Width           =   1095
   End
   Begin DBPush DBPush6 
      Alignment       =   2  'Center
      AutoPictureDisabled=   0  'Off
      BackColor       =   &H00C0C0C0&
      BackColorDown   =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "Load List"
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
      Left            =   3360
      Outline         =   -1  'True
      OutlineSize     =   1
      Picture         =   MAINMMER.FRX:BC12
      PictureDown     =   MAINMMER.FRX:CB7E
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   4
      Top             =   3720
      TransColor      =   &H00000000&
      Transparent     =   0   'False
      VertAlignment   =   2  'Center
      Width           =   1095
   End
   Begin DBPush DBPush7 
      Alignment       =   2  'Center
      AutoPictureDisabled=   0  'Off
      BackColor       =   &H00C0C0C0&
      BackColorDown   =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "Save List"
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
      Left            =   2280
      Outline         =   -1  'True
      OutlineSize     =   1
      Picture         =   MAINMMER.FRX:DAEA
      PictureDown     =   MAINMMER.FRX:EA56
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   3
      Top             =   3720
      TransColor      =   &H00000000&
      Transparent     =   0   'False
      VertAlignment   =   2  'Center
      Width           =   1095
   End
   Begin DBPush DBPush8 
      Alignment       =   2  'Center
      AutoPictureDisabled=   0  'Off
      BackColor       =   &H00C0C0C0&
      BackColorDown   =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "Add Name"
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
      Left            =   2280
      Outline         =   -1  'True
      OutlineSize     =   1
      Picture         =   MAINMMER.FRX:F9C2
      PictureDown     =   MAINMMER.FRX:1092E
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   2
      Top             =   3240
      TransColor      =   &H00000000&
      Transparent     =   0   'False
      VertAlignment   =   2  'Center
      Width           =   1095
   End
   Begin PictureBox Picture1 
      Height          =   2725
      Left            =   2160
      Picture         =   MAINMMER.FRX:1189A
      ScaleHeight     =   2690.11
      ScaleMode       =   0  'User
      ScaleWidth      =   2268
      TabIndex        =   0
      Top             =   480
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
   Begin Label Label8 
      Caption         =   "0"
      Height          =   255
      Left            =   1680
      TabIndex        =   26
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin Label Label10 
      Caption         =   "1"
      Height          =   135
      Left            =   1560
      TabIndex        =   25
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin Label Label11 
      Caption         =   "0"
      Height          =   135
      Left            =   1920
      TabIndex        =   24
      Top             =   720
      Visible         =   0   'False
      Width           =   135
   End
   Begin Label Label12 
      Height          =   135
      Left            =   1800
      TabIndex        =   23
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin Label Label2 
      Caption         =   "0"
      Height          =   135
      Left            =   1800
      TabIndex        =   22
      Top             =   720
      Visible         =   0   'False
      Width           =   135
   End
   Begin Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Message:"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
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
      Left            =   2160
      TabIndex        =   8
      Top             =   240
      Width           =   1815
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
      Left            =   3960
      TabIndex        =   7
      Top             =   240
      Width           =   495
   End
End
' MAINMMER.FRM
Option Explicit
Const mc0092 = 513 ' &H201%

Sub DBPush1_Click ()
Dim l0028 As Variant
Form1.WindowState = 1
form10.Show
If  Text2.Enabled = True Then
form10.Label11.Caption = Text2.Text
End If
form10.List1.Clear
For l0028 = 0 To List1.ListCount - 1
form10.List1.AddItem List1.List(l0028)
Next l0028
List1.Clear
End Sub

Sub DBPush2_Click ()
Dim l0032 As Integer
l0032% = fn0101()
If  l0032% = 0 Then
MsgBox "Must be in a chat to use"
Exit Sub
End If
sub0095 List1
Label6.Caption = List1.ListCount
End Sub

Sub DBPush3_Click ()
Dim l003C As Variant
l003C = MsgBox("Are you sure you want to Clear the list?", 4, "Clear MM list")
If  l003C = 6 Then
List1.Clear
Exit Sub
End If
If  l003C = 7 Then Exit Sub
End Sub

Sub DBPush4_Click ()
Dim l0040 As Variant
l0040 = MsgBox("Are you sure you want Exit?", 4, "Exit UPSS")
If  l0040 = 6 Then
End
End If
If  l0040 = 7 Then Exit Sub
End Sub

Sub DBPush5_Click ()
options.Show
End Sub

Sub DBPush6_Click ()
Dim l0048 As Variant
Dim l004C As String
Dim l004E As Integer
On Error Resume Next
CMDialog1.Flags = &H1000& Or &H4& Or &H800& Or &H4000&
CMDialog1.DefaultExt = "MMN"
CMDialog1.DialogTitle = "Load Name List"
CMDialog1.Filter = "MM List (*.MML)|*.MML|Text Files (*.TXT)|*.TXT|All Files (*.*)|*.*|"
CMDialog1.MaxFileSize = 2000
CMDialog1.Filename = ""
CMDialog1.CancelError = True
CMDialog1.Action = 1
Select Case Err
    Case 0:
    List1.Clear
    Open CMDialog1.Filename For Input As #1
l0048 = 0
    Do
    Input #1, l004C$
    l004C$ = Trim(l004C$)
    If  l004C$ <> "" Then List1.AddItem l004C$
    l004E% = DoEvents()
l0048 = l0048 + 1
    Loop Until EOF(1)
  Close #1
End Select
Label6.Caption = List1.ListCount
Exit Sub
End Sub

Sub DBPush7_Click ()
Dim l0056 As Variant
Dim l005A As Variant
Dim l005E As String
Dim l0060 As String
If  List1.ListCount = 0 Then MsgBox "There are no names.", 48: Exit Sub
On Error Resume Next
CMDialog1.Flags = &H4& Or &H2& Or &H800&
CMDialog1.DefaultExt = "MMN"
CMDialog1.DialogTitle = "Save Names as ..."
CMDialog1.Filter = "MM List (*.MML)|*.MML|Text Files (*.TXT)|*.TXT|All Files (*.*)|*.*|"
CMDialog1.MaxFileSize = 2000
CMDialog1.Filename = ""
CMDialog1.CancelError = True
CMDialog1.Action = 2
Select Case Err
    Case 0:
    l0056 = List1.ListCount - 1
    For l005A = 0 To l0056 - 1
    l005E$ = List1.List(l005A)
    l0060$ = l0060$ + l005E$ & ","
    Next l005A
    l0060$ = l0060$ + List1.List(l0056)
        Open CMDialog1.Filename For Output As #1
            Print #1, l0060$
        Close #1
    Case 32755:
        Err = False
    Case Else:
        MsgBox "Unexpected error:" & Str(Err) & Chr$(13) & "Error message: " & Error$(Err), 16, "Unexpected Error!!!"
End Select
End Sub

Sub DBPush8_Click ()
Dim l0062 As Variant
Dim l0066 As Variant
l0066 = InputBox("Screen Name Here", "Add who?")
If  l0066 = "" Then Exit Sub
For l0062 = 0 To List1.ListCount - 1
If  UCase$(l0066) = UCase$(List1.List(l0062)) Then Exit Sub
Next l0062
List1.AddItem l0066
Label6.Caption = List1.ListCount
End Sub

Sub DBPush9_Click ()
Dim l006C As Integer
Dim l0070 As Integer
l006C% = extfn04BA("AOL FRAME25", 0&)
l0070% = extfn0510(l006C%, "Welcome, ")
If  l0070% = 0 Then
MsgBox "Not signed on!"
Exit Sub
End If
If  Text2.Enabled = True And Text2.Text = "" Then
MsgBox "enter a real # of mails to send"
Exit Sub
End If
If  Option7.Value = True Then
If  List1.ListCount = 0 Then
MsgBox "get some peeps on the mm first"
Exit Sub
End If
Form3.Show
Exit Sub
End If
If  List1.ListCount = 0 Then
MsgBox "get some peeps on the mm first"
Exit Sub
End If
Form1.Visible = False
Load Form2
End Sub

Sub Form_Load ()
Dim l0082 As Variant
sub0193 Me
l0082 = fn0141()
If  l0082 = "" Then
Text1.Text = "UPSS MM. Enjoy the free WaReZ"
options.Show
Exit Sub
End If
Text1.Text = "UPSS MM by: " + l0082 + ". Enjoy the free WaReZ!"
options.Show
MsgBox "This release of my massmailer werks great. There is only a few bugs left!"
End Sub

Sub List1_DblClick ()
List1.RemoveItem List1.ListIndex
Label6.Caption = List1.ListCount
End Sub

Sub Option2_Click ()
Text2.Enabled = True
End Sub

Sub Option3_Click ()
Text2.Enabled = False
End Sub

Sub Option7_Click ()
Text2.Enabled = False
End Sub

Sub Timer1_Timer ()
Dim l008C As Integer
Dim l0090 As Integer
' Const mc0092 = 513 ' &H201%
Dim l0096 As Integer
' Const mc0098 = 514 ' &H202%
l008C% = extfn04BA("_AOL_Palette", "America Online Timer")
l0090% = extfn0524(l008C%, "_AOL_Button")
l0096% = extfn02DF(l0090%, mc0092, &HD, 0)
l0096% = extfn02DF(l0090%, mc0098, &HD, 0)
End Sub
