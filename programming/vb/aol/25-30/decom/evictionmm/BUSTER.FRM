VERSION 2.00
Begin Form Form12 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Room Bust"
   ClientHeight    =   2208
   ClientLeft      =   5832
   ClientTop       =   2988
   ClientWidth     =   2556
   Height          =   2544
   Icon            =   BUSTER.FRX:0000
   Left            =   5784
   LinkTopic       =   "Form1"
   ScaleHeight     =   2208
   ScaleWidth      =   2556
   Top             =   2700
   Width           =   2652
   Begin SSCommand Command3D1 
      Caption         =   "Command3D1"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2400
      TabIndex        =   15
      Top             =   1800
      Width           =   135
   End
   Begin OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "Menu"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   12
      Top             =   840
      Width           =   855
   End
   Begin OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Keyword"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   840
      Width           =   975
   End
   Begin Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1200
      Top             =   2760
   End
   Begin Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   2760
   End
   Begin DBPush DBPush4 
      Alignment       =   2  'Center
      AutoPictureDisabled=   0  'Off
      BackColor       =   &H00C0C0C0&
      BackColorDown   =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "Kill Wait"
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
      Picture         =   BUSTER.FRX:0302
      PictureDown     =   BUSTER.FRX:126E
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   7
      Top             =   1320
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
      Left            =   1320
      Outline         =   -1  'True
      OutlineSize     =   1
      Picture         =   BUSTER.FRX:21DA
      PictureDown     =   BUSTER.FRX:3146
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   6
      Top             =   1320
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
      Caption         =   "Bust"
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
      Picture         =   BUSTER.FRX:40B2
      PictureDown     =   BUSTER.FRX:501E
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   5
      Top             =   1080
      TransColor      =   &H00000000&
      Transparent     =   0   'False
      VertAlignment   =   2  'Center
      Width           =   1095
   End
   Begin DBPush DBPush2 
      Alignment       =   2  'Center
      AutoPictureDisabled=   0  'Off
      BackColor       =   &H00C0C0C0&
      BackColorDown   =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "Stop"
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
      Left            =   1320
      Outline         =   -1  'True
      OutlineSize     =   1
      Picture         =   BUSTER.FRX:5F8A
      PictureDown     =   BUSTER.FRX:6EF6
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   4
      Top             =   1080
      TransColor      =   &H00000000&
      Transparent     =   0   'False
      VertAlignment   =   2  'Center
      Width           =   1095
   End
   Begin TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin VBMsg VBMsg1 
      Height          =   420
      Left            =   0
      MessageCount    =   BUSTER.FRX:7E62
      MessageList     =   BUSTER.FRX:7E64
      MessageTypes    =   0  'Selected Messages
      PostDefault     =   0   'False
      Top             =   3120
      Width           =   420
   End
   Begin Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "Bust attempts:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   1920
      Width           =   1335
   End
   Begin Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "Bust Method:"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   600
      Width           =   1815
   End
   Begin Label Label6 
      Caption         =   "Label6"
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   3120
      Width           =   135
   End
   Begin Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   3720
      Width           =   255
   End
   Begin Label Label4 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   1920
      Width           =   735
   End
   Begin Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "What room?"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin Label Label2 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   1680
      Width           =   1095
   End
   Begin Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Status:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   735
   End
End
' BUSTER.FRM
Option Explicit
Const mc0050 = 12 ' &HC%
Const mc0058 = 258 ' &H102%
Const mc0096 = 2 ' &H2%
Const mc00A0 = 513 ' &H201%
Const mc00A4 = 514 ' &H202%
Const mc00AA = 16 ' &H10%

Sub Command3D1_Click ()
Dim l0028 As Integer
Dim l002C As Integer
Dim l002E As Integer
Dim l0048 As String
Dim l004C As Integer
Dim l004E As Integer
' Const mc0050 = 12 ' &HC%
Dim l0054 As Integer
Dim l0056 As Integer
' Const mc0058 = 258 ' &H102%
Dim l005C As Variant
Label4.Caption = "0"
Label6.Caption = "0"
l0028% = extfn04BA("AOL Frame25", 0&)
l002C% = extfn0524(l0028%, "MDIClient")
l002E% = extfn04BA("#32770", 0&)
VBMsg1.SubClasshWnd = l002C%
Timer1.Enabled = True
Label2.Caption = "Busting!"
l0048$ = "" & Text1.Text & ""
l004C% = extfn0510(l0028%, "keyword")
If  l004C% = 0 Then
Call sub020F(l0028%, "keyword...")
End If
Do
DoEvents
l004C% = extfn0510(l0028%, "keyword")
Loop Until l004C% <> 0
Do
DoEvents
l004E% = extfn0524(l004C%, "_AOL_EDIT")
Loop Until l004E% <> 0
l0054% = extfn02F3(l004E%, mc0050, 0, "aol://2719:2-2-" + l0048$ + "")
Do
DoEvents
l0056% = extfn0524(l004C%, "_AOL_ICON")
Loop Until l0056% <> 0
Do
DoEvents
l0056% = extfn0524(l004C%, "_AOL_ICON")
If  Label2.Caption = "Stopped" Then
Exit Sub
End If
l005C = extfn02DF(l0056%, mc0058, 13, 0)
If  Label2.Caption = "Busting!" Then
Label4 = Label4 + 1
End If
sub01B8
Loop Until Label2.Caption = "stopped"
End Sub

Sub DBPush1_Click ()
Dim l0060 As Integer
Dim l0062 As Integer
Dim l0070 As Integer
Dim l0072 As Integer
Dim l007A As String
Dim l007C As Integer
Dim l007E As Integer
Dim l0080 As Integer
Dim l0082 As Integer
Dim l0084 As Variant
Dim l008A As Variant
Dim l008E As Variant
Dim l0092 As Integer
Dim l0094 As Integer
' Const mc0096 = 2 ' &H2%
Dim l009C As Integer
Dim l009E As Integer
' Const mc00A0 = 513 ' &H201%
Dim l00A2 As Integer
' Const mc00A4 = 514 ' &H202%
Dim l00A6 As Variant
' Const mc00AA = 16 ' &H10%
Dim l00AE As Integer
l0060% = extfn04BA("AOL FRAME25", 0&)
l0062% = extfn0510(l0060%, "Welcome, ")
If  l0062% = 0 Then
MsgBox "Not signed on!"
Exit Sub
End If
If  Option1.Value = False And Option2.Value = False Then
MsgBox "choose a method"
Exit Sub
End If
If  Option1.Value = True Then
Label4.Caption = "0"
Label6.Caption = "0"
l0060% = extfn04BA("AOL Frame25", 0&)
l0070% = extfn0524(l0060%, "MDIClient")
l0072% = extfn04BA("#32770", 0&)
VBMsg1.SubClasshWnd = l0070%
Timer1.Enabled = True
Timer2.Enabled = True
Label2.Caption = "Busting!"
l007A$ = "" & Text1.Text & ""
l007C% = extfn0510(l0060%, "keyword")
If  l007C% = 0 Then
Call sub020F(l0060%, "keyword...")
End If
Do
DoEvents
l007C% = extfn0510(l0060%, "keyword")
Loop Until l007C% <> 0
Do
DoEvents
l007E% = extfn0524(l007C%, "_AOL_EDIT")
Loop Until l007E% <> 0
l0080% = extfn02F3(l007E%, mc0050, 0, "aol://2719:2-2-" + l007A$ + "")
Do
DoEvents
l0082% = extfn0524(l007C%, "_AOL_ICON")
Loop Until l0082% <> 0
Do
DoEvents
l0082% = extfn0524(l007C%, "_AOL_ICON")
If  Label2.Caption = "Stopped" Then
Exit Sub
End If
l0084 = extfn02DF(l0082%, mc0058, 13, 0)
If  Label2.Caption = "Busting!" Then
Label4 = Label4 + 1
End If
l008A = fn00CD()
sub01B8
Loop Until Label2.Caption = "stopped" Or UCase(l008A) = UCase(Text1.Text)
Timer1.Enabled = False
Timer2.Enabled = False
l008E = Label4 / Label6
sub0263 "²º·‚¸‚• UPSS Room Buster by: EvicTioN"
sub0263 "²º·‚¸‚• Busted into " + Text1.Text + " in " + Label6 + " seconds"
sub0263 "²º·‚¸‚• In only: " + Label4.Caption + " tries!"
sub0263 "²º·‚¸‚• Running at " + l008E + " tries per second"
VBMsg1.SubClasshWnd = 0
Label2.Caption = "Stopped"
Label4.Caption = ""
Label6.Caption = "0"
Label4.Caption = "0"
End If
If  Option2.Value = True Then
Label4.Caption = "0"
Label6.Caption = "0"
l0060% = extfn04BA("AOL Frame25", 0&)
Call sub020F(l0060%, "Edit Go To Menu")
Do
DoEvents
l0092% = extfn04BA("_AOL_MODAL", "Favorite Places")
l0094% = extfn0524(l0092%, "_AOL_EDIT")
gv00CC% = extfn055D(l0094%, mc0096)
l009C% = extfn0524(l0092%, "_AOL_Button")
Loop Until l0092% <> 0
l009E% = extfn02F3(l0094%, mc0050, 0, Text1)
l009E% = extfn02F3(gv00CC%, mc0050, 0, "aol://2719:2-2-" + Text1)
l00A2% = extfn02DF(l009C%, mc00A0, &HD, 0)
l00A2% = extfn02DF(l009C%, mc00A4, &HD, 0)
Timer1.Enabled = True
Timer2.Enabled = True
Label2.Caption = "Busting!"
Do
DoEvents
l0060% = extfn04BA("AOL Frame25", 0&)
Call sub020F(l0060%, Text1)
l00A6 = fn00CD()
If  Not Label2.Caption = "Stopped" Then
Label4 = Label4 + 1
End If
sub01B8
Loop Until Label2.Caption = "Stopped" Or UCase(l00A6) = UCase(Text1.Text)
If  Label2.Caption = "Stopped" Then
Exit Sub
End If
Timer1.Enabled = False
Timer2.Enabled = False
l008E = Label4 / Label6
sub0263 "²º·‚¸‚• UPSS Room Buster by: EvicTioN"
sub0269 (.0001#)
sub0263 "²º·‚¸‚• Busted into " + Text1.Text + " in " + Label6 + " seconds"
sub0269 (.0001#)
sub0263 "²º·‚¸‚• In only: " + Label4.Caption + " tries!"
sub0269 (.0001#)
sub0263 "²º·‚¸‚• Running at " + l008E + " tries per second"
sub0269 (.0001#)
VBMsg1.SubClasshWnd = 0
l007C% = extfn0510(l0060%, "keyword")
l00AE% = extfn0398(l007C%, mc00AA, 0, 0)
Label2.Caption = "Stopped"
Label4.Caption = ""
Label6.Caption = "0"
Label4.Caption = "0"
End If
End Sub

Sub DBPush2_Click ()
VBMsg1.SubClasshWnd = 0
Timer1.Enabled = False
Timer2.Enabled = False
Label2.Caption = "Stopped"
Label4.Caption = ""
End Sub

Sub DBPush3_Click ()
VBMsg1.SubClasshWnd = 0
Timer1.Enabled = False
Timer2.Enabled = False
Label2.Caption = "Stopped"
Label4.Caption = ""
Form12.Hide
Unload Me
End Sub

Sub DBPush4_Click ()
sub01B8
End Sub

Sub Form_Activate ()
sub0193 Me
End Sub

Sub Form_Load ()
sub0193 Me
End Sub

Sub Timer1_Timer ()
sub01AD
End Sub

Sub Timer2_Timer ()
Label6 = Label6 + 1
End Sub

Sub VBMsg1_WindowMessage (hWindow As Integer, Msg As Integer, wParam As Integer, lParam As Long, RetVal As Long, CallDefProc As Integer)
CallDefProc = 0
End Sub
