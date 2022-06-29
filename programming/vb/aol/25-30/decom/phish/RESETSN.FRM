VERSION 2.00
Begin Form ResetSN 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "SN Replacer"
   ClientHeight    =   1335
   ClientLeft      =   2895
   ClientTop       =   420
   ClientWidth     =   3750
   FontBold        =   -1  'True
   FontItalic      =   0   'False
   FontName        =   "MS Serif"
   FontSize        =   8.25
   FontStrikethru  =   0   'False
   FontUnderline   =   0   'False
   Height          =   1740
   Icon            =   RESETSN.FRX:0000
   Left            =   2835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   3750
   Top             =   75
   Width           =   3870
   Begin SSPanel Panel3D1 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   0
      RoundedCorners  =   0   'False
      ShadowColor     =   1  'Black
      TabIndex        =   7
      Top             =   0
      Width           =   3735
      Begin SSCommand Command3D32 
         BevelWidth      =   1
         Caption         =   "Replace SN"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2520
         RoundedCorners  =   0   'False
         TabIndex        =   3
         Top             =   840
         Width           =   1095
      End
      Begin TextBox Text1 
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   315
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   0
         Top             =   120
         Width           =   1575
      End
      Begin TextBox Text2 
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Text            =   "C:\AOL25"
         Top             =   840
         Width           =   975
      End
      Begin TextBox Text3 
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   315
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
      Begin Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "SN to replace (EXACT!):"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   1935
      End
      Begin Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Text to replace with:"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
      Begin Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "AOL Directory:"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
   End
End
' RESETSN.FRM
Option Explicit
Const c0068 = 16 ' &H10%
Const c0084 =-1 ' &HFFFF%
Const c0086 = 3 ' &H3%

Sub Command3D1_Click ()
Unload Me
End Sub

Sub Command3D32_Click ()
Dim l001E As String
Dim l002E As String
Dim l0036 As String
Dim l003E As String
Dim l0040 As Long
Dim l0042 As Variant
Dim l0046 As String
Dim l0048 As Variant
Dim l004C As String
Dim l004E As Variant
Dim l0052 As Variant
Dim l0056 As Variant
Dim l005A As Variant
Dim l005E As Long
Dim l0060 As Long
Dim l0064 As Variant
' Const c0068 = 16 ' &H10%
Dim l006C As Variant
Dim l0074 As Variant
Dim l007A As Variant
Dim l0080 As Variant
' Const c0084 =-1 ' &HFFFF%
' Const c0086 = 3 ' &H3%
On Error Resume Next
l001E$ = App.Path + "\phish.INI"
Open l001E$ For Output As #1
Print #1, Text2.Text
Close #1
ResetSN.Hide
WaitForReset.Show
DoEvents
l002E$ = Text2.Text + "\idb\main.idx"
l0036$ = Text1.Text
Text3.Text = Text3.Text + Space$(Len(Text1) - Len(Text3))
l003E$ = Text3.Text
Open l002E$ For Binary Access Read Write As #1
l0040& = LOF(1)
l0042 = 16384
l0046$ = ""
l0048 = Len(l0036$) - 1
l004C$ = String$(l0048, 1)
l004E = Int(l0040& / l0042)
l0052 = l0040& - l004E * l0042
For l0056 = 1 To l004E
  l005A = 0
  l0046$ = l004C$ + Input$(l0042, 1)
L6FE0:
  l005A = InStr(l005A + 1, l0046$, l0036$)
  If  l005A = 0 Then GoTo L7076
  l005E& = (l0056 - 1) * l0042 + l005A - l0048
  l0060& = Seek(1)
  Seek #1, l005E&
  Put #1, , l003E$
  Seek #1, l0060&
  GoTo L6FE0
L7076:
  l004C$ = Right$(l0046$, l0048)
Next l0056
If  l0052 = 0 Then GoTo L7146
l005A = 0
l0046$ = l004C$ + Input$(l0042, 1)
L70CE:
  l005A = InStr(l005A + 1, l0046$, l0036$)
  If  l005A = 0 Then GoTo L7146
  l005E& = l004E * l0042 + l005A - l0048
  Seek #1, l005E&
  Put #1, , l003E$
  GoTo L70CE
L7146:
  Close #1
l0064 = fn3F58("Welcome")
If  l0064 = 0 Then l0064 = fn3F58("Goodbye")
l006C = extfn6E8(l0064, c0068, 0&, 0&)
DoEvents
sub3FC8 ("%g~")
DoEvents
l0064 = fn4038("Welcome")
l0074 = extfnC8(l0064, "_AOL_Combobox")
l007A = extfn1478(l0074)
SendKeys "{UP}{UP}{UP}{UP}{UP}{UP}{UP}"
l0080 = extfn170(App.Path + "\created.wav", 1)
Unload WaitForReset
Unload Me
End Sub

Sub Form_Load ()
Dim l008A As Variant
Dim l0092 As Variant
Dim l0096 As String
Dim l0098 As String
On Error Resume Next
l008A = extfn4120(ResetSN.hWnd, c0084, 0, 0, 0, 0, c0086)
ResetSN.Left = (Screen.Width - ResetSN.Width) / 2
ResetSN.Top = 0
l0092 = extfn1478(Text1.hWnd)
l0096$ = App.Path + "\phish.INI"
Open l0096$ For Input As #1
Line Input #1, l0098$
Close #1
If  l0098$ <> "" Then Text2 = l0098$
End Sub

Sub Form_Unload (p1 As Integer)
Phisher.Show
End Sub
