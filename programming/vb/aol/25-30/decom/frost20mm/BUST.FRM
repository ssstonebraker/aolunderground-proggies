VERSION 2.00
Begin Form bust 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "FrostBite v2.o Room Buster"
   ClientHeight    =   1272
   ClientLeft      =   2712
   ClientTop       =   2796
   ClientWidth     =   3516
   ControlBox      =   0   'False
   Height          =   1608
   Left            =   2664
   LinkTopic       =   "Form1"
   ScaleHeight     =   1272
   ScaleWidth      =   3516
   Top             =   2508
   Width           =   3612
   Begin SSFrame Frame3D2 
      Alignment       =   2  'Center
      Caption         =   "Attempts"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   2160
      TabIndex        =   7
      Top             =   0
      Width           =   1215
      Begin Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "0"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   18
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
   End
   Begin SSFrame Frame3D1 
      Alignment       =   2  'Center
      Caption         =   "Room"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   1935
      Begin ComboBox Combo1 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   330
         Left            =   120
         TabIndex        =   6
         Top             =   290
         Width           =   1695
      End
   End
   Begin OptionButton Option2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Goto Menu"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Keyword"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   720
      Value           =   -1  'True
      Width           =   1215
   End
   Begin CommandButton Command2 
      Caption         =   "Exit"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin CommandButton Command1 
      Caption         =   "Start"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
End
' BUST.FRM
Option Explicit
Dim m001A As Variant
Const mc0054 = 12 ' &HC%
Const mc005C = 258 ' &H102%
Const mc0062 = 16 ' &H10%
Const mc006C = 2 ' &H2%

Sub Combo1_KeyPress (KeyAscii As Integer)
If  KeyAscii = 13 Then Command1_Click
End Sub

Sub Command1_Click ()
Dim l0040 As String
Dim l004A As Integer
Dim l004E As Integer
Dim l0052 As Integer
' Const mc0054 = 12 ' &HC%
Dim l0058 As Variant
' Const mc005C = 258 ' &H102%
Dim l0060 As Integer
' Const mc0062 = 16 ' &H10%
Dim l006A As Integer
' Const mc006C = 2 ' &H2%
Dim l0070 As Integer
Dim l0072 As Integer
If  Command1.Caption = "Start" Then
Command1.Caption = "Stop"
m001A = 0
If  Combo1.Text = "" Then MsgBox "Please enter a room to enter", 16, "Error": Exit Sub
Label1.Caption = "1"
l0040$ = fn0246()
If  Option1.Value = True Then
LAEF6:
If  m001A = 1 Then Exit Sub
Call sub045F("Keyword")
l004A% = extfn0B02("AOL Frame25", 0&)
Do
DoEvents
If  m001A = 1 Then Exit Sub
l004E% = extfn0B10(l004A%, "Keyword")
Loop Until l004E% > 0
l0052% = extfn0B24(l004E%, "_AOL_Edit")
l0058 = extfn0AC9(l0052%, mc0054, 0, "aol://2719:2-2-" & Combo1.Text)
l0058 = extfn0AB5(l0052%, mc005C, 13, 0)
Do
DoEvents
If  m001A = 1 Then Exit Sub
l0060% = extfn0B02("#32770", "America Online")
Loop Until l0060% > 0 Or fn0246() <> l0040$
If  l0060% > 0 Then
If  m001A = 1 Then Exit Sub
l0060% = extfn0B02("#32770", "America Online")
l0058 = extfn0AB5(l0060%, mc0062, 0, 0)
Label1.Caption = Label1.Caption + 1
GoTo LAEF6
ElseIf fn0246() <> l0040$ Then
If  m001A = 1 Then Exit Sub
Command1.Caption = "Start"
End If
ElseIf Option2.Value = True Then
sub045F "Edit Go To Menu"
    Do
    DoEvents
    If  m001A = 1 Then Exit Sub
    l006A% = extfn0B02("_AOL_Modal", "Favorite Places")
    Loop Until l006A% > 0
    l0052% = extfn0B24(l006A%, "_AOL_Edit")
    l0058 = extfn0AC9(l0052%, mc0054, 0, "")
    l0058 = extfn0AC9(l0052%, mc0054, 0, Combo1.Text)
    l0070% = extfn0B50(l0052%, mc006C)
    l0058 = extfn0AC9(l0070%, mc0054, 0, "")
    If  m001A = 1 Then Exit Sub
    l0058 = extfn0AC9(l0070%, mc0054, 0, "aol://2719:2-2-" & Combo1.Text)
    l0072% = extfn0B10(l006A%, "Save Changes")
    sub02B3 l0072%
    Do
    DoEvents
    If  m001A = 1 Then Exit Sub
    l006A% = extfn0B02("_AOL_Modal", 0&)
    Loop Until l006A% = 0
LB4E0:
    sub045F (Combo1.Text)
    Do
    DoEvents
    l0060% = extfn0B02("#32770", "America Online")
    If  m001A = 1 Then Exit Sub
    Loop Until l0060% > 0 Or fn0246() <> l0040$
    If  l0060% > 0 Then
    If  m001A = 1 Then Exit Sub
    l0060% = extfn0B02("#32770", "America Online")
    l0058 = extfn0AB5(l0060%, mc0062, 0, 0)
    Label1.Caption = Label1.Caption + 1
    GoTo LB4E0
    ElseIf fn0246() > l0040$ Then
    If  m001A = 1 Then Exit Sub
    Command1.Caption = "Start"
    End If
End If
ElseIf Command1.Caption = "Stop" Then
Command1.Caption = "Start"
m001A = 1
End If
End Sub

Sub Command2_Click ()
MMMain.Show
Unload bust
End Sub

Sub Form_Load ()
sub05BC Me
sub02A9 Me
Combo1.AddItem "Fate"
Combo1.AddItem "Fate2"
Combo1.AddItem "Fate3"
Combo1.AddItem "Fate4"
Combo1.AddItem "Fate5"
Combo1.AddItem "Fate6"
Combo1.AddItem "Fate7"
Combo1.AddItem "Fate8"
Combo1.AddItem "Fate9"
Combo1.AddItem "Server"
Combo1.AddItem "Server2"
Combo1.AddItem "Server3"
Combo1.AddItem "Server4"
Combo1.AddItem "Server5"
Combo1.AddItem "Server6"
Combo1.AddItem "Server7"
Combo1.AddItem "Server8"
Combo1.AddItem "Server9"
Combo1.AddItem "FrostBite"
End Sub
