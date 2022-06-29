VERSION 2.00
Begin Form loadd 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Load File"
   ClientHeight    =   1800
   ClientLeft      =   2796
   ClientTop       =   2880
   ClientWidth     =   4044
   Height          =   2136
   Left            =   2748
   LinkTopic       =   "Form2"
   ScaleHeight     =   1800
   ScaleWidth      =   4044
   Top             =   2592
   Width           =   4140
   Begin CommonDialog CMDialog1 
      Left            =   2040
      Top             =   480
   End
   Begin TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin FileListBox File1 
      BackColor       =   &H00000000&
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   1080
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1815
   End
   Begin SSCommand Command3D2 
      Caption         =   "Cancel"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin SSCommand Command3D1 
      Caption         =   "Load"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin DriveListBox Drive1 
      BackColor       =   &H00000000&
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2160
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin DirListBox Dir1 
      BackColor       =   &H00000000&
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   750
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin ComboBox Combo1 
      BackColor       =   &H00000000&
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Text            =   "*.frs"
      Top             =   120
      Width           =   1815
   End
End
' LOADMM.FRM
Option Explicit

Sub Command3D1_Click ()
Dim l0020 As Variant
Dim l0030 As String
Dim l0032 As Integer
Dim l0038 As String
l0020 = Dir1.Path
If  Right$(l0020, 1) <> "\" Then l0020 = l0020 + "\"
l0020 = l0020 + File1.List(File1.ListIndex)
Text1 = l0020
    l0030$ = Text1.Text
    l0032% = FreeFile
    Open l0030$ For Input As l0032%
    MMMain.List1.Clear
    Do: DoEvents
    Line Input #1, l0038$
    If  l0038$ <> "" Then MMMain.List1.AddItem l0038$
    Loop Until EOF(l0032%)
    Close l0032%
    MMMain.Frame3D1.Caption = "People " + MMMain.List1.ListCount
    Unload loadd
End Sub

Sub Command3D2_Click ()
Unload Me
End Sub

Sub Dir1_Change ()
File1.Path = Dir1.Path
End Sub

Sub File1_DblClick ()
Dim l003E As Variant
Dim l0042 As String
Dim l0044 As Integer
Dim l0046 As String
l003E = Dir1.Path
If  Right$(l003E, 1) <> "\" Then l003E = l003E + "\"
l003E = l003E + File1.List(File1.ListIndex)
Text1 = l003E
    l0042$ = Text1.Text
    l0044% = FreeFile
    Open l0042$ For Input As l0044%
    MMMain.List1.Clear
    Do: DoEvents
    Line Input #1, l0046$
    If  l0046$ <> "" Then MMMain.List1.AddItem l0046$
    Loop Until EOF(l0044%)
    Close l0044%
    MMMain.Frame3D1.Caption = "People " + MMMain.List1.ListCount
    Unload loadd
End Sub

Sub Form_Load ()
sub05BC Me
sub02A9 Me
Combo1.AddItem "*.frs"
Combo1.ListIndex = 0
File1.Pattern = "*.frs"
End Sub
