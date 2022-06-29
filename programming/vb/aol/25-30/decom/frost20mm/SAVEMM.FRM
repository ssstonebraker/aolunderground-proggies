VERSION 2.00
Begin Form save 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Save File"
   ClientHeight    =   1764
   ClientLeft      =   2676
   ClientTop       =   1836
   ClientWidth     =   4044
   ControlBox      =   0   'False
   Height          =   2100
   Left            =   2628
   LinkTopic       =   "Form1"
   ScaleHeight     =   1764
   ScaleWidth      =   4044
   Top             =   1548
   Width           =   4140
   Begin CommonDialog CMDialog1 
      Left            =   2040
      Top             =   960
   End
   Begin TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   615
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
      FontName        =   "MS Sans Serif"
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
      Caption         =   "Save"
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
      TabIndex        =   2
      Top             =   120
      Width           =   1815
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
      TabIndex        =   1
      Top             =   960
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
      Text            =   "Combo1"
      Top             =   120
      Width           =   1815
   End
End
' SAVEMM.FRM
Option Explicit

Sub Combo1_KeyPress (KeyAscii As Integer)
Dim l001C As Variant
Dim l0026 As Variant
Dim l0036 As String
Dim l0038 As String
Dim l003A As Variant
If  l001C = 13 Then
l0026 = Dir1.Path
If  Right$(l0026, 1) <> "\" Then l0026 = l0026 + "\"
l0026 = l0026 + Combo1.Text
Text1 = l0026
l0036$ = Text1.Text
l0038$ = ""
For l003A = 0 To MMMain.List1.ListCount - 1
l0038$ = l0038$ + MMMain.List1.List(l003A) + Chr(13) + Chr(10)
Next l003A
Open l0036$ For Output As #1
Print #1, l0038$
Close #1
Unload save
End If
End Sub

Sub Command3D1_Click ()
Dim l0046 As Variant
Dim l004A As String
Dim l004C As String
Dim l004E As Variant
l0046 = Dir1.Path
If  Right$(l0046, 1) <> "\" Then l0046 = l0046 + "\"
l0046 = l0046 + Combo1.Text
Text1 = l0046
l004A$ = Text1.Text
l004C$ = ""
For l004E = 0 To MMMain.List1.ListCount - 1
l004C$ = l004C$ + MMMain.List1.List(l004E) + Chr(13) + Chr(10)
Next l004E
Open l004A$ For Output As #1
Print #1, l004C$
Close #1
Unload save
End Sub

Sub Command3D2_Click ()
Unload Me
End Sub

Sub Dir1_Change ()
File1.Path = Dir1.Path
End Sub

Sub Form_Load ()
sub05BC Me
sub02A9 Me
Combo1.AddItem "*.frs"
Combo1.ListIndex = 0
File1.Pattern = "*.frs"
End Sub
