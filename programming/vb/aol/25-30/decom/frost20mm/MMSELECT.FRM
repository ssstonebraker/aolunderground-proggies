VERSION 2.00
Begin Form MMselect 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Select Mails"
   ClientHeight    =   3132
   ClientLeft      =   1368
   ClientTop       =   1488
   ClientWidth     =   6384
   ControlBox      =   0   'False
   Height          =   3468
   Left            =   1320
   LinkTopic       =   "Form1"
   ScaleHeight     =   3132
   ScaleWidth      =   6384
   Top             =   1200
   Width           =   6480
   Begin SSCommand Command3D4 
      Caption         =   "Finished"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin SSCommand Command3D3 
      Caption         =   "Select With..."
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin SSCommand Command3D2 
      Caption         =   "De-Select All"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin SSCommand Command3D1 
      Caption         =   "Select All"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   2760
      Width           =   1095
   End
   Begin TextBox Text1 
      BackColor       =   &H00000000&
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   9.6
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   3480
      TabIndex        =   2
      Text            =   "•UPS•"
      Top             =   2760
      Width           =   1695
   End
   Begin SSFrame Frame3D1 
      Alignment       =   2  'Center
      Caption         =   "Select Mails"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin ListBox List1 
         BackColor       =   &H00000000&
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   2340
         Left            =   120
         MultiSelect     =   1  'Simple
         TabIndex        =   1
         Top             =   240
         Width           =   6135
      End
   End
End
' MMSELECT.FRM
Option Explicit

Sub Command3D1_Click ()
Dim l001A As Variant
For l001A = 0 To List1.ListCount - 1
    List1.Selected(l001A) = True
Next l001A
End Sub

Sub Command3D2_Click ()
Dim l0024 As Variant
For l0024 = 0 To List1.ListCount - 1
    List1.Selected(l0024) = False
Next l0024
End Sub

Sub Command3D4_Click ()
Dim l002C As Variant
MMMain.List2.Clear
For l002C = 0 To MMselect.List1.ListCount - 1
    If  MMselect.List1.Selected(l002C) = True Then
        MMMain.List2.AddItem l002C
    End If
Next l002C
If  MMMain.List2.ListCount = 0 Then MsgBox "Please choose at least 1 mail", 16, "Error": Exit Sub
Me.Hide
End Sub

Sub Form_Load ()
sub05BC Me
sub02A9 Me
End Sub
