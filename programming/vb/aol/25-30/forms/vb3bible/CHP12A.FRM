VERSION 2.00
Begin Form formMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Font Project"
   ClientHeight    =   4200
   ClientLeft      =   870
   ClientTop       =   1485
   ClientWidth     =   4605
   Height          =   4635
   Icon            =   CHP12A.FRX:0000
   Left            =   795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   4605
   Top             =   1125
   Width           =   4755
   Begin CommandButton cmndExit 
      Caption         =   "E&xit"
      Height          =   612
      Left            =   2640
      TabIndex        =   5
      Top             =   2280
      Width           =   1812
   End
   Begin CommandButton cmndDisplay 
      Caption         =   "&Display Fonts"
      Height          =   1332
      Left            =   2640
      TabIndex        =   4
      Top             =   720
      Width           =   1812
   End
   Begin TextBox textSample 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   1092
      Left            =   120
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3000
      Width           =   4332
   End
   Begin Frame framEffects 
      Caption         =   "Special Effects"
      Height          =   2292
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2412
      Begin CheckBox chekUnderline 
         Caption         =   "&Underline"
         Height          =   192
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   1452
      End
      Begin CheckBox chekStrikethru 
         Caption         =   "&StrikeThru"
         Height          =   252
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   1572
      End
      Begin CheckBox chekItalic 
         Caption         =   "&Italic"
         Height          =   252
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1812
      End
      Begin CheckBox chekBold 
         Caption         =   "&Bold"
         Height          =   252
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1812
      End
   End
   Begin ComboBox cmboSize 
      Height          =   288
      Left            =   3000
      TabIndex        =   1
      Text            =   "Combo2"
      Top             =   120
      Width           =   1452
   End
   Begin ComboBox cmboFont 
      Height          =   288
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   2772
   End
End

Sub chekBold_Click ()
   textSample.FontBold = Not textSample.FontBold
End Sub

Sub chekItalic_Click ()
    textSample.FontItalic = Not textSample.FontItalic
End Sub

Sub chekStrikethru_Click ()
    textSample.FontStrikethru = Not textSample.FontStrikethru
End Sub

Sub chekUnderline_Click ()
    textSample.FontUnderline = Not textSample.FontUnderline
End Sub

Sub cmboFont_Click ()
     textSample.FontName = cmboFont.Text
     textSample.Text = cmboFont.Text
End Sub

Sub cmboSize_Click ()
     textSample.FontSize = Val(cmboSize.Text)
     cmboSize.Text = Str$(textSample.FontSize)
End Sub

Sub cmndDisplay_Click ()
    formMain.Hide            'hide this form
    Load formFontList        'load the font list form, don't display
    formFontList.Show        'and display after the form has had a chance to create list
End Sub

Sub cmndExit_Click ()
     End
End Sub

Sub Form_Load ()
    For F = 0 To (Screen.FontCount - 1)           'iterate through all screen fonts
        cmboFont.AddItem Screen.Fonts(F)          'add the font name to the combo box
    Next F
    For P = 4 To 36                               'iterate through some common point sizes
        cmboSize.AddItem Str$(P)                  'and add point size to the combo box
    Next P
    cmboFont.Text = textSample.FontName           'set the combo box to the default font
    cmboSize.Text = Str$(textSample.FontSize)     'make combo box read the correct font size
End Sub

