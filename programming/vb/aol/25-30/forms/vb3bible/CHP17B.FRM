VERSION 2.00
Begin Form formAbout 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "About"
   ClientHeight    =   2100
   ClientLeft      =   624
   ClientTop       =   2556
   ClientWidth     =   4236
   Height          =   2544
   Left            =   564
   LinkTopic       =   "Form2"
   ScaleHeight     =   2100
   ScaleWidth      =   4236
   Top             =   2172
   Width           =   4356
   Begin CommandButton cmndOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   444
      Left            =   1428
      TabIndex        =   1
      Top             =   1284
      Width           =   1476
   End
   Begin Label lablProject 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dialog Project"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   24
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   552
      Index           =   0
      Left            =   516
      TabIndex        =   0
      Top             =   324
      Width           =   3216
   End
End

Sub cmndOK_Click ()
     Hide
End Sub

Sub Form_Load ()
     offset = 14                             'how deep is inset effect
     Load lablProject(1)                     'create two identical
     Load lablProject(2)                     'label copies to create inset
     lablProject(1).ForeColor = QBColor(15)  'White (highlight)
     lablProject(1).Left = lablProject(0).Left - 1.8 * offset
     lablProject(1).Top = lablProject(0).Top - 1.8 * offset
     lablProject(1).Visible = True
     lablProject(2).ForeColor = QBColor(7)   'Light Gray (same as background)
     lablProject(2).Left = lablProject(0).Left - offset
     lablProject(2).Top = lablProject(0).Top - offset
     lablProject(2).Visible = True
     lablProject(2).ZOrder                   'bring to front of stack
End Sub

