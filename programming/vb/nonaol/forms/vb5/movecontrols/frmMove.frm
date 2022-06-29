VERSION 5.00
Begin VB.Form frmMove 
   Caption         =   "Moving Controls During Runtime | plastik@dosfx.com"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   200
      Width           =   5260
      Begin VB.ListBox lstList 
         Height          =   2870
         IntegralHeight  =   0   'False
         ItemData        =   "frmMove.frx":0000
         Left            =   3885
         List            =   "frmMove.frx":000D
         TabIndex        =   4
         Top             =   120
         Width           =   1350
      End
      Begin VB.TextBox txt2 
         Height          =   285
         Left            =   25
         TabIndex        =   3
         Top             =   2700
         Width           =   3855
      End
      Begin VB.TextBox txt1 
         Height          =   2580
         Left            =   25
         TabIndex        =   2
         Top             =   120
         Width           =   3855
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Moving Controls During Runtime | plastik@dosfx.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5260
   End
End
Attribute VB_Name = "frmMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'________________________________________________________'
'                                                        '
'   Moving Controls During Runtime | plastik@dosfx.com   '
'                                                        '
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
'________________________________________________________'
'///////////////////////////|\\\\\\\\\\\\\\\\\\\\\\\\\\\\'
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
'  Here you will notice that the controls resize in the  '
'  forms resize procedure, thats the code that will be   '
'  processed whenever the form is resized.  If you try   '
'  to resize the form to a very small state it will      '
'  error because controls will start trying to go in the '
'  negatives, so to prevent this write a code that will  '
'  keep the form from going too small.                   '
'________________________________________________________'
'\\\\\\\\\\\\\\\\\\\\\\\\\\\|////////////////////////////'
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'

'________________________________________________________'
'///////////////////////////|\\\\\\\\\\\\\\\\\\\\\\\\\\\\'
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
'                       Move Method                      '
'                                                        '
' Syntax:  Move (left, top, width, height)               '
'                                                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Part   'Description                                     '
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
'left   ' Representing the left property of the control  '
'       ' when you resize.                               '
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
'top    ' Representing the top property of the control   '
'       ' when you resize.                               '
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
'width  ' Representing the width of the control when you '
'       ' resize.                                        '
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
'height ' Representing the height of the control when you'
'       ' resize.                                        '
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
'________________________________________________________'
'\\\\\\\\\\\\\\\\\\\\\\\\\\\|////////////////////////////'
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'

Private Sub Form_Resize()
Call lblTitle.Move(0, 0, ScaleWidth, 255)
Call Frame1.Move(0, 200, ScaleWidth, ScaleHeight - 200)
Call txt1.Move(25, 120, ScaleWidth - 55 - lstList.Width, ScaleHeight - 350 - txt2.Height)
Call lstList.Move(txt1.Left + txt1.Width, 120, 1350, ScaleHeight - 340)
Call txt2.Move(25, txt1.Top + txt1.Height, ScaleWidth - lstList.Width - 55, 285)
End Sub
