VERSION 5.00
Begin VB.Form DragExmple 
   BorderStyle     =   0  'None
   Caption         =   "Morphine's Drag Example"
   ClientHeight    =   3315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2385
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   2385
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   $"DragExmple.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   30
      TabIndex        =   3
      Top             =   1755
      Width           =   2340
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   2385
      Y1              =   1635
      Y2              =   1635
   End
   Begin VB.Label yLBL 
      Caption         =   "this label is needed for the cursor position of Y"
      Height          =   885
      Left            =   1215
      TabIndex        =   2
      Top             =   615
      Width           =   990
   End
   Begin VB.Label xLBL 
      Caption         =   "this label is needed for the cursor position of X"
      Height          =   855
      Left            =   195
      TabIndex        =   1
      Top             =   600
      Width           =   870
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "hit run if your not sure what this does"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   480
      Left            =   105
      TabIndex        =   0
      Top             =   75
      Width           =   2175
   End
End
Attribute VB_Name = "DragExmple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Label1.Caption = "drag me!"
DragExmple.Height = 990
End Sub


Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aa = Button
If aa = 1 Then  'if left mouse button is down...
dd = (X - xLBL.Caption)  'dd = cursor's X position, keeping the
                            'same spot on the label
ee = (Y - yLBL.Caption)  'ee = cursor's Y position, keeping the
                            'same spot on the label
Left = Left + (dd) 'form.left = cursor's X position
Top = Top + (ee)  'form.top = cursor's Y position
Exit Sub
End If  'end if
xLBL.Caption = X  'gives xLBL cursor's X position
yLBL.Caption = Y  'gives yLBL cursor's Y position
'---------------------
'remember, this code goes in MouseMove, NOT MouseDown.
'the xLBL and yLBL dont have to be visible, after all,
'they're not too actractive.
'QUESTIONS?: morfeen@n2.com
'AIM: o morphine o
'                       -morphine
End Sub


