VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "CONFIG"
   ClientHeight    =   1980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3540
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtWhite 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Text            =   "White"
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtBlack 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Text            =   "Black"
      Top             =   240
      Width           =   1935
   End
   Begin VB.Line Line4 
      BorderWidth     =   5
      X1              =   3525
      X2              =   3525
      Y1              =   0
      Y2              =   1920
   End
   Begin VB.Line Line3 
      BorderWidth     =   5
      X1              =   0
      X2              =   3600
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1920
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   0
      X2              =   3480
      Y1              =   1960
      Y2              =   1960
   End
   Begin VB.Label Label2 
      Caption         =   "White : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Black : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Visible = False
Form1.lblWhite.Caption = Form2.txtWhite
Form1.lblBlack.Caption = Form2.txtBlack
End Sub

