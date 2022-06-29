VERSION 5.00
Begin VB.Form Form24 
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form24"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   23
      Left            =   4320
      TabIndex        =   24
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   22
      Left            =   4320
      TabIndex        =   23
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   21
      Left            =   4320
      TabIndex        =   22
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   20
      Left            =   4320
      TabIndex        =   21
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   19
      Left            =   4320
      TabIndex        =   20
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   18
      Left            =   4320
      TabIndex        =   19
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   17
      Left            =   2880
      TabIndex        =   18
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   16
      Left            =   2880
      TabIndex        =   17
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   15
      Left            =   2880
      TabIndex        =   16
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   14
      Left            =   2880
      TabIndex        =   15
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   13
      Left            =   2880
      TabIndex        =   14
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   12
      Left            =   2880
      TabIndex        =   13
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   11
      Left            =   1440
      TabIndex        =   12
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   10
      Left            =   1440
      TabIndex        =   11
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   9
      Left            =   1440
      TabIndex        =   10
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   8
      Left            =   1440
      TabIndex        =   9
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   7
      Left            =   1440
      TabIndex        =   8
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   6
      Left            =   1440
      TabIndex        =   7
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   5
      Left            =   0
      TabIndex        =   6
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   4
      Left            =   0
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2160
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "Form24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    SendChat msg
End Function

Private Sub Command18_Click()

End Sub
