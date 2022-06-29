VERSION 5.00
Object = "{5D11CFD2-4E58-11D2-A11D-549F06C10000}#1.0#0"; "MAKE3D.OCX"
Begin VB.Form frm3D 
   Caption         =   "Make It 3D 1.0"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   ScaleHeight     =   1710
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMake3D 
      Caption         =   "Make 3D"
      Height          =   495
      Left            =   1560
      TabIndex        =   7
      Top             =   1080
      Width           =   2055
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   975
      Left            =   3840
      TabIndex        =   6
      Top             =   600
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   600
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   4320
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin Make3D.MakeIt3D MakeIt3D1 
      Left            =   4080
      Top             =   960
      _ExtentX        =   2646
      _ExtentY        =   926
   End
End
Attribute VB_Name = "frm3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMake3D_Click()
    MakeIt3D1.Make3D Command1, 3, 0, True
    MakeIt3D1.Make3D Command2, 3, 0, True
    MakeIt3D1.Make3D Command3, 3, 0, True
    MakeIt3D1.Make3D cmdMake3D, 3, 0, True
    MakeIt3D1.Make3D Text1, 3, 0, True
    MakeIt3D1.Make3D HScroll1, 3, 0, True
    MakeIt3D1.Make3D VScroll1, 3, 0, True
    MakeIt3D1.Make3D List1, 3, 0, True
End Sub
