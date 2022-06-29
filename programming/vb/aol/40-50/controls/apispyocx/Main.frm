VERSION 5.00
Object = "{018F500F-755E-11D1-A237-D87C90F57546}#9.0#0"; "ApI SpY.ocx"
Begin VB.Form Form1 
   Caption         =   "ApI SpY OcX Ex -x- Izekial83 -x- Funkdemon@yahoo.com"
   ClientHeight    =   2310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6375
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2310
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4920
      Top             =   360
   End
   Begin ApISpYxIzekial83.ApI ApI1 
      Height          =   750
      Left            =   4920
      TabIndex        =   9
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1323
   End
   Begin VB.Label Label9 
      Caption         =   "Parent Module:"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   2040
      Width           =   11895
   End
   Begin VB.Label Label8 
      Caption         =   "Parent Class:"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   1800
      Width           =   11895
   End
   Begin VB.Label Label7 
      Caption         =   "Parent Text:"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1560
      Width           =   11775
   End
   Begin VB.Label Label6 
      Caption         =   "Parent Handle:"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1320
      Width           =   11895
   End
   Begin VB.Label Label5 
      Caption         =   "Window ID Number:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   11895
   End
   Begin VB.Label Label4 
      Caption         =   "Window Style:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   11895
   End
   Begin VB.Label Label3 
      Caption         =   "Window Text:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   11895
   End
   Begin VB.Label Label2 
      Caption         =   "Window Class:"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   11895
   End
   Begin VB.Label Label1 
      Caption         =   "Window Handle:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer1_Timer()
Call ApI1.WindowSPYLabels(Label1, Label2, Label3, Label4, Label5, Label6, Label7, Label8, Label9)
End Sub

