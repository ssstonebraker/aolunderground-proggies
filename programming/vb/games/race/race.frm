VERSION 5.00
Object = "{126CD610-201F-11D1-BBAB-00AA0064C217}#1.0#0"; "RENDERAXXCONTROL.OCX"
Begin VB.Form Mainform 
   Caption         =   "Race-Framerate per second is 0 Website http://www.eldermage.com"
   ClientHeight    =   4800
   ClientLeft      =   -345
   ClientTop       =   2280
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   8025
   Begin RenderAXXControl.RenderAX RenderAX1 
      Height          =   4005
      Left            =   0
      OleObjectBlob   =   "race.frx":0000
      TabIndex        =   0
      Top             =   0
      Width           =   7995
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   6720
      Top             =   4200
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   960
      Top             =   3960
   End
   Begin VB.Label lblLaps 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   4
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Laps Done:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2760
      TabIndex        =   3
      Top             =   4200
      Width           =   2460
   End
   Begin VB.Label Label1 
      Caption         =   "Speed:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label lblSpeed 
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1680
      TabIndex        =   1
      Top             =   4200
      Width           =   300
   End
End
Attribute VB_Name = "Mainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    Mainform.Top = (Screen.Height - Mainform.Height) / 2
    Mainform.Left = (Screen.Width - Mainform.Width) / 2
End Sub

Private Sub Form_Terminate()
End

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub RenderAX1_OnLineCross(ByVal actorno As Long, ByVal line As Long, ByVal sectorno As Long)
    
    'Test to see if the person goes around the tack or not
    If sectorno = 0 Then gbOkay = 1
    If sectorno = 3 And gbOkay = 0 Then
       lblLaps.Caption = lblLaps.Caption + 1
    Else
        If sectorno = 3 Then gbOkay = 0
    End If
End Sub

Private Sub Timer1_Timer()
    StuffToDo
End Sub

Private Sub Timer2_Timer()
    Mainform.Caption = "Race-Framerate per second is " & giFrameRate & "  Website http://www.eldermage.com"
    giFrameRate = 0
End Sub
