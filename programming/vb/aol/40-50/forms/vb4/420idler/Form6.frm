VERSION 4.00
Begin VB.Form Form6 
   BackColor       =   &H00C0C000&
   BorderStyle     =   0  'None
   Caption         =   "Form6"
   ClientHeight    =   1095
   ClientLeft      =   5070
   ClientTop       =   3315
   ClientWidth     =   3405
   Height          =   1500
   Left            =   5010
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   Top             =   2970
   Width           =   3525
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "exit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1080
      TabIndex        =   1
      Text            =   "www.blazeitup.com"
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "send"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Line Line28 
      X1              =   360
      X2              =   3360
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line27 
      X1              =   840
      X2              =   3000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line26 
      X1              =   3360
      X2              =   3360
      Y1              =   120
      Y2              =   1080
   End
   Begin VB.Line Line25 
      X1              =   0
      X2              =   3360
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line24 
      X1              =   0
      X2              =   0
      Y1              =   600
      Y2              =   1080
   End
   Begin VB.Line Line23 
      X1              =   2640
      X2              =   3360
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line22 
      X1              =   2520
      X2              =   3240
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line21 
      X1              =   2640
      X2              =   3360
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line20 
      X1              =   1080
      X2              =   2520
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line19 
      X1              =   840
      X2              =   840
      Y1              =   360
      Y2              =   600
   End
   Begin VB.Line Line18 
      X1              =   960
      X2              =   960
      Y1              =   240
      Y2              =   600
   End
   Begin VB.Line Line17 
      X1              =   2280
      X2              =   2640
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line16 
      X1              =   960
      X2              =   2640
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line15 
      X1              =   3360
      X2              =   3360
      Y1              =   0
      Y2              =   1080
   End
   Begin VB.Line Line14 
      X1              =   3240
      X2              =   3360
      Y1              =   480
      Y2              =   600
   End
   Begin VB.Line Line13 
      X1              =   3240
      X2              =   3360
      Y1              =   120
      Y2              =   0
   End
   Begin VB.Line Line12 
      X1              =   960
      X2              =   960
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line11 
      X1              =   960
      X2              =   2400
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line10 
      X1              =   1080
      X2              =   960
      Y1              =   480
      Y2              =   600
   End
   Begin VB.Line Line9 
      X1              =   1080
      X2              =   960
      Y1              =   120
      Y2              =   0
   End
   Begin VB.Line Line8 
      X1              =   0
      X2              =   840
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line7 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   600
   End
   Begin VB.Line Line6 
      X1              =   840
      X2              =   840
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   840
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   0
      Y1              =   480
      Y2              =   600
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   0
      Y1              =   120
      Y2              =   0
   End
   Begin VB.Line Line2 
      X1              =   720
      X2              =   840
      Y1              =   120
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   720
      X2              =   840
      Y1              =   480
      Y2              =   600
   End
End
Attribute VB_Name = "Form6"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call ChatSend("< a href=" & Text1 & ">click here!</a>")
End Sub

Private Sub Form_Load()
Call FormOnTop(Me)
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


    FormDrag Me
    End Sub



Private Sub Frame1_Click()
Unload Form6
End Sub

