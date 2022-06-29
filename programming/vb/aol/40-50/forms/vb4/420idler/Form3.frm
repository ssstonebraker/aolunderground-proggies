VERSION 4.00
Begin VB.Form Form3 
   BackColor       =   &H00C0C000&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   1125
   ClientLeft      =   4890
   ClientTop       =   3285
   ClientWidth     =   3735
   Height          =   1530
   Left            =   4830
   LinkTopic       =   "Form3"
   ScaleHeight     =   1125
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   Top             =   2940
   Width           =   3855
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
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   3135
   End
   Begin VB.Line Line21 
      X1              =   3720
      X2              =   3360
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line20 
      X1              =   3600
      X2              =   2640
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line19 
      X1              =   120
      X2              =   360
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line18 
      X1              =   0
      X2              =   360
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line17 
      X1              =   3720
      X2              =   3840
      Y1              =   1800
      Y2              =   2280
   End
   Begin VB.Line Line16 
      X1              =   0
      X2              =   0
      Y1              =   1920
      Y2              =   1800
   End
   Begin VB.Line Line15 
      X1              =   120
      X2              =   480
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line14 
      X1              =   0
      X2              =   240
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line13 
      X1              =   360
      X2              =   3240
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line12 
      X1              =   3720
      X2              =   3720
      Y1              =   1080
      Y2              =   0
   End
   Begin VB.Line Line11 
      X1              =   3720
      X2              =   3720
      Y1              =   0
      Y2              =   1080
   End
   Begin VB.Line Line10 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1080
   End
   Begin VB.Line Line9 
      X1              =   0
      X2              =   3720
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line8 
      X1              =   3600
      X2              =   3720
      Y1              =   960
      Y2              =   1080
   End
   Begin VB.Line Line7 
      X1              =   3600
      X2              =   3720
      Y1              =   120
      Y2              =   0
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   0
      Y1              =   120
      Y2              =   0
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   0
      Y1              =   960
      Y2              =   1080
   End
   Begin VB.Line Line4 
      X1              =   3600
      X2              =   3600
      Y1              =   120
      Y2              =   960
   End
   Begin VB.Line Line3 
      X1              =   360
      X2              =   3240
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   120
      Y1              =   960
      Y2              =   120
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3600
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   $"Form3.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "Form3"
Attribute VB_Creatable = False
Attribute VB_Exposed = False



Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


    FormDrag Me
    End Sub


Private Sub Form_Load()
Call FormOnTop(Me)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


    FormDrag Me
    End Sub


Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


    FormDrag Me
    End Sub




Private Sub Frame1_Click()
Unload Form3
End Sub

