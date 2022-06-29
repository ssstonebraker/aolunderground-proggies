VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Greetz"
   ClientHeight    =   2490
   ClientLeft      =   2505
   ClientTop       =   2025
   ClientWidth     =   1560
   Height          =   2895
   Left            =   2445
   LinkTopic       =   "Form1"
   ScaleHeight     =   2490
   ScaleWidth      =   1560
   Top             =   1680
   Width           =   1680
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   240
      Top             =   2520
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "To You"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Greetz"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   2640
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()
timer1.Enabled = True
End Sub


Private Sub timer1_Timer()
Label1.Top = Label1.Top - 50
Label2.Top = Label2.Top - 50
If Label2.Top = -100 Then timer1.Enabled = False
End Sub


