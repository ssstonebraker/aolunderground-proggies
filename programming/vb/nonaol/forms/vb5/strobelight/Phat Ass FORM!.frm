VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Strobe Light v1 By ßiohazard"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Left            =   2280
      Top             =   -120
   End
   Begin VB.Timer Timer1 
      Left            =   2040
      Top             =   -120
   End
   Begin VB.CommandButton command2 
      BackColor       =   &H000000FF&
      Caption         =   "Stop"
      Height          =   195
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   11
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label8 
      Caption         =   "10"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "100"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DotSpeed ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ColorSpeed ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H00000000&
      Height          =   2295
      Left            =   240
      Top             =   840
      Width           =   4455
   End
   Begin VB.Shape Shape9 
      Height          =   2055
      Left            =   360
      Top             =   960
      Width           =   4215
   End
   Begin VB.Shape Shape8 
      Height          =   135
      Left            =   1320
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Shape Shape7 
      Height          =   375
      Left            =   1200
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Shape Shape6 
      Height          =   615
      Left            =   1080
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Shape Shape5 
      Height          =   855
      Left            =   960
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Shape Shape4 
      Height          =   1095
      Left            =   840
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Shape Shape3 
      Height          =   1335
      Left            =   720
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Shape Shape2 
      Height          =   1575
      Left            =   600
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Shape Shape1 
      Height          =   1815
      Left            =   480
      Top             =   1080
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    StartStrobeLight
End Sub

Private Sub Command2_Click()
    StopStrobeLight
End Sub

Private Sub StartStrobeLight()

    Timer1.Enabled = True
    Timer1.Interval = Label7.Caption
    Timer2.Enabled = True
    Timer2.Interval = Label8.Caption
    Label1.BackColor = &HFF&
    Shape1.Visible = True
    Shape2.Visible = True
    Shape3.Visible = True
    Shape4.Visible = True
    Shape5.Visible = True
    Shape6.Visible = True
    Shape7.Visible = True
    Shape8.Visible = True
    Shape9.Visible = True
    Shape10.Visible = True
    
End Sub

Private Sub StopStrobeLight()
    Timer1.Enabled = False
    Timer2.Enabled = False
    Form1.BackColor = &HC0C0C0
Label1.BackColor = &HC0C0C0
Label2.BackColor = &HC0C0C0
Label3.BackColor = &HC0C0C0
Label4.BackColor = &HC0C0C0
Label5.BackColor = &HC0C0C0
Label6.BackColor = &HC0C0C0
Label7.BackColor = &HC0C0C0
Label8.BackColor = &HC0C0C0
Label9.BackColor = &HC0C0C0
Label10.BackColor = &HC0C0C0
    
End Sub

Private Sub Label10_Click()
Label4.Caption = Val(Label4) - 1
Label8.Caption = Val(Label8) + 1
End Sub

Private Sub Label2_Click()
Label5.Caption = Val(Label5) + 1
Label7.Caption = Val(Label7) - 1
End Sub

Private Sub Label3_Click()
Label5.Caption = Val(Label5) - 1
End Sub

Private Sub Label9_Click()
Label4.Caption = Val(Label4) + 1
Label8.Caption = Val(Label8) - 1
End Sub

Private Sub Timer1_Timer()
    If Shape1.BorderStyle = vbBSDot Then
        Shape1.BorderStyle = vbBSDashDot
    Else
        Shape1.BorderStyle = vbBSDot
    End If
    
    If Shape2.BorderStyle = vbBSDot Then
        Shape2.BorderStyle = vbBSDashDot
    Else
        Shape2.BorderStyle = vbBSDot
    End If
    
    
        If Shape3.BorderStyle = vbBSDot Then
        Shape3.BorderStyle = vbBSDashDot
    Else
        Shape3.BorderStyle = vbBSDot
    End If
    
        If Shape4.BorderStyle = vbBSDot Then
        Shape4.BorderStyle = vbBSDashDot
    Else
        Shape4.BorderStyle = vbBSDot
    End If
    
        If Shape5.BorderStyle = vbBSDot Then
        Shape5.BorderStyle = vbBSDashDot
    Else
        Shape5.BorderStyle = vbBSDot
    End If
    
        If Shape6.BorderStyle = vbBSDot Then
        Shape6.BorderStyle = vbBSDashDot
    Else
        Shape6.BorderStyle = vbBSDot
    End If
    
        If Shape7.BorderStyle = vbBSDot Then
        Shape7.BorderStyle = vbBSDashDot
    Else
        Shape7.BorderStyle = vbBSDot
    End If
    
        If Shape8.BorderStyle = vbBSDot Then
        Shape8.BorderStyle = vbBSDashDot
    Else
        Shape8.BorderStyle = vbBSDot
    End If
    
        If Shape9.BorderStyle = vbBSDot Then
        Shape9.BorderStyle = vbBSDashDot
    Else
        Shape9.BorderStyle = vbBSDot
    End If
    
        If Shape10.BorderStyle = vbBSDot Then
        Shape10.BorderStyle = vbBSDashDot
    Else
        Shape10.BorderStyle = vbBSDot
    End If
    
      
End Sub

Private Sub Timer2_Timer()
Dim a As Variant
a = Int(Rnd * 10)
If a = 1 Then
Form1.BackColor = &HFF&
Command1.BackColor = &HFF&
command2.BackColor = &HFF&
Label1.BackColor = &HFF&
Label2.BackColor = &HFF&
Label3.BackColor = &HFF&
Label4.BackColor = &HFF&
Label5.BackColor = &HFF&
Label6.BackColor = &HFF&
Label7.BackColor = &HFF&
Label8.BackColor = &HFF&
Label9.BackColor = &HFF&
Label10.BackColor = &HFF&
End If
If a = 2 Then
Form1.BackColor = &HFF0000
Command1.BackColor = &HFF0000
command2.BackColor = &HFF0000
Label1.BackColor = &HFF0000
Label2.BackColor = &HFF0000
Label3.BackColor = &HFF0000
Label4.BackColor = &HFF0000
Label5.BackColor = &HFF0000
Label6.BackColor = &HFF0000
Label7.BackColor = &HFF0000
Label8.BackColor = &HFF0000
Label9.BackColor = &HFF0000
Label10.BackColor = &HFF0000
End If
If a = 3 Then
Form1.BackColor = &HFF00&
Command1.BackColor = &HFF00&
command2.BackColor = &HFF00&
Label1.BackColor = &HFF00&
Label2.BackColor = &HFF00&
Label3.BackColor = &HFF00&
Label4.BackColor = &HFF00&
Label5.BackColor = &HFF00&
Label6.BackColor = &HFF00&
Label7.BackColor = &HFF00&
Label8.BackColor = &HFF00&
Label9.BackColor = &HFF00&
Label10.BackColor = &HFF00&
End If
If a = 4 Then
Form1.BackColor = &HFFFF&
Command1.BackColor = &HFFFF&
command2.BackColor = &HFFFF&
Label1.BackColor = &HFFFF&
Label2.BackColor = &HFFFF&
Label3.BackColor = &HFFFF&
Label4.BackColor = &HFFFF&
Label5.BackColor = &HFFFF&
Label6.BackColor = &HFFFF&
Label7.BackColor = &HFFFF&
Label8.BackColor = &HFFFF&
Label9.BackColor = &HFFFF&
Label10.BackColor = &HFFFF&
End If
If a = 5 Then
Form1.BackColor = &H808080
Command1.BackColor = &H808080
command2.BackColor = &H808080
Label1.BackColor = &H808080
Label2.BackColor = &H808080
Label3.BackColor = &H808080
Label4.BackColor = &H808080
Label5.BackColor = &H808080
Label6.BackColor = &H808080
Label7.BackColor = &H808080
Label8.BackColor = &H808080
Label9.BackColor = &H808080
Label10.BackColor = &H808080
End If
End Sub
