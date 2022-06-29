VERSION 5.00
Object = "{8FEDFC6C-2D9F-4548-BE95-683B026F98B3}#14.0#0"; "Q3Style.ocx"
Begin VB.Form frmTest 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test Form"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   348
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "solid"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "glass"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4200
      Top             =   2520
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0000FF00&
      Caption         =   "glass"
      Height          =   1095
      Left            =   4080
      TabIndex        =   2
      Top             =   0
      Width           =   975
      Begin Q3Style.Glass Glass1 
         Left            =   240
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "q3button"
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1815
      Begin Q3Style.Q3Button Q3Button1 
         Height          =   780
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   1376
         ButtonType      =   1
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "numberdisplay"
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   5175
      Begin Q3Style.NumberDisplay NumberDisplay1 
         Height          =   540
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   4860
         _ExtentX        =   8573
         _ExtentY        =   953
         Caption         =   "8888888888"
         Border          =   1
         DigitCount      =   10
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C000C0&
      Caption         =   "status"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   4815
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   160
      X2              =   184
      Y1              =   168
      Y2              =   192
   End
   Begin VB.Line Line7 
      BorderWidth     =   4
      X1              =   184
      X2              =   152
      Y1              =   200
      Y2              =   232
   End
   Begin VB.Line Line6 
      BorderWidth     =   4
      X1              =   40
      X2              =   16
      Y1              =   232
      Y2              =   200
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   48
      X2              =   16
      Y1              =   168
      Y2              =   192
   End
   Begin VB.Line Line3 
      BorderWidth     =   4
      X1              =   48
      X2              =   144
      Y1              =   232
      Y2              =   232
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   56
      X2              =   152
      Y1              =   168
      Y2              =   168
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'that enforces variable delcaration
Private Number As Long
'declare the number here so it can get _
values added to it

Private Sub Command1_Click()
    Glass1.Glassify Me, True
    'call the sub, me refers to the form, true means glass
    'replace Me with frmTest and it will still work
    Label1.Caption = "Lines don't work when the form has a border.."
End Sub

Private Sub Command2_Click()
    Glass1.Glassify Me, False
    'call the sub, me refrs to the form, false means solid
    'replace Me with frmTest and it will still work
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.Caption = "The Glass control is invisible at runtime just like the timer control."
End Sub

Private Sub Q3Button1_Click()
    Select Case Q3Button1.ButtonType
        Case Is = 0
            If MsgBox("Click (Custom) in the properties window to set up button..." & vbCrLf & "Do you want to quit and set up your button?", vbYesNo, "Q3Button") = vbYes Then
                End
            Else
                Q3Button1.ButtonType = Q3Button1.ButtonType + 1
            End If
        Case Else
            Label1.Caption = " You can do different things when the button type is changed.." & Q3Button1.ButtonType
            If Q3Button1.ButtonType = 15 Then
                Q3Button1.ButtonType = 0
            Else
                Q3Button1.ButtonType = Q3Button1.ButtonType + 1
            End If
    End Select
    Debug.Print "Click"
End Sub

Private Sub Q3Button1_DblClick()
    Debug.Print "DBClick"
End Sub

Private Sub Q3Button1_DragDrop(Source As Control, X As Single, Y As Single)
    Debug.Print "Dragdrop: " & Source & " : " & X & " : " & Y
End Sub

Private Sub Q3Button1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Debug.Print "dragover: " & Source & " : " & X & " : " & Y
 End Sub

Private Sub Q3Button1_GotFocus()
    Debug.Print "Got Focus"
End Sub

Private Sub Q3Button1_KeyDown(KeyCode As Integer, Shift As Integer)
    Debug.Print "KeyDown : " & KeyCode & " : " & Shift
End Sub

Private Sub Q3Button1_KeyPress(KeyAscii As Integer)
    Debug.Print "keypress: " & KeyAscii
End Sub

Private Sub Q3Button1_KeyUp(KeyCode As Integer, Shift As Integer)
    Debug.Print "keyup : " & KeyCode & " : " & Shift
End Sub

Private Sub Q3Button1_LostFocus()
    Debug.Print "Lost focus"
End Sub

Private Sub Q3Button1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "MouseDown : " & Button & " : " & Shift & " : " & X & " : " & Y
End Sub

Private Sub Q3Button1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "Mousemove : " & Button & " : " & Shift & " : " & X & " : " & Y
End Sub

Private Sub Q3Button1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "MouseUp : " & Button & " : " & Shift & " : " & X & " : " & Y
End Sub

Private Sub Q3Button1_Validate(Cancel As Boolean)
    Debug.Print "Validate : " & Cancel
End Sub

Private Sub Timer1_Timer()
    Number& = Number& + 100
    'add a number
    NumberDisplay1.Caption = Str$(Number&)
    'be sure to change it to a string...
End Sub
