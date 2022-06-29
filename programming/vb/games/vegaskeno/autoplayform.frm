VERSION 5.00
Begin VB.Form autoplayform 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AutoPlay Settings  (lazy day keno)"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4065
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cancelbutton 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton gobutton 
      Caption         =   "GO"
      Default         =   -1  'True
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox bettext 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2280
      TabIndex        =   8
      ToolTipText     =   "Enter Amount of Bet."
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox repeattext 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      TabIndex        =   0
      ToolTipText     =   "Enter Number of Deals"
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Amount to bet."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label with1bet 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1740
      TabIndex        =   7
      Top             =   960
      Width           =   585
   End
   Begin VB.Label with2bet 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Label5"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1740
      TabIndex        =   6
      Top             =   1320
      Width           =   585
   End
   Begin VB.Label with3bet 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Label4"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1740
      TabIndex        =   5
      Top             =   1680
      Width           =   585
   End
   Begin VB.Label with4bet 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Label4"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1740
      TabIndex        =   4
      Top             =   2040
      Width           =   585
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Autoplay will not accept settings that exceed your current amount of credits."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the number of times to repeat."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   2760
      Width           =   2055
   End
End
Attribute VB_Name = "autoplayform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancelbutton_Click()
board.Enabled = True
options.Enabled = True
Unload Me
End Sub

Private Sub form_load()
bettext.Text = "0"
repeattext.Text = "0"
options.Enabled = False
board.Enabled = False
Label2.Caption = "You have " & dollars & " credits."
with1bet.Caption = "You can play " & (dollars / 1) & " times with one bet."
If dollars >= 2 Then
with2bet.Caption = "You can play " & Format(((dollars / 2) - 1), "###,###") & " times with two bet."
Else
with2bet.Caption = "You can play " & Format((dollars / 2), "###,###") & " times with two bet."
End If
If dollars >= 3 Then
with3bet.Caption = "You can play " & Format(((dollars / 3) - 1), "###,###") & " times with three bet."
Else
with3bet.Caption = "You can play " & Format((dollars / 3), "###,###") & " times with three bet."
End If
If dollars >= 4 Then
with4bet.Caption = "You can play " & Format(((dollars / 4) - 1), "###,###") & " times with four bet."
repeattext.Text = Format(((dollars / 4) - 1), "###,###")
bettext.Text = "4"
Else
with4bet.Caption = "You can play " & Format((dollars / 4), "###,###") & " times with four bet."
End If

End Sub

Private Sub gobutton_Click()
On Error GoTo ErrHandler
Dim startdollars As Double
Dim enddollars As Double
Dim x As Long
Dim y As Long
Dim f As Long
Dim result As Double
x = 0
y = 0
f = 0
startdollars = 0
enddollars = 0
result = 0
x = repeattext.Text
y = bettext.Text
f = (x * y)
If f > dollars Or y > 4 Then
MsgBox "This is not a Valid Setting, Please Re-enter", vbOKOnly
repeattext.Text = ""
bettext.Text = ""
autoplayform.Refresh
Exit Sub
Else
Me.Hide
options.Hide
board.Enabled = True
options.Enabled = True
startdollars = dollars
autokenomode = 1
board.change_button.Enabled = False
Call board.autokeno(x, y)
enddollars = dollars
result = enddollars - startdollars
If result = 0 Then
MsgBox "You broke even!", vbApplicationModal, "Even Steven"
ElseIf result > 0 Then
MsgBox "Your Have Won " & Format(result, "###,###") & " Credits", vbOKOnly
ElseIf result < 0 Then
result = Abs(result)
MsgBox "You Have Lost " & Format(result, "###,###") & " Credits", vbOKOnly
End If
autokenomode = 0
board.change_button.Enabled = True
Unload Me
Exit Sub
End If
ErrHandler:
    autokenomode = 0
    MsgBox "This is not a Valid Setting, Please Re-enter", vbOKOnly
    repeattext.Text = ""
    bettext.Text = ""
    autoplayform.Refresh
Exit Sub
End Sub

