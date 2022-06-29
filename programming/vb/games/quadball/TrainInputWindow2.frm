VERSION 5.00
Begin VB.Form InputWindow2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "You Have Beaten The Best Time!"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox NewName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   90
      TabIndex        =   0
      Top             =   2205
      Width           =   4380
   End
   Begin VB.Label Title 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You Have Betten The High Time For Quad-Ball !"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1635
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4515
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Your Name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   90
      TabIndex        =   3
      Top             =   1890
      Width           =   1500
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   465
      Left            =   660
      TabIndex        =   2
      Top             =   2655
      Width           =   585
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   465
      Left            =   2745
      TabIndex        =   1
      Top             =   2655
      Width           =   1155
   End
End
Attribute VB_Name = "InputWindow2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
'This Form Is Like The "INPUTWINDOW" but Is Used For High Time Scores '
'_____________________________________________________________________'
Private Sub Form_Load()
 InputLoaded = True
 Call Set_Mouse_X_Y(Me.Left / Screen.TwipsPerPixelX + 100, Me.Top / Screen.TwipsPerPixelY + 100)
 Exit2Mouse = False
 Me.Picture = ParentForm.StatBox.Picture
 Me.Show
 Me.Refresh
 LimitMovement
End Sub
Private Sub LimitMovement()
Do
  If Exit2Mouse = True Then GoTo nd:
  Me.ZOrder 0
  DoEvents
  KeepMouseOnForm
 Loop Until Exit2Mouse = True
nd:
Unload Me
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 InputLoaded = False
 Exit2Mouse = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
 InputLoaded = False
End Sub
Private Sub Label3_Click()
 Call SaveTimeTraining(NewName.Text, ParentForm.AllTime.Caption)
 Unload Me
End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Label3.Tag = "yes" Then Exit Sub
 Label3.Tag = "yes"
 If Label4.Tag = " yes" Then
  Label4.ForeColor = RGB(0, 90, 0)
  Label4.Top = Label4.Top + 50
  Label4.FontSize = Label4.FontSize - 5
  Label4.Tag = "no"
 End If
 WAVPlay "click.qbs"
 Label3.Top = Label3.Top - 50
 Label3.ForeColor = RGB(0, 255, 0)
 Label3.FontSize = Label3.FontSize + 5
 End Sub
Private Sub Label4_Click()
 Dim Result As VbMsgBoxResult
 Result = MsgBox("Are You Sure?, If You Click OK Your New Top Score Will not Be Saved!", vbOKCancel, "Confirmation")
 If Result = vbCancel Then
  Exit Sub
 Else
  Exit2Mouse = True
  Unload Me
 End If
End Sub
Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Label4.Tag = "yes" Then Exit Sub
 Label4.Tag = "yes"
 If Label3.Tag = " yes" Then
  Label3.ForeColor = RGB(0, 90, 0)
  Label3.Top = Label3.Top + 50
  Label3.FontSize = Label3.FontSize - 5
  Label3.Tag = "no"
 End If
 WAVPlay "click.qbs"
 Label4.ForeColor = RGB(0, 255, 0)
 Label4.Top = Label4.Top - 50
 Label4.FontSize = Label4.FontSize + 5
 End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Label4.Tag = "yes" Then
  Label4.ForeColor = RGB(0, 90, 0)
  Label4.Top = Label4.Top + 50
  Label4.FontSize = Label4.FontSize - 5
  Label4.Tag = "no"
 End If
 If Label3.Tag = "yes" Then
  Label3.ForeColor = RGB(0, 90, 0)
  Label3.Top = Label3.Top + 50
  Label3.FontSize = Label3.FontSize - 5
  Label3.Tag = "no"
 End If
 Label4.ForeColor = RGB(0, 90, 0)
 Label3.ForeColor = RGB(0, 90, 0)
End Sub
Private Sub KeepMouseOnForm()
If Exit2Mouse = True Then GoTo nd:
Me.ZOrder 0
If Get_Mouse_X >= Int((Me.Left + Me.Width) _
 / Screen.TwipsPerPixelX) - 5 Then _
 Set_Mouse_X ((Me.Left + Me.Width) / Screen.TwipsPerPixelX) - 5
If Get_Mouse_X <= Int((Me.Left) _
 / Screen.TwipsPerPixelX) Then _
 Set_Mouse_X ((Me.Left) / Screen.TwipsPerPixelX) + 5
If Get_Mouse_Y <= Int(Me.Top / Screen.TwipsPerPixelY) + 5 Then _
 Set_Mouse_Y (Me.Top / Screen.TwipsPerPixelY) + 5
If Get_Mouse_Y >= Int((Me.Top + Me.Height) / Screen.TwipsPerPixelY) - 5 Then _
 Set_Mouse_Y ((Me.Top + Me.Height) / Screen.TwipsPerPixelY) - 5
nd:
End Sub



