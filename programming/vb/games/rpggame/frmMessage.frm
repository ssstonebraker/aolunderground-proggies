VERSION 5.00
Begin VB.Form frmMessage 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6660
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "RPG Message"
      BeginProperty Font 
         Name            =   "Home"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "An error has occured"
      BeginProperty Font 
         Name            =   "Home"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1440
      TabIndex        =   1
      Top             =   960
      Width           =   3285
   End
   Begin VB.Label lblReturn 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "Home"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   2400
      TabIndex        =   0
      Top             =   1560
      Width           =   1635
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim enter As Boolean

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'make the return label white
    lblReturn.ForeColor = QBColor(15)
End Sub

Private Sub lblReturn_Click()
    enter = True
End Sub

Public Function showMessage(msg As String, frmName As Form)
    
    
    'show the message in the label
    lblMessage.Caption = msg
    'center the message on the form
    lblMessage.Left = (frmMessage.Width - lblMessage.Width) / 2
    
    frmName.Enabled = False
        
    'show the message frm
    frmMessage.Visible = True
    
    'play the message sound
    Call sndPlaySound(sndMessage, &H1)
    frmMessage.SetFocus
    
    While enter = False
    DoEvents
    Wend
    
    enter = False
    frmName.Enabled = True
    frmName.SetFocus
    
End Function

Private Sub lblReturn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblReturn.ForeColor = QBColor(12)
End Sub
