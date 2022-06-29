VERSION 5.00
Begin VB.Form frmStartup 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Pokémon Adventure"
   ClientHeight    =   2310
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   5985
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmStartup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrFade 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   450
      Top             =   1650
   End
   Begin VB.Label lblAfter 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pokémon Adventure"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2205
      Left            =   150
      TabIndex        =   2
      Top             =   -90
      Visible         =   0   'False
      Width           =   5880
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pokémon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   2220
      Left            =   75
      TabIndex        =   0
      Top             =   30
      Width           =   5835
   End
   Begin VB.Label lblTitle2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Adventure"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   75
      TabIndex        =   1
      Top             =   945
      Width           =   5835
   End
   Begin VB.Shape Shape1 
      Height          =   2310
      Left            =   0
      Top             =   0
      Width           =   5985
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    FormOnTop Me
End Sub
Private Sub Form_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Me.Show
    Red = 0
    Green = 0
    Blue = 255
    Red1 = 0
    Green1 = 0
    Blue1 = 0
    num1 = 0
    num2 = 0
    num3 = 0
    num4 = 0
    tmrFade.Enabled = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmSplash.Show
End Sub
Private Sub lblAfter_Click()
    Unload Me
End Sub
Private Sub lblTitle_Click()
    Unload Me
End Sub
Private Sub lblTitle2_Click()
    Unload Me
End Sub
Private Sub tmrFade_Timer()
    If Not num1 = 50 Then
        Blue = Blue - 5
        Color = RGB(Red, Green, Blue)
        Me.BackColor = Color
        Me.Refresh
        num1 = num1 + 1
    ElseIf Not num2 = 70 Then
        lblTitle.FontSize = lblTitle.FontSize + 8
        num2 = num2 + 1
    ElseIf Not num3 = 25 Then
        lblTitle2.FontSize = lblTitle2.FontSize + 8
        num3 = num3 + 1
    ElseIf Not num4 = 50 Then
        lblAfter.Visible = True
        Red1 = Red1 + 5
        Color = RGB(Red1, Green1, Blue1)
        lblAfter.ForeColor = Color
        Me.Refresh
        num4 = num4 + 1
    ElseIf num1 = 50 And num2 = 70 And num3 = 25 And num4 = 50 Then
        Unload Me
     End If
End Sub
