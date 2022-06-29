VERSION 4.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFFF&
   Caption         =   "                      Circle Form!"
   ClientHeight    =   2625
   ClientLeft      =   2355
   ClientTop       =   2565
   ClientWidth     =   4485
   Height          =   3030
   Icon            =   "Form1.frx":0000
   Left            =   2295
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   4485
   Top             =   2220
   Width           =   4605
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0442
      ForeColor       =   &H00008000&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gamerhere@aol.com"
      Height          =   195
      Left            =   1440
      TabIndex        =   2
      Top             =   2160
      Width           =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Circle Example"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   400
         size            =   24
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   3195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This is an example of a circular form using visual basic!!!"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3945
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, _
 ByVal bRedraw As Boolean) As Long


Private Sub Form_Load()

SetWindowRgn hWnd, _
  CreateEllipticRgn(0, 0, 300, 200), True
Label3.Caption = "By" & vbCrLf & "Gamerhere@aol.com"
End Sub

