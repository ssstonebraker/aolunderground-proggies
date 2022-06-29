VERSION 5.00
Begin VB.Form hopper 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6300
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   6300
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "An attendant has been called."
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   1980
      TabIndex        =   2
      Top             =   1395
      Width           =   2265
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait While The Hopper is Refilled."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   225
      TabIndex        =   1
      Top             =   2085
      Visible         =   0   'False
      Width           =   6120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You Emptied The Hopper!"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   1725
      TabIndex        =   0
      Top             =   225
      Width           =   2640
   End
End
Attribute VB_Name = "hopper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

