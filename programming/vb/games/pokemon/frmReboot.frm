VERSION 5.00
Begin VB.Form frmReboot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Restarting..."
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3105
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   3105
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrReboot 
      Interval        =   1500
      Left            =   2520
      Top             =   1110
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   420
      Left            =   735
      TabIndex        =   0
      Top             =   1095
      Width           =   1650
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Server Restarting"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   105
      TabIndex        =   1
      Top             =   45
      Width           =   2910
   End
End
Attribute VB_Name = "frmReboot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    End
End Sub
Private Sub Form_Activate()
    FormOnTop Me
End Sub
Private Sub tmrReboot_Timer()
    frmMain.Show
    frmTrayIcon.SetCallback frmMain
    frmMain.Hide
    Unload Me
End Sub
