VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Port For Chat:"
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2775
   Icon            =   "joinports.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   2775
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Text            =   "4000"
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton frmports 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Port for This chat:"
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1260
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
'OK was clicked.

'Redefine the settings.
glPort = txtPort.text
Call WriteToINI("lastport", "port", txtPort.text, "c:\windows\happychat.ini")
Form1.Timer7.Enabled = True
Unload Me
End Sub

Private Sub Label2_Click()

End Sub

Private Sub frmports_Click()
Unload Me
End Sub
