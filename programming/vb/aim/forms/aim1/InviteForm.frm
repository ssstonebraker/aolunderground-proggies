VERSION 4.00
Begin VB.Form InviteForm 
   Caption         =   "InviteTest Form"
   ClientHeight    =   915
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   2295
   Height          =   1320
   Left            =   1080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   2295
   Top             =   1170
   Width           =   2415
   Begin VB.CommandButton Command3 
      Caption         =   "Find Ivo"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Open Invo"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close Invo"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "InviteForm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call CloseInvite
End Sub


Private Sub Command2_Click()
Call OpenInvite
End Sub


Private Sub Command3_Click()
Call FindInvite
End Sub


Private Sub Form_Unload(Cancel As Integer)
Main.Show
End Sub


