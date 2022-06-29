VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "SWFLASH.OCX"
Begin VB.Form frmFlash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5220
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flash 
      Height          =   5055
      Left            =   -105
      TabIndex        =   0
      Top             =   90
      Width           =   7800
      _cx             =   13758
      _cy             =   8916
      Movie           =   "C:\windows\desktop\bulba.swf"
      Src             =   "C:\windows\desktop\bulba.swf"
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   0   'False
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   -1  'True
      BGColor         =   ""
      SWRemote        =   ""
   End
   Begin VB.Label lblPokemon1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   6945
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "frmFlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    FormOnTop Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmSplash.Show
End Sub
Private Sub lblPokemon1_Click()
    Unload Me
End Sub
