VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "SWFLASH.OCX"
Begin VB.Form about_frm 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About caustik converter"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5385
   Icon            =   "about_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   5385
   StartUpPosition =   2  'CenterScreen
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash aboutmovie 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      _cx             =   4203855
      _cy             =   4201315
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "Best"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
   End
End
Attribute VB_Name = "about_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End Sub

Private Sub aboutmovie_FSCommand(ByVal command As String, ByVal args As String)
Me.Visible = False
aboutmovie.Movie = "NONE"
End Sub

Private Sub Form_Load()
StayOnTop Me
End Sub

Private Sub List1_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
aboutmovie.Movie = "NONE"
End Sub
