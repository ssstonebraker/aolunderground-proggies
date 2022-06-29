VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "SWFLASH.OCX"
Begin VB.Form help_frm 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caustik Converter Help"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5625
   Icon            =   "help_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash helpmovie 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      _cx             =   4204279
      _cy             =   4204067
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
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
Attribute VB_Name = "help_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
StayOnTop Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
helpmovie.Stop
helpmovie.Movie = "NONE"
End Sub

Private Sub helpmovie_FSCommand(ByVal command As String, ByVal args As String)
Me.Visible = False
helpmovie.Stop
helpmovie.Movie = "NONE"
End Sub

