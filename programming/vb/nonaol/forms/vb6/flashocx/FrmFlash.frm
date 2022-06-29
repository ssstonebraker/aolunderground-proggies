VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "SWFLASH.OCX"
Begin VB.Form FrmFlash 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Flash ocx Help!"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _cx             =   4202585
      _cy             =   4199622
      Movie           =   "http://www.sng.net/computers/flymanvb/peoplechat.swf"
      Src             =   "http://www.sng.net/computers/flymanvb/peoplechat.swf"
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   "http://www.sng.net/computers/flymanvb/peoplechat.swf"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   "000000"
      SWRemote        =   ""
   End
End
Attribute VB_Name = "FrmFlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Hey listen.. click the lil flash screen go to the editor
'then were it has my url to my flash rechange it with
'your url with your flash project!
'This was made by Flyman
'f1yman@gnuspy.com -email
Private Sub ShockwaveFlash1_OnReadyStateChange(newState As Long)

End Sub
