VERSION 5.00
Begin VB.Form main 
   Caption         =   "Main"
   ClientHeight    =   996
   ClientLeft      =   2868
   ClientTop       =   2952
   ClientWidth     =   2640
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   996
   ScaleWidth      =   2640
   Begin VB.CommandButton Command1 
      Caption         =   "Create SnapShot"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2412
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
snap1
End Sub

Private Sub Command2_Click()
'I got this code from micosoft visual basic
'help topics

  ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files" & _
    "(*.txt)|*.txt|Batch Files (*.bat)|*.bat" & _
    "Bitmap"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Save dialog box
    CommonDialog1.ShowSave
    ' Display name of selected file

    MsgBox CommonDialog1.filename
    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
    Exit Sub

End Sub
