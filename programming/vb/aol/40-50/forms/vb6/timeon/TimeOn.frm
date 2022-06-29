VERSION 5.00
Begin VB.Form TimeOn 
   Caption         =   "Timeonline by Chron__x"
   ClientHeight    =   1290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   ScaleHeight     =   1290
   ScaleWidth      =   3870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Check TimeOnline"
      Height          =   285
      Left            =   60
      TabIndex        =   0
      Top             =   15
      Width           =   3810
   End
   Begin VB.Label Label1 
      Height          =   675
      Left            =   30
      TabIndex        =   1
      Top             =   405
      Width           =   3765
   End
End
Attribute VB_Name = "TimeOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Time0n$ = Timeonline
    Label1.Caption = Time0n$
    MsgBox (Time0n$)
End Sub

Private Sub Form_Load()
MsgBox "This is a simple example that only touches the basics of 32 bit win api. If you dont understand it then email me at" & Chr(13) & Chr(10) & "Chron__x@hotmail.com" & Chr(13) & Chr(10) & "Visit my site" & Chr(13) & Chr(10) & "http://chronx.cjb.net" & Chr(13) & Chr(10) & "If you use this form (which i dont like) or any of these subs please just add me to your greets", vbInformation, "Msg From Chron__x"

End Sub
