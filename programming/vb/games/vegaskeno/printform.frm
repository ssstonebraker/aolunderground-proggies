VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form print2form 
   Caption         =   "Bragging Board"
   ClientHeight    =   1665
   ClientLeft      =   6510
   ClientTop       =   4695
   ClientWidth     =   2535
   Icon            =   "printform.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1665
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cancel 
      Cancel          =   -1  'True
      Caption         =   "&CANCEL"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton printoptions 
      Caption         =   "Print Keno Stats"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton printboard 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "print2form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancel_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If board.Visible = False Then
printboard.Enabled = False
printboard.Caption = "Display Keno Board To Print"
Else
printboard.Enabled = True
printboard.Caption = "Print Keno Board"
End If
End Sub

Private Function printboardsub()
If board.Visible = True Then
    Dim BeginPage, EndPage, NumCopies, i
    ' Set Cancel to True
  CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Display the Print dialog box
   CommonDialog1.ShowPrinter
    ' Get user-selected values from the dialog box
    BeginPage = CommonDialog1.FromPage
    EndPage = CommonDialog1.ToPage
    NumCopies = CommonDialog1.Copies
    For i = 1 To NumCopies
    board.PrintForm
    Next i
    Exit Function
ErrHandler:
    ' User pressed the Cancel button
    Exit Function
    Else
        MsgBox " display keno board to print", vbOKOnly
        End If
End Function
Private Sub form_load()
Call ShadeForm(Me)
If board.Visible = False Then
printboard.Enabled = False
printboard.Caption = "Display Keno Board To Print"
End If

End Sub

Private Sub printboard_Click()
Call printboardsub
End Sub

Private Sub printoptions_Click()
statsform.Show
End Sub
Sub ShadeForm(Frm As Form)


    Dim iLoop As Integer
    Dim NumberOfRects As Integer
    Dim GradColor As Long
    Dim GradValue As Integer
    Frm.ScaleMode = 3
    Frm.DrawStyle = 6
    Frm.DrawWidth = 2
    Frm.AutoRedraw = True
    NumberOfRects = 64
    

    For iLoop = 1 To 64
        GradValue = 255 - (iLoop * 4 - 1)
        
        GradColor = RGB(GradValue, GradValue, GradValue)
        Frm.Line (0, Frm.ScaleHeight * (iLoop - 1) / 64)-(Frm.ScaleWidth, Frm.ScaleHeight * iLoop / 64), GradColor, BF
        
    Next iLoop


    Frm.Refresh
End Sub
