VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form WebsiteFrm 
   Caption         =   "WebBrowser"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   5100
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   4680
      Top             =   1560
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Refresh"
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Search"
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Stop"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Home"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Back"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Go"
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Forward"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Text            =   "Http://meltingpot.fortunecity.com/france/700"
      Top             =   960
      Width           =   4095
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   4815
      ExtentX         =   8493
      ExtentY         =   5953
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Address"
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "WebsiteFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Timer1.Enabled = True
    WebBrowser1.Navigate Combo1
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        Combo1_Click
    End If
End Sub

Private Sub Command1_Click()
WebBrowser1.GoForward
End Sub

Private Sub Command2_Click()

WebBrowser1.Navigate Combo1
End Sub

Private Sub Command3_Click()
WebBrowser1.GoBack
End Sub

Private Sub Command4_Click()
WebBrowser1.GoHome
End Sub

Private Sub Command5_Click()
WebBrowser1.Stop
End Sub

Private Sub Command6_Click()
WebBrowser1.GoSearch
End Sub

Private Sub Command7_Click()
WebBrowser1.Refresh
End Sub

Private Sub Form_Load()
WebsiteFrm.Height = 8000
WebsiteFrm.Width = 9000
    WebsiteFrm.Left = (Screen.Width - WebsiteFrm.Width) / 2
     WebsiteFrm.Top = (Screen.Height - WebsiteFrm.Height) / 2
WebBrowser1.Navigate Combo1
End Sub

Private Sub Form_Resize()
If WebsiteFrm.Height < 2000 Then Exit Sub
If WebsiteFrm.Width < 6015 Then WebsiteFrm.Width = 6015
If WebsiteFrm.Height < 5160 Then WebsiteFrm.Height = 5160
Combo1.Width = WebsiteFrm.Width - 1320
WebBrowser1.Height = WebsiteFrm.Height - 1800
WebBrowser1.Width = WebsiteFrm.Width - 350

End Sub

Private Sub Timer1_Timer()
 If WebBrowser1.Busy = False Then
       Timer1.Enabled = False
        Me.Caption = WebBrowser1.LocationName
    Else
        Me.Caption = "Loading Page..."
    End If
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
 Combo1.AddItem WebBrowser1.LocationURL, 0
End Sub

