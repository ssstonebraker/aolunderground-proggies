VERSION 2.00
Begin Form formInfo 
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Info"
   ClientHeight    =   5580
   ClientLeft      =   1320
   ClientTop       =   1356
   ClientWidth     =   5004
   Height          =   6024
   Left            =   1260
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   5004
   Top             =   972
   Width           =   5124
   Begin VScrollBar scrl 
      Height          =   4932
      Left            =   4620
      TabIndex        =   3
      Top             =   540
      Width           =   252
   End
   Begin PictureBox pictContainer 
      Height          =   4932
      Left            =   120
      ScaleHeight     =   4908
      ScaleWidth      =   4488
      TabIndex        =   1
      Top             =   540
      Width           =   4512
      Begin PictureBox pictInfo 
         AutoRedraw      =   -1  'True
         Height          =   1632
         Left            =   300
         ScaleHeight     =   1608
         ScaleWidth      =   1728
         TabIndex        =   2
         Top             =   1320
         Width           =   1752
      End
   End
   Begin CommandButton cmndClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   312
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   912
   End
End

Sub cmndClose_Click ()
    
    Me.Hide

End Sub

Sub Form_Load ()

    pictInfo.Left = 0
    pictInfo.Top = 0
    pictInfo.Width = pictContainer.ScaleWidth
    pictInfo.Height = 3 * pictContainer.ScaleHeight
    scrl.Min = 0
    scrl.Max = pictInfo.Height
    scrl.LargeChange = scrl.Max \ 20
    scrl.SmallChange = scrl.Max \ 100

End Sub

Sub scrl_Change ()

    pictInfo.Top = -scrl.Value

End Sub

