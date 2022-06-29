VERSION 2.00
Begin Form Form1 
   Caption         =   "Scroll Bars Project"
   ClientHeight    =   3555
   ClientLeft      =   1590
   ClientTop       =   2070
   ClientWidth     =   3480
   Height          =   3990
   Icon            =   CHP20.FRX:0000
   Left            =   1515
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   3480
   Top             =   1710
   Width           =   3630
   Begin CommandButton cmndExit 
      Caption         =   "E&xit"
      Height          =   372
      Left            =   2760
      TabIndex        =   7
      Top             =   3120
      Width           =   612
   End
   Begin CommandButton cmndBottomRight 
      Caption         =   "Bottom Right"
      Height          =   372
      Left            =   1440
      TabIndex        =   6
      Top             =   3120
      Width           =   1212
   End
   Begin CommandButton cmndTopRight 
      Caption         =   "Top Right"
      Height          =   372
      Left            =   1440
      TabIndex        =   5
      Top             =   2640
      Width           =   1212
   End
   Begin CommandButton cmndBottomLeft 
      Caption         =   "Bottom Left"
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   1212
   End
   Begin CommandButton cmndTopLeft 
      Caption         =   "Top Left"
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1212
   End
   Begin HScrollBar HScroll1 
      Height          =   252
      LargeChange     =   100
      Left            =   0
      SmallChange     =   10
      TabIndex        =   2
      Top             =   2280
      Width           =   3252
   End
   Begin VScrollBar VScroll1 
      Height          =   2292
      LargeChange     =   100
      Left            =   3240
      SmallChange     =   10
      TabIndex        =   1
      Top             =   0
      Width           =   252
   End
   Begin PictureBox Picture1 
      Height          =   2292
      Left            =   0
      ScaleHeight     =   2265
      ScaleWidth      =   3225
      TabIndex        =   0
      Top             =   0
      Width           =   3252
      Begin PictureBox Picture2 
         AutoSize        =   -1  'True
         Height          =   3495
         Left            =   0
         Picture         =   CHP20.FRX:0302
         ScaleHeight     =   3465
         ScaleWidth      =   4755
         TabIndex        =   8
         Top             =   -60
         Width           =   4785
      End
   End
End

Sub cmndBottomLeft_Click ()
     HScroll1.Value = HScroll1.Min
     VScroll1.Value = VScroll1.Max
End Sub

Sub cmndBottomRight_Click ()
     HScroll1.Value = HScroll1.Max
     VScroll1.Value = VScroll1.Max
End Sub

Sub cmndExit_Click ()
     End
End Sub

Sub cmndTopLeft_Click ()
     HScroll1.Value = HScroll1.Min
     VScroll1.Value = VScroll1.Min
End Sub

Sub cmndTopRight_Click ()
     HScroll1.Value = HScroll1.Max
     VScroll1.Value = VScroll1.Min
End Sub

Sub Form_Load ()
     HScroll1.Max = (Picture2.Width - Picture1.Width)
     VScroll1.Max = (Picture2.Height - Picture1.Height)
     HScroll1.Min = 0
     VScroll1.Min = 0

End Sub

Sub HScroll1_Change ()
     Picture2.Left = -HScroll1.Value
End Sub

Sub VScroll1_Change ()
     Picture2.Top = -VScroll1.Value
End Sub

