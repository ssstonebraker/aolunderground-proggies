VERSION 2.00
Begin Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Idle Loops"
   ClientHeight    =   2130
   ClientLeft      =   2340
   ClientTop       =   3600
   ClientWidth     =   5355
   Height          =   2535
   Left            =   2280
   LinkTopic       =   "Form1"
   ScaleHeight     =   2130
   ScaleWidth      =   5355
   Top             =   3255
   Width           =   5475
   Begin PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   1665
      Left            =   255
      ScaleHeight     =   1635
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   225
      Width           =   4905
      Begin ComboBox comSecond 
         Height          =   300
         Left            =   270
         TabIndex        =   3
         Top             =   615
         Width           =   4290
      End
      Begin ComboBox comThird 
         Height          =   300
         Left            =   270
         TabIndex        =   2
         Top             =   1065
         Width           =   4290
      End
      Begin ComboBox comFirst 
         Height          =   300
         Left            =   270
         TabIndex        =   1
         Top             =   210
         Width           =   4290
      End
   End
End
Option Explicit

Sub Form_Load ()


    Static nindex As Integer


    ' Show the form on the screen
    frmMain.Show

    
    ' Build the combo boxes. A Do loop is used instead of a For loop to make
    ' the example in the book easier to follow.
    nindex = 0
    Do While nindex < 2001

        nindex = nindex + 1

        comFirst.AddItem "Item " & nindex
        comSecond.AddItem "Item " & nindex
        comThird.AddItem "Item " & nindex

    Loop

    ' This line of code is necessary only when using very large, or slow
    ' Do While Doevents loops, such as those you will build from the example
    ' in the book. It also guarantees that there will be no screen corruption
    ' on a slower machine.
    frmMain.Refresh

End Sub

