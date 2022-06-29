VERSION 4.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Example of a srollable Picture Box"
   ClientHeight    =   4935
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   5145
   Height          =   5340
   Left            =   1080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   Top             =   1170
   Width           =   5265
   Begin VB.PictureBox Picture1 
      Height          =   4695
      Left            =   120
      ScaleHeight     =   4635
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   5
         Left            =   0
         Max             =   100
         TabIndex        =   3
         Top             =   4400
         Width           =   4600
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   0
         Picture         =   "PicScr~1.frx":0000
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   2
         Top             =   0
         Width           =   480
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   4400
         LargeChange     =   5
         Left            =   4600
         Max             =   100
         TabIndex        =   1
         Top             =   0
         Width           =   270
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private Sub HScroll1_Change()
Static LastPos As Integer
If HScroll1.Value < LastPos Then
    Picture2.Left = Picture2.Left - (((Picture1.Width - (VScroll1.Width + 100)) - Picture2.Width) / 100) * (LastPos - HScroll1.Value)
Else
    Picture2.Left = Picture2.Left + (((Picture1.Width - (VScroll1.Width + 100)) - Picture2.Width) / 100) * (HScroll1.Value - LastPos)
End If
LastPos = HScroll1.Value
End Sub


Private Sub VScroll1_Change()
Static LastPos As Integer
If VScroll1.Value < LastPos Then
    Picture2.Top = Picture2.Top - (((Picture1.Height - (HScroll1.Height + 100)) - Picture2.Height) / 100) * (LastPos - VScroll1.Value)
Else
    Picture2.Top = Picture2.Top + (((Picture1.Height - (HScroll1.Height + 100)) - Picture2.Height) / 100) * (VScroll1.Value - LastPos)
End If
LastPos = VScroll1.Value

End Sub

