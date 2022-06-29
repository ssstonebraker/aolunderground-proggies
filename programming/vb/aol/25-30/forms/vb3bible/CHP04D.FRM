VERSION 2.00
Begin Form formModeless 
   Caption         =   "Modeless Form"
   ClientHeight    =   2940
   ClientLeft      =   975
   ClientTop       =   1890
   ClientWidth     =   6240
   Height          =   3375
   Left            =   900
   LinkTopic       =   "Form3"
   ScaleHeight     =   2940
   ScaleWidth      =   6240
   Top             =   1530
   Width           =   6390
   Begin TextBox Text1 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   1452
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "This form is modeless. Try clicking on another window in the Forms application. The focus leaves this window and goes to the other one. Try maximizing, minimizing and restoring the window to normal too."
      Top             =   360
      Width           =   3372
   End
   Begin CommandButton cmndClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   372
      Left            =   4440
      TabIndex        =   3
      Top             =   2160
      Width           =   1572
   End
   Begin CommandButton cmndMinimize 
      Caption         =   "&Minimize"
      Height          =   372
      Left            =   4440
      TabIndex        =   2
      Top             =   1320
      Width           =   1572
   End
   Begin CommandButton cmndNormal 
      Caption         =   "&Normal"
      Height          =   372
      Left            =   4440
      TabIndex        =   1
      Top             =   840
      Width           =   1572
   End
   Begin CommandButton cmndMaximize 
      Caption         =   "Ma&ximize"
      Height          =   372
      Left            =   4440
      TabIndex        =   0
      Top             =   360
      Width           =   1572
   End
End

Sub cmndClose_Click ()
     Hide
End Sub

Sub cmndMaximize_Click ()
     WindowState = 2
End Sub

Sub cmndMinimize_Click ()
     WindowState = 1
End Sub

Sub cmndNormal_Click ()
     WindowState = 0
End Sub

