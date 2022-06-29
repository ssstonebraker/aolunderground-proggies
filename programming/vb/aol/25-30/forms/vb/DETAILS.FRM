VERSION 2.00
Begin Form frmMain 
   BackColor       =   &H00C00000&
   Caption         =   "Entry validation example"
   ClientHeight    =   2355
   ClientLeft      =   2925
   ClientTop       =   2985
   ClientWidth     =   4755
   Height          =   2760
   Left            =   2865
   LinkTopic       =   "Form1"
   ScaleHeight     =   2355
   ScaleWidth      =   4755
   Top             =   2640
   Width           =   4875
   Begin CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   1860
      TabIndex        =   4
      Top             =   1485
      Width           =   1215
   End
   Begin TextBox txtAge 
      BackColor       =   &H00FFFF00&
      Height          =   285
      Left            =   2115
      MaxLength       =   3
      TabIndex        =   3
      Top             =   840
      Width           =   2070
   End
   Begin TextBox txtForename 
      BackColor       =   &H00FFFF00&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2145
      TabIndex        =   1
      Top             =   300
      Width           =   2070
   End
   Begin Label Label2 
      BackColor       =   &H00C00000&
      Caption         =   "Enter your age"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   315
      TabIndex        =   2
      Top             =   870
      Width           =   1320
   End
   Begin Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "Enter your forename"
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   300
      TabIndex        =   0
      Top             =   315
      Width           =   1815
   End
End
Option Explicit

Sub cmdOK_Click ()

    '**** Place the new code here
    
    
    Unload Me


End Sub

