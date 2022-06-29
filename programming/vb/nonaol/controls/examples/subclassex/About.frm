VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Subclass"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "http://www.softcircuits.com"
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label4 
      Caption         =   "Redistributed by Permission"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "<Copyright>"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "<Version>"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   240
      Picture         =   "About.frx":000C
      Top             =   480
      Width           =   420
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   4080
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label1 
      Caption         =   "<Description>"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Subclass - Visual Basic Subclass Control
'Copyright (c) 1997 SoftCircuits Programming (R)
'Redistributed by Permission
'
'This code demonstrates how to write a subclassing control in Visual Basic
'(version 5 or later). The code takes advantage of the new AddressOf
'keyword, which can only be used from a BAS module. A common BAS module
'keeps track of all the current control instances within the current
'process and then intercepts Windows messages, calling specific control
'instances as appropriate.
'
'WARNING: This software is copyrighted. You may only use this software in
'compliance with the following conditions. By using this software, you
'indicate your acceptance of these conditions.
'
' 1.    You may freely use and distribute the supplied Subclass.ocx with your
'       own programs as long as you assume responsibility for such programs
'       and hold the original author(s) harmless from any resulting
'       liabilities.
'
' 2.    You may use this source code within your own programs (embedded within
'       the resulting EXE or DLL file) as long as you assume responsibility
'       for such programs and hold the original author(s) harmless from any
'       resulting liabilities.
'
' 3.    You may NOT distribute this source code, regardless of the extent of
'       modifications, except as part of the original, unmodified
'       Subclass.zip.
'
' 4.    You may NOT recompile this source code and distribute the resulting
'       Subclass.ocx, regardless of the extent of modifications.
'
'The reason for these conditions is to prevent the distribution of different
'versions of Subclass.ocx. Multiple versions would prevent enforcement of
'backwards compatibility and increase problems encountered by programs that
'are distributed with Subclass.ocx. Please respect these conditions. If you
'find a problem or would like an enhancement, please contact SoftCircuits.
'
'This example program was provided by:
' SoftCircuits Programming
' http://www.softcircuits.com
' P.O. Box 16262
' Irvine, CA 92623
Option Explicit

Private Sub Form_Load()
    Label1 = App.FileDescription
    Label2 = "Version " & CStr(App.Major) & _
        "." & Format$(App.Minor, "00")
    Label3 = App.LegalCopyright
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub
