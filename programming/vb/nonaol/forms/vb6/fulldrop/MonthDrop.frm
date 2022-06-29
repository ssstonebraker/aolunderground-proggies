VERSION 5.00
Begin VB.Form FMonthDrop 
   Caption         =   "Form1"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   4050
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1560
      Width           =   2235
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Hook Combo"
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   840
      Width           =   2235
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   2235
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Standard Combo:"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   1320
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Modified Combo:"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   1185
   End
End
Attribute VB_Name = "FMonthDrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *********************************************************************
'  Copyright ©1999 Karl E. Peterson, All Rights Reserved
'  http://www.mvps.org/vb
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' *********************************************************************
Option Explicit

Private DropHandler As CFullDrop

Private Sub Check1_Click()
   ' hand off message processing to class
   If Check1.Value = vbChecked Then
      DropHandler.hWnd = Combo1.hWnd
   Else
      DropHandler.hWnd = 0
   End If
End Sub

Private Sub Form_Load()
   Dim i As Long
   ' gussy up form a bit
   Set Me.Icon = Nothing
   Me.Caption = "CFullDrop Demo"
   
   ' fill combos with each month name
   For i = 1 To 12
      Combo1.AddItem Format(DateSerial(99, i, 1), "mmmm")
      Combo2.AddItem Combo1.List(i - 1)
   Next i
   Combo1.ListIndex = 0
   Combo2.ListIndex = 0
   
   ' instantiate new drop handler class
   Set DropHandler = New CFullDrop
   ' and turn on subclassing
   Check1.Value = vbChecked
End Sub

