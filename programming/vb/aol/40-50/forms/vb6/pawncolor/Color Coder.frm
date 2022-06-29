VERSION 5.00
Begin VB.Form Color 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Color Coder by PawN"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3330
   Icon            =   "Color Coder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   -10.646
   ScaleMode       =   0  'User
   ScaleWidth      =   3330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "000"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "000"
      Top             =   915
      Width           =   375
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "000"
      Top             =   600
      Width           =   375
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   255
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   255
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   255
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "<FONT COLOR=#000000>"
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "PawN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   2400
      TabIndex        =   7
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   840
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   480
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   210
   End
End
Attribute VB_Name = "Color"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function CLRBars(RedBar As Control, GreenBar As Control, BlueBar As Control)
CLRBars = RGB(RedBar.Value, GreenBar.Value, BlueBar.Value)
End Function
Private Sub Form_Load()
' This shit was poorly coded by whoever made it.
' I got it off KnK and I fixed it myself, now it's tight.  PawN OwNz.
FormOnTop Me
End Sub
Private Sub HScroll1_Change()
Text2.Text = HScroll1.Value
Label7.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
If HScroll1.Value = 0 Then
Label2.Caption = "00"
ElseIf Hex(HScroll1.Value) = "A" Or Hex(HScroll1.Value) = "B" Or Hex(HScroll1.Value) = "C" Or Hex(HScroll1.Value) = "D" Or Hex(HScroll1.Value) = "E" Or Hex(HScroll1.Value) = "F" Then
Label2.Caption = "0" + Hex(HScroll1.Value)
ElseIf (HScroll1.Value) < 10 Then
Label2.Caption = "0" + Hex(HScroll1.Value)
Else
Label2.Caption = Hex(HScroll1.Value)
End If
Text1.Text = "<FONT COLOR=#" + Label2.Caption + Label4.Caption + Label5.Caption + ">"
End Sub
Private Sub HScroll2_Change()
Text3.Text = HScroll2.Value
Label7.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text1.Text = "<FONT COLOR=#" + Label2.Caption + Label4.Caption + Label5.Caption + ">"
If HScroll2.Value = 0 Then
Label4.Caption = "00"
ElseIf Hex(HScroll2.Value) = "A" Or Hex(HScroll2.Value) = "B" Or Hex(HScroll2.Value) = "C" Or Hex(HScroll2.Value) = "D" Or Hex(HScroll2.Value) = "E" Or Hex(HScroll2.Value) = "F" Then
Label4.Caption = "0" + Hex(HScroll2.Value)
ElseIf (HScroll2.Value) < 10 Then
Label4.Caption = "0" + Hex(HScroll2.Value)
Else
Label4.Caption = Hex(HScroll2.Value)
End If
Text1.Text = "<FONT COLOR=#" + Label2.Caption + Label4.Caption + Label5.Caption + ">"
End Sub
Private Sub HScroll3_Change()
Text4.Text = HScroll3.Value
Label7.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text1.Text = "<FONT COLOR=#" + Label2.Caption + Label4.Caption + Label5.Caption + ">"
If HScroll3.Value = 0 Then
Label5.Caption = "00"
ElseIf Hex(HScroll3.Value) = "A" Or Hex(HScroll3.Value) = "B" Or Hex(HScroll3.Value) = "C" Or Hex(HScroll3.Value) = "D" Or Hex(HScroll3.Value) = "E" Or Hex(HScroll3.Value) = "F" Then
Label5.Caption = "0" + Hex(HScroll3.Value)
ElseIf (HScroll3.Value) < 10 Then
Label5.Caption = "0" + Hex(HScroll3.Value)
Else
Label5.Caption = Hex(HScroll3.Value)
End If
Text1.Text = "<FONT COLOR=#" + Label2.Caption + Label4.Caption + Label5.Caption + ">"
End Sub
Private Sub Label8_Click()
Keyword "http://xpawnx.cjb.net"
End Sub
