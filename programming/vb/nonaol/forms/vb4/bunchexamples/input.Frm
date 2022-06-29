VERSION 4.00
Begin VB.Form Form1 
   ClientHeight    =   810
   ClientLeft      =   1950
   ClientTop       =   3900
   ClientWidth     =   4200
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "Arial"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Height          =   1215
   Left            =   1890
   LinkTopic       =   "Form1"
   ScaleHeight     =   810
   ScaleWidth      =   4200
   Top             =   3555
   Width           =   4320
   Begin VB.CommandButton Command1 
      Caption         =   "click this now"
      Height          =   300
      Left            =   2640
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
'iz representing eXcel 2001

z = InputBox("enter your first name", "name thing", edfault) ' this is how you make an input box :)
Label1.Caption = "hey " & z & " how you doing?" ' displays what you typed in the input box + a little hello
End Sub

