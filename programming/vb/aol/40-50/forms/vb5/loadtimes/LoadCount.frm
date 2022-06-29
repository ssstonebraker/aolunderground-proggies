VERSION 4.00
Begin VB.Form LoadCount 
   Caption         =   "Form1"
   ClientHeight    =   750
   ClientLeft      =   2235
   ClientTop       =   1515
   ClientWidth     =   2850
   Height          =   1155
   Left            =   2175
   LinkTopic       =   "Form1"
   Picture         =   "LoadCount.frx":0000
   ScaleHeight     =   750
   ScaleWidth      =   2850
   Top             =   1170
   Width           =   2970
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "LoadCount"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()
Times = GetSetting("KnK", "Count", "stunt")
If Times = "" Then
SaveSetting "KnK", "Count", "stunt", 1
Else
SaveSetting "KnK", "Count", "stunt", Times + 1
Label1 = Times
End If
End Sub


