VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Download File"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "stop"
      Height          =   285
      Left            =   2760
      TabIndex        =   7
      Top             =   570
      Width           =   1425
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   1440
      TabIndex        =   6
      Text            =   "C:\Windows\Desktop\file.zip"
      Top             =   2325
      Width           =   3705
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1425
      TabIndex        =   5
      Text            =   "http://www.yourname.com/file.zip"
      Top             =   1965
      Width           =   3675
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   195
      Left            =   915
      TabIndex        =   2
      Top             =   285
      Width           =   3270
      _ExtentX        =   5768
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "start"
      Height          =   285
      Left            =   930
      TabIndex        =   1
      Top             =   615
      Width           =   1365
   End
   Begin Project1.download n1 
      Height          =   420
      Left            =   4350
      TabIndex        =   0
      Top             =   3720
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   741
   End
   Begin VB.Label Label4 
      Caption         =   "website:"
      Height          =   270
      Left            =   165
      TabIndex        =   9
      Top             =   1965
      Width           =   1125
   End
   Begin VB.Label Label2 
      Caption         =   "saved path:"
      Height          =   270
      Left            =   210
      TabIndex        =   8
      Top             =   2355
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Speed and Time left."
      Height          =   1350
      Left            =   375
      TabIndex        =   4
      Top             =   2985
      Width           =   3345
   End
   Begin VB.Label Label1 
      Caption         =   "File Information"
      Height          =   765
      Left            =   165
      TabIndex        =   3
      Top             =   1065
      Width           =   4980
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'use this code as freeware
'you dont need to mention my name since i didnt make this code
'originally found in planet-source-code.com by Jeff
'made it into a control, so its a lot easier and it can be used by anyone..
'..without drawing anything on their forms.
'i know the form is crappy, but this is a simple example...youll get it.

Private Sub Command1_Click()
'to let you know, filling in all the options in the download file
'isnt necessary.
'remember i didnt make this code, and i posted this
'so others can find it easier.
n1.DownloadFile Text1.Text, Text2.Text, ProgressBar1, Label1, Label3
End Sub

Private Sub Command2_Click()
n1.CancelDownload
End Sub

Private Sub Form_Unload(Cancel As Integer)
n1.CancelDownload
End Sub
