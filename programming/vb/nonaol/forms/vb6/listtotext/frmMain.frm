VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Listbox to textbox"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5160
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Easy ASCII"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   3593
      TabIndex        =   3
      Top             =   120
      Width           =   1455
      Begin VB.CommandButton Command1 
         Caption         =   "Enter"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1680
         Width           =   975
      End
      Begin VB.ListBox List2 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1110
         ItemData        =   "frmMain.frx":0BC2
         Left            =   120
         List            =   "frmMain.frx":0BC4
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listbox to textbox"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   113
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.ListBox List1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1620
         ItemData        =   "frmMain.frx":0BC6
         Left            =   120
         List            =   "frmMain.frx":0BE2
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1695
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  Listbox to textbox is your basic example of how to take
'  a listbox item and move it to a textbox when clicked.
'  I also show how to set the cursor on the textbox exactly
'  where it should go so people would not have to
'  constantly click the textbox to set focus on it.
'  I also show how to fill up a listbox with your basic
'  all the ASCII charcters in three lines of code.
'  With this example you should be able to make a sweet
'  ASCII shop in no time
'  - void

Private Sub Command1_Click()
For i = 32 To 255
    List2.AddItem Chr(i)
Next i
End Sub

Private Sub List1_Click()
'  This works flawlessly.  Even if the person places the cursor in the
'  middle of the text, it still works perfect.
start = Text1.SelStart
lstindex = Len(List1.List(List1.ListIndex))
Text1.SelText = List1.List(List1.ListIndex)
Text1.SetFocus
Text1.SelStart = start + lstindex
End Sub
