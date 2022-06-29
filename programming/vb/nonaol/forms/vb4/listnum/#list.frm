VERSION 4.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "# the list"
   ClientHeight    =   1950
   ClientLeft      =   3120
   ClientTop       =   1860
   ClientWidth     =   2880
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Height          =   2355
   Icon            =   "#list.frx":0000
   Left            =   3060
   LinkTopic       =   "Form1"
   ScaleHeight     =   1950
   ScaleWidth      =   2880
   Top             =   1515
   Width           =   3000
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   2400
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1200
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   3
      Left            =   2400
      Top             =   840
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "+#"
      Height          =   375
      Left            =   120
      MouseIcon       =   "#list.frx":0ECA
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1440
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
list2.AddItem ""
list1.List(list1.ListIndex) = Form1.Text2 + 1


End Sub

Private Sub Form_Load()
list1.AddItem "0"
End Sub

Private Sub Timer1_Timer()
Text2.Text = list2.ListCount
End Sub

