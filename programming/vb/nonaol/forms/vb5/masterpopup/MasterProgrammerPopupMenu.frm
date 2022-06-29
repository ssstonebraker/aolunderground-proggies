VERSION 5.00
Begin VB.Form MasterProgrammerPopupMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master Programer Popup Menu"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MasterProgrammerPopupMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Exit?"
      Top             =   2160
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "            Test3"
      ToolTipText     =   "Test3 - Left Click"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Text            =   "            Test3"
      ToolTipText     =   "Test3 - Right Click"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Test1"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      ToolTipText     =   "Test1 - Left Click"
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test1"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Test1 - Right Click"
      Top             =   720
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   1920
      X2              =   1920
      Y1              =   120
      Y2              =   1920
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LEFT CLICK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      ToolTipText     =   "Left Click"
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RIGHT CLICK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "Right Click"
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Test2"
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      ToolTipText     =   "Test2 - Left Click"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Test2"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Test2 - Right Click"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Menu MasterProgramerPopupMenu 
      Caption         =   "Master Programer Popup Menu"
      Visible         =   0   'False
      Begin VB.Menu Master 
         Caption         =   "Master"
      End
      Begin VB.Menu Programer 
         Caption         =   "Programer"
      End
      Begin VB.Menu Popup 
         Caption         =   "Popup"
      End
      Begin VB.Menu Menu 
         Caption         =   "Menu"
      End
   End
   Begin VB.Menu ¤Exit 
      Caption         =   "Exit?"
      Visible         =   0   'False
      Begin VB.Menu AreYouSure 
         Caption         =   "Are You Sure?"
         Begin VB.Menu Yes 
            Caption         =   "Yes"
         End
         Begin VB.Menu No 
            Caption         =   "No"
         End
      End
   End
End
Attribute VB_Name = "MasterProgrammerPopupMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Button = 2 Then
Exit Sub
Else
    PopupMenu MasterProgramerPopupMenu, 0, 300, 1100
End If
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu MasterProgramerPopupMenu, 0, 2100, 1100
End If
End Sub

Private Sub Command3_Click()
    PopupMenu ¤Exit, 0, 1150, 2650
End Sub

Private Sub Form_Load()
'Ok This Is Different Than Any Other Code To Make
'Popup Menus, This Gives You An Idea On How To
'Make Them Popup Where Ever You Want Them To.
'I Like To Make Them Popup Center Right Under The
'Button. It Kinda Hard To Get The Place Where You
'Want It But Its Cool And Cool Looking At The End.
'This Code Was Made By UGHHH So If You Use
'It Put Me In Greetz. If You Don't Your A BIG LOSER!
'Because If You Don't All You Are Is A Code Taking
'Loser. So If You Want To Be One Of The Cool Programers
'Put Me In Greetz Or Be A Loser!
            'E-Mail Me At MastaP5@aol.com
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Exit Sub
Else
    PopupMenu MasterProgramerPopupMenu, 0, 300, 1450
End If
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu MasterProgramerPopupMenu, 0, 2100, 1450
End If
End Sub

Private Sub No_Click()
    Exit Sub
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Exit Sub
Else
    PopupMenu MasterProgramerPopupMenu, 0, 300, 1880
End If
End Sub

Private Sub Text2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'I Would Say Don't Try Right Click On A Text Box
'Because The Other Popup Menu Comes Up.
'You Know The One That Says Copy, Cut, Select All
'And Stuff.
If Button = 2 Then
    PopupMenu MasterProgramerPopupMenu, 0, 2100, 1880
End If
End Sub

Private Sub Yes_Click()
    Unload Me
End Sub
