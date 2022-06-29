VERSION 4.00
Begin VB.Form Form10 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mass IMz"
   ClientHeight    =   3360
   ClientLeft      =   2820
   ClientTop       =   1875
   ClientWidth     =   5595
   Height          =   3765
   Icon            =   "Form10.frx":0000
   Left            =   2760
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   Top             =   1530
   Width           =   5715
   Begin VB.TextBox Text1 
      BackColor       =   &H000000C0&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MadAve"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form10.frx":030A
      Top             =   1800
      Width           =   1935
   End
   Begin VB.ListBox List1 
      BackColor       =   &H000000C0&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MadAve"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1230
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   2655
      Left            =   2640
      TabIndex        =   6
      Top             =   360
      Width           =   2415
      _Version        =   65536
      _ExtentX        =   4260
      _ExtentY        =   4683
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   4
      BorderWidth     =   2
      BevelInner      =   1
      Outline         =   -1  'True
      Begin VB.TextBox Text2 
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MadAve"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Text            =   "Name to Add"
         Top             =   1080
         Width           =   1935
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1920
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Start Mass IMz"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MadAve"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
      End
      Begin Threed.SSCommand SSCommand3 
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Clear List"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MadAve"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Font3D          =   2
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Add Name"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MadAve"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Font3D          =   2
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Add RooM"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MadAve"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Font3D          =   2
      End
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "IMz Sent 0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MadAve"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "M  e  s  s  a  g  e"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MadAve"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "N  a  m  e   s"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MadAve"
         Size            =   6.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Double Click to remove of List"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MadAve"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Number On List 0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MadAve"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "Form10"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()
StayOnTop Me
SendChat "<b><i><s><font Face= Arial>" & BlueBlackBlue("ØRëO ¹·° • ÍMz Š†á†úš • (v)ass ÍMz LÒaded •")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
SendChat "<b><i><s><font Face= Arial>" & BlueBlackBlue("ØRëO ¹·° • ÍMz Š†á†úš •(v)ass ÍMz Unloaded•")
Form2.Show
Unload Form10
End Sub


Private Sub List1_DblClick()
List1.RemoveItem List1.ListIndex
End Sub


Private Sub SSCommand1_Click()
Call AddRoomToListBox(List1)
Label1.caption = "Number on list: " & List1.ListCount
SendChat "<b><i><s><font Face= Arial>" & BlackRedBlack("ØRëO ¹·° • Š†á†úš • RooM Added •")
End Sub


Private Sub SSCommand2_Click()
If Text1 = "" Then
MsgBox "enter a screen name first", 64, "Enter Name"
Exit Sub
End If
List1.AddItem Text2
Text2 = ""
Label1.caption = "Number on List: " & List1.ListCount
End Sub


Private Sub SSCommand3_Click()
List1.Clear
Label1.caption = "Number on List: 0"
End Sub

Private Sub SSCommand4_Click()
Call IsUserOnline
If IsUserOnline = 0 Then
MsgBox "You are not signed on", 64, "Sign on"
Exit Sub
End If
Call IsUserOnline
If List1.ListCount = 0 Then
MsgBox "enter a name first", 64, "Enter SN"
Exit Sub
End If
If Text1.Text = "" Then
MsgBox "type a message to send", 64, "enter message"
Exit Sub
End If
For i = 0 To List1.ListCount - 1
a = List1.ListCount - i
B = List1.ListCount - a
C = B + 1
X = (Text1) + "<br>"
Y = ("<b><s><i><FONT COLOR=" + "#bbbbbb" + ">" + "(¯`·¸OrEO¹·º (V)ass I(V) # " & C & " Out of: " & List1.ListCount & " ¸·´¯)")
Call IMKeyword(List1.List(i), Text1 + Chr(10) + Chr(13) + Chr(10) + Chr(13) + "      " + Y)
Label5.caption = "IMs Sent: " & C & ""
Next i

End Sub





