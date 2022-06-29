VERSION 5.00
Object = "{33155A3D-0CE0-11D1-A6B4-444553540000}#1.0#0"; "SYSTRAY.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5370
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":08CA
   ScaleHeight     =   2145
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SysTray.SystemTray SystemTray1 
      Left            =   480
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      SysTrayText     =   ""
      IconFile        =   0
   End
   Begin VB.VScrollBar Blue2 
      Height          =   1095
      Left            =   4680
      Max             =   255
      TabIndex        =   16
      Top             =   600
      Width           =   135
   End
   Begin VB.VScrollBar Green2 
      Height          =   1095
      Left            =   4440
      Max             =   255
      TabIndex        =   15
      Top             =   600
      Width           =   135
   End
   Begin VB.VScrollBar Red2 
      Height          =   1095
      Left            =   4200
      Max             =   255
      TabIndex        =   14
      Top             =   600
      Width           =   135
   End
   Begin VB.PictureBox Color2 
      BackColor       =   &H80000012&
      Height          =   135
      Left            =   2880
      ScaleHeight     =   75
      ScaleWidth      =   1155
      TabIndex        =   13
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2880
      TabIndex        =   12
      Text            =   "http://fear99.cjb.net"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2880
      TabIndex        =   11
      Text            =   "Description"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2880
      TabIndex        =   10
      Text            =   "Person"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Text            =   "http://fear99.cjb.net"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Text            =   "Description"
      Top             =   720
      Width           =   1215
   End
   Begin VB.PictureBox Color1 
      BackColor       =   &H80000012&
      Height          =   135
      Left            =   1200
      ScaleHeight     =   75
      ScaleWidth      =   1155
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin VB.VScrollBar Blue1 
      Height          =   1095
      Left            =   840
      Max             =   255
      TabIndex        =   5
      Top             =   600
      Width           =   135
   End
   Begin VB.VScrollBar Green1 
      Height          =   1095
      Left            =   600
      Max             =   255
      TabIndex        =   4
      Top             =   600
      Width           =   135
   End
   Begin VB.VScrollBar Red1 
      Height          =   1095
      Left            =   360
      Max             =   255
      TabIndex        =   3
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Advertize"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   18
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Send IM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   17
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   2640
      X2              =   2640
      Y1              =   360
      Y2              =   2040
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Send It"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "illusionation Link Sender ¹·º"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Blue1_Change()
Color1.BackColor = RGB(Red1.Value, Green1.Value, Blue1.Value)
End Sub

Private Sub Blue2_Change()
Color2.BackColor = RGB(Red2.Value, Green2.Value, Blue2.Value)
End Sub

Private Sub Form_Load()
StayOnTop Me
FormTopLeft Me
ChatSend "<font face=""Arial Narrow""></B><I></U></S>" & BlackGreen("‹‹››•^v› illusionation Link Sender ¹·º")
ChatSend "<font face=""Arial Narrow""></B><I></U></S>" & BlackGreen("‹‹››•^v› By FeaR ")
ChatSend "<font face=""Arial Narrow""></B><I></U></S>" & BlackGreen("‹‹››•^v› Now Loaded ")
ChatSend "<font face=""Arial Narrow""></B><I></U></S>" & BlackGreen("‹‹››•^v› " & UserSN + " • " & TrimTime2 + " • " & TrimDate + "")
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MvFrm Me
End Sub

Private Sub Green1_Change()
Color1.BackColor = RGB(Red1.Value, Green1.Value, Blue1.Value)
End Sub

Private Sub Green2_Change()
Color2.BackColor = RGB(Red2.Value, Green2.Value, Blue2.Value)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MvFrm Me
End Sub

Private Sub Label2_Click()
SystemTray1.icon = Val(Form1.icon)
SystemTray1.SysTrayText = "Form1"
SystemTray1.Action = sys_Add
Form1.Hide
End Sub

Private Sub Label3_Click()
ChatSend "<font face=""Arial Narrow""></B><I></U></S>" & BlackGreen("‹‹››•^v› illusionation Link Sender ¹·º")
ChatSend "<font face=""Arial Narrow""></B><I></U></S>" & BlackGreen("‹‹››•^v› By FeaR ")
ChatSend "<font face=""Arial Narrow""></B><I></U></S>" & BlackGreen("‹‹››•^v› Now UnLoaded ")
ChatSend "<font face=""Arial Narrow""></B><I></U></S>" & BlackGreen("‹‹››•^v› " & UserSN + " • " & TrimTime2 + " • " & TrimDate + "")
End
End Sub

Private Sub Label4_Click()
ChatSend "<font face=""Arial Narrow""></B><I></U></S>" & BlackGreen("‹‹››•^v› illusionation Link Sender ¹·º")
ChatSend "<font face=""Arial Narrow""></B><I></U></S>" & BlackGreen("‹‹››•^v› By FeaR ")
ChatSend "<font face=""Arial Narrow""></B><I></U></S>" & BlackGreen("‹‹››•^v› Incoming Link")
TimeOut 2
FadedText$ = FadeByColor2(Color1.BackColor, Color2.BackColor, Text1.Text, False)
Let X = (FadedText$)
ChatSend ("< a href=" & Chr(34) & Text2.Text & Chr(34) & "></U></B></I>" & (X) & "</a>")
End Sub

Private Sub Label5_Click()
ChatSend "<font face=""Arial Narrow""></B><I></U></S>" & BlackGreen("‹‹››•^v› illusionation Link Sender ¹·º")
ChatSend "<font face=""Arial Narrow""></B><I></U></S>" & BlackGreen("‹‹››•^v› By FeaR ")
ChatSend "<font face=""Arial Narrow""></B><I></U></S>" & BlackGreen("‹‹››•^v› Sending IM Link To " & Text3.Text + "")
TimeOut 2
FadedText$ = FadeByColor2(Color1.BackColor, Color2.BackColor, Text4.Text, False)
Let X = (FadedText$)
Call IM_Keyword(Text3.Text, "< a href=" & Chr(34) & Text5.Text & Chr(34) & "></U></B></I>" & (X) & "</a>")
End Sub

Private Sub Label6_Click()
ChatSend "<font face=""Arial Narrow""></B><I></U></S>" & BlackGreen("¸‚.-·~¬ˆ‘´¸‚·ª˜¨˜ª·,¸`ˆ'¬~·-.,¸")
ChatSend "<font face=""Arial Narrow""></B><I></U></S>" & BlackGreen("  `·.,¸¸,.·´`·.,¸¸,.·´`·.,¸¸,.·´")
ChatSend "<font face=""Arial Narrow""></B><I></U></S>" & BlackGreen("   ¸,.·´ illusionation  `·.,¸")
ChatSend "<font face=""Arial Narrow""></B><I></U></S>" & BlackGreen("   ¯\_  Link Sender   _/¯")
ChatSend "<font face=""Arial Narrow""></B><I></U></S>" & BlackGreen("       '·.¸,.¸‚·ª˜¨˜ª·,¸.¸,¸.·'")
ChatSend "<font face=""Arial Narrow""></B><I></U></S>" & BlackGreen(" ")
ChatSend "<font face=""Arial Narrow""></B><I></U></S>" & BlackGreen("            By FeaR")
End Sub

Private Sub Red1_Change()
Color1.BackColor = RGB(Red1.Value, Green1.Value, Blue1.Value)
End Sub

Private Sub Red2_Change()
Color2.BackColor = RGB(Red2.Value, Green2.Value, Blue2.Value)
End Sub

Private Sub SystemTray1_MouseDblClk(ByVal Button As Integer)
Form1.Show
End Sub
