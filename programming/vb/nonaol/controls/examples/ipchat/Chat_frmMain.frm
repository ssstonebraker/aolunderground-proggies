VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat Example By Mist"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   Icon            =   "Chat_frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   5220
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   4935
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Text            =   "Press enter to send text"
      Top             =   3960
      Width           =   5055
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   2010
      TabIndex        =   8
      Text            =   "Mr. Bob"
      Top             =   1020
      Width           =   2055
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   7
      Top             =   4440
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   2566
            Text            =   "  No connection...  "
            TextSave        =   "  No connection...  "
            Key             =   "STATUS"
            Object.Tag             =   ""
            Object.ToolTipText     =   "The current status of the connection"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   6562
            TextSave        =   ""
            Key             =   "DATA"
            Object.Tag             =   ""
            Object.ToolTipText     =   "The last data transfer through the modem"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtRemotePort 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   2010
      TabIndex        =   2
      Text            =   "5"
      Top             =   720
      Width           =   2985
   End
   Begin VB.TextBox txtLocalPort 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   2010
      TabIndex        =   1
      Text            =   "5"
      Top             =   420
      Width           =   2985
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   4
      Top             =   1080
      Width           =   885
   End
   Begin VB.TextBox txtRemoteIP 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   2010
      TabIndex        =   0
      Top             =   120
      Width           =   2985
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   600
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nick Name"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   90
      TabIndex        =   9
      Top             =   1020
      Width           =   1905
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Remote Port :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   90
      TabIndex        =   6
      Top             =   720
      Width           =   1905
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Local Port :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   90
      TabIndex        =   5
      Top             =   420
      Width           =   1875
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Connect with IP :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   90
      TabIndex        =   3
      Top             =   120
      Width           =   1905
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private IgnoreText As Boolean

Private Sub cmdClear_Click()
Text1 = ""
With Text2
.Text = " "
 .SetFocus
End With
End Sub

Private Sub cmdConnect_Click()


On Error GoTo ErrHandler

With Winsock1
 .RemoteHost = Trim(txtRemoteIP)
  .RemotePort = Trim(txtRemotePort)
   If .LocalPort = Empty Then
      .LocalPort = Trim(txtLocalPort)
      Frame2.Caption = .LocalIP
      .Bind .LocalPort
   End If
End With
txtLocalPort.Locked = True
StatusBar1.Panels(1).Text = "  Connected to " & Winsock1.RemoteHost & "  "

Frame1.Enabled = True
Frame2.Enabled = True
Label4.Visible = True

Text2.SetFocus
Exit Sub

ErrHandler:
MsgBox "Winsock failed to establish connection with remote server", vbCritical
End Sub

Private Sub Form_Load()
Show
MsgBox "Heres an example of how to talk over a internet conenction using the winsock bas. Put your friends IP address in the box and talk away. Only two people can be connected to the same IP. If you use this please give me credit", vbInformation, "Chat Example By Mist"
txtRemoteIP = Winsock1.LocalIP
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)


Static Last_Line_Feed As Long
Dim New_Line As String

If Trim(Text2) = vbNullString Then Last_Line_Feed = 0
If KeyAscii = 13 Then
KeyAscii = 0
   New_Line = Mid(Text2, Last_Line_Feed + 1)
   Last_Line_Feed = Text2.SelStart
   Winsock1.SendData (Text3.Text + ": " + New_Line + Chr$(13) + Chr$(10))
  Text2.Text = ""
   StatusBar1.Panels(2).Text = "  Sent " & (LenB(New_Line) / 2) & " bytes  "
End If

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

Dim New_Text As String

Winsock1.GetData New_Text
Text1.SelText = New_Text
Frame1.Caption = Winsock1.RemoteHostIP
StatusBar1.Panels(2).Text = "  Recieved " & bytesTotal & " bytes  "
End Sub

