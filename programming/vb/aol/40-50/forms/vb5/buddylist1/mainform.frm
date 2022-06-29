VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Status: Offline"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3420
   ControlBox      =   0   'False
   Icon            =   "mainform.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "mainform.frx":0442
   ScaleHeight     =   4020
   ScaleWidth      =   3420
   Begin VB.CommandButton Command30 
      Caption         =   "Command30"
      Height          =   375
      Left            =   120
      TabIndex        =   80
      Top             =   4920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   4440
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   5880
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3000
      Top             =   4200
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Command29"
      Height          =   255
      Left            =   3120
      TabIndex        =   79
      Top             =   4680
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Timer Timer4 
      Interval        =   100
      Left            =   1920
      Top             =   1800
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   2520
      TabIndex        =   78
      Text            =   "35"
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Show Bottom"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   74
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Hide Bottom"
      Height          =   375
      Left            =   1920
      TabIndex        =   75
      Top             =   3600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text11 
      Height          =   1575
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   73
      Text            =   "mainform.frx":2FE0
      Top             =   4320
      Width           =   2565
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   720
      MaxLength       =   15
      TabIndex        =   70
      Top             =   4920
      Width           =   2055
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   720
      MaxLength       =   10
      TabIndex        =   69
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Host This Chat"
      Height          =   495
      Left            =   1800
      TabIndex        =   68
      ToolTipText     =   "Host the Chat"
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Join this Chat"
      Height          =   495
      Left            =   720
      TabIndex        =   67
      ToolTipText     =   "Join the Chat "
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Host/Join Chat"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   66
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton Command23 
      Caption         =   "<-"
      Height          =   735
      Left            =   2880
      TabIndex        =   60
      Top             =   2880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1920
      Top             =   360
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Get Message of the Day"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   59
      ToolTipText     =   "Download the Message of the Day from Concept"
      Top             =   3240
      Width           =   2655
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Define Your Location/Message"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   58
      ToolTipText     =   "Define your location/message to others"
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   3480
      MaxLength       =   45
      TabIndex        =   57
      Text            =   "Your Location"
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CommandButton Command20 
      Caption         =   "->"
      Height          =   735
      Left            =   2880
      TabIndex        =   56
      Top             =   2880
      Width           =   495
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   1920
      Top             =   1320
   End
   Begin VB.CommandButton Command18 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   2400
      TabIndex        =   52
      ToolTipText     =   "Exit the Program"
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton Command19 
      Caption         =   "&Min"
      Height          =   375
      Left            =   1920
      TabIndex        =   53
      ToolTipText     =   "Minimize the window"
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Check &All"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   51
      ToolTipText     =   "check Status of All Buddys"
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command16 
      Caption         =   "About...."
      Height          =   375
      Left            =   120
      TabIndex        =   54
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton butClose 
      Caption         =   "&Disconnect"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   26
      ToolTipText     =   "Click Here to Disconnect form the AW Buddy Server"
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "&Connect"
      Height          =   375
      Left            =   120
      TabIndex        =   41
      ToolTipText     =   "Click Here to connect to the AW Buddy Server"
      Top             =   2880
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1920
      Top             =   840
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Check If Online"
      Enabled         =   0   'False
      Height          =   255
      Left            =   7320
      TabIndex        =   48
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Check If Online"
      Enabled         =   0   'False
      Height          =   255
      Left            =   7320
      TabIndex        =   47
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Check If Online"
      Enabled         =   0   'False
      Height          =   255
      Left            =   7320
      TabIndex        =   46
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Check If Online"
      Enabled         =   0   'False
      Height          =   255
      Left            =   7320
      TabIndex        =   45
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Check If Online"
      Enabled         =   0   'False
      Height          =   255
      Left            =   7320
      TabIndex        =   44
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   7320
      TabIndex        =   43
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Check If Online"
      Enabled         =   0   'False
      Height          =   255
      Left            =   7320
      TabIndex        =   42
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   960
      MaxLength       =   12
      TabIndex        =   40
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Semantec"
      Height          =   852
      Left            =   3120
      TabIndex        =   22
      Top             =   9000
      Width           =   2172
      Begin VB.OptionButton Option4 
         Caption         =   "Passive"
         Height          =   252
         Left            =   240
         TabIndex        =   24
         Top             =   480
         Width           =   1092
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Active (default)"
         Height          =   252
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Value           =   -1  'True
         Width           =   1452
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Dir"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   7080
      Width           =   1212
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Binary"
      Height          =   255
      Left            =   1320
      TabIndex        =   21
      Top             =   9240
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "ASCII"
      Height          =   255
      Left            =   1320
      TabIndex        =   20
      Top             =   9480
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "SetCurrentDir"
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   8040
      Width           =   1212
   End
   Begin VB.CommandButton Command6 
      Caption         =   "GetCurrentDir"
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   7560
      Width           =   1212
   End
   Begin VB.CommandButton Command5 
      Caption         =   "GetFile"
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   9000
      Width           =   1212
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Put Large File "
      Height          =   372
      Left            =   5400
      TabIndex        =   11
      Top             =   9480
      Width           =   1212
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Connect"
      Height          =   372
      Left            =   5400
      TabIndex        =   3
      Top             =   6600
      Width           =   1212
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   372
      Left            =   5400
      TabIndex        =   12
      Top             =   9960
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PutFile"
      Height          =   372
      Left            =   5400
      TabIndex        =   9
      Top             =   8520
      Width           =   1212
   End
   Begin VB.TextBox Text5 
      Height          =   372
      Left            =   1080
      TabIndex        =   6
      Top             =   8520
      Width           =   4212
   End
   Begin VB.TextBox Text4 
      Height          =   372
      Left            =   1080
      TabIndex        =   8
      Text            =   "c:\windows\"
      Top             =   8040
      Width           =   4212
   End
   Begin VB.TextBox Text3 
      Height          =   372
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "trats"
      Top             =   7560
      Width           =   4212
   End
   Begin VB.TextBox Text2 
      Height          =   372
      Left            =   1080
      TabIndex        =   1
      Text            =   "AWbuddy"
      Top             =   7080
      Width           =   4212
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   1080
      TabIndex        =   0
      Text            =   "ftp.xoom.com"
      Top             =   6600
      Width           =   4212
   End
   Begin VB.Frame Frame2 
      Caption         =   "Transfer Type"
      Height          =   852
      Left            =   1080
      TabIndex        =   25
      Top             =   9000
      Width           =   1932
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   3600
      TabIndex        =   77
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "Message of the Day:"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3960
      TabIndex        =   76
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H000000FF&
      Height          =   1815
      Left            =   600
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Chat Room:"
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   720
      TabIndex        =   72
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "NickName:"
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   720
      TabIndex        =   71
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   135
      Left            =   3000
      Top             =   1560
      Width           =   375
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   135
      Left            =   3000
      Top             =   960
      Width           =   375
   End
   Begin VB.Line Line12 
      BorderColor     =   &H000000FF&
      X1              =   2880
      X2              =   3480
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   3600
      TabIndex        =   65
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   3600
      TabIndex        =   64
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   3600
      TabIndex        =   63
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   3600
      TabIndex        =   62
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   3600
      TabIndex        =   61
      Top             =   360
      Width           =   2415
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      Height          =   135
      Left            =   3360
      Shape           =   2  'Oval
      Top             =   1320
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      Height          =   135
      Left            =   3240
      Shape           =   2  'Oval
      Top             =   1200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      Height          =   135
      Left            =   3120
      Shape           =   2  'Oval
      Top             =   1320
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      Height          =   135
      Left            =   3000
      Shape           =   2  'Oval
      Top             =   1200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      Height          =   135
      Left            =   2880
      Shape           =   2  'Oval
      Top             =   1320
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Line Line11 
      BorderColor     =   &H000000FF&
      X1              =   2880
      X2              =   3480
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Location/Message:"
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   3960
      TabIndex        =   55
      Top             =   0
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   2175
      Left            =   3480
      Top             =   240
      Width           =   2625
   End
   Begin VB.Line Line10 
      BorderColor     =   &H000000FF&
      X1              =   0
      X2              =   1440
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line9 
      BorderColor     =   &H000000FF&
      X1              =   0
      X2              =   1440
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line8 
      BorderColor     =   &H000000FF&
      X1              =   0
      X2              =   1440
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      X1              =   0
      X2              =   1440
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      X1              =   0
      X2              =   1440
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      X1              =   0
      X2              =   0
      Y1              =   240
      Y2              =   2400
   End
   Begin VB.Image Image1 
      Height          =   555
      Left            =   720
      Picture         =   "mainform.frx":3010
      Top             =   2760
      Width           =   1500
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Status of Buddy:"
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   1560
      TabIndex        =   50
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label21 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   120
      TabIndex        =   49
      Top             =   2520
      Width           =   975
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      X1              =   -2040
      X2              =   3480
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   2880
      X2              =   2880
      Y1              =   240
      Y2              =   2400
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   1440
      X2              =   1440
      Y1              =   240
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   -2040
      X2              =   2880
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label20 
      BackColor       =   &H80000012&
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   1920
      TabIndex        =   39
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label19 
      BackColor       =   &H80000012&
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   1920
      TabIndex        =   38
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label18 
      BackColor       =   &H80000012&
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   1920
      TabIndex        =   37
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label17 
      BackColor       =   &H80000012&
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   1920
      TabIndex        =   36
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label16 
      BackColor       =   &H80000012&
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   1920
      TabIndex        =   35
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label15 
      BackColor       =   &H80000012&
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   1920
      TabIndex        =   34
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000012&
      Caption         =   "none"
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000012&
      Caption         =   "none"
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000012&
      Caption         =   "none"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000012&
      Caption         =   "none"
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000012&
      Caption         =   "none"
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000012&
      Caption         =   "Concept"
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000007&
      Caption         =   "Buddys:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label7 
      Height          =   375
      Left            =   2880
      TabIndex        =   19
      Top             =   9960
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "Large Transfer   Byte Count:"
      Height          =   375
      Left            =   1080
      TabIndex        =   18
      Top             =   9960
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      Caption         =   "Remote File or Directory"
      Height          =   615
      Left            =   240
      TabIndex        =   17
      Top             =   8400
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Local File:"
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   8040
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Password:"
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   7560
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "User:"
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Top             =   7080
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Server:"
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   6600
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   2085
      Left            =   3480
      Picture         =   "mainform.frx":5BAE
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   4005
   End
   Begin VB.Image Image3 
      Height          =   600
      Left            =   3480
      Picture         =   "mainform.frx":324DC
      Top             =   3480
      Width           =   2550
   End
   Begin VB.Menu buddy1 
      Caption         =   "buddy1"
      Visible         =   0   'False
      Begin VB.Menu check1 
         Caption         =   "Check if Online"
      End
      Begin VB.Menu asdfee 
         Caption         =   "Check Location/Message"
      End
      Begin VB.Menu define1 
         Caption         =   "Define Buddy"
      End
   End
   Begin VB.Menu buddy2 
      Caption         =   "buddy2"
      Visible         =   0   'False
      Begin VB.Menu a 
         Caption         =   "Check if Online"
      End
      Begin VB.Menu asdfeeee 
         Caption         =   "Check Location/Message"
      End
      Begin VB.Menu dt 
         Caption         =   "Define Buddy"
      End
   End
   Begin VB.Menu buddy3 
      Caption         =   "buddy3"
      Visible         =   0   'False
      Begin VB.Menu aweraw 
         Caption         =   "Check if Online"
      End
      Begin VB.Menu aweraaaaa 
         Caption         =   "Check Location/Message"
      End
      Begin VB.Menu afgg 
         Caption         =   "Define Buddy"
      End
   End
   Begin VB.Menu buddy4 
      Caption         =   "buddy4"
      Visible         =   0   'False
      Begin VB.Menu asdf 
         Caption         =   "Check if Online"
      End
      Begin VB.Menu aeeeaaa 
         Caption         =   "Check Location/Message"
      End
      Begin VB.Menu wetaw 
         Caption         =   "Define Buddy"
      End
   End
   Begin VB.Menu buddy5 
      Caption         =   "buddy5"
      Visible         =   0   'False
      Begin VB.Menu aweradf 
         Caption         =   "Check if Online"
      End
      Begin VB.Menu asdfwaeawr 
         Caption         =   "Check Location/Message"
      End
      Begin VB.Menu werasdf 
         Caption         =   "Define Buddy"
      End
   End
   Begin VB.Menu buddy6 
      Caption         =   "buddy6"
      Visible         =   0   'False
      Begin VB.Menu asd 
         Caption         =   "Check if Online"
      End
      Begin VB.Menu weee 
         Caption         =   "Check Location/Message"
      End
      Begin VB.Menu aweras 
         Caption         =   "Define Buddy"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''''''''''''''''''''''''LOTS OF THANKS TO AARON
''''''''''''''''''''''''YOUNG AT REDWING FOR
''''''''''''''''''''''''TEACHING ME THIS CODE
''''''''''''''''''''''''
Private Type RECT      '
        Left As Long   '
        Top As Long    '
        Right As Long  '''''''''RECTANGLE TYPE
        Bottom As Long '
End Type               '
''''''''''''''''''''''''
''''''''''''''''''''''''GETS WINDOW RECTANGULAR DIMENSIONS
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

''''''''''''''''''''''''GETS A WINDOW'S HANDLE (TASKBAR)
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

''''''''''''''''''''''''OUR RECT TYPE (PUBLIC/PRIVATE/GLOBAL)
Private tRect As RECT
Dim hOpen As Long, hConnection As Long, hFile As Long
Dim dwType As Long
Dim dwSeman As Long

Private Sub ErrorOut(ByVal dwError As Long, ByRef szFunc As String)
Dim dwRet As Long
Dim dwTemp As Long
Dim szString As String * 2048
Dim szErrorMessage As String

dwRet = FormatMessage(FORMAT_MESSAGE_FROM_HMODULE, _
                  GetModuleHandle("wininet.dll"), dwError, 0, _
                  szString, 256, 0)
szErrorMessage = szFunc & " error code: " & dwError & " Message: " & szString
Debug.Print szErrorMessage
MsgBox szErrorMessage, , "SimpleFtp"
If (dwError = 12003) Then
    ' Extended error information was returned
    dwRet = InternetGetLastResponseInfo(dwTemp, szString, 2048)
    Debug.Print szString
    Form2.Show
    Form2.Text1.text = szString
End If
End Sub

Private Sub a_Click()
If Label10.Caption = "none" Then

Exit Sub
End If
' for ASCII files use FTP_TRANSFER_TYPE_ASCII
Text4.text = "c:\windows\" + LCase(Label10.Caption) + ".txt"
Text5.text = LCase(Label10.Caption) + ".txt"
' add INTERNET_FLAG_NO_CACHE_WRITE to avoid local caching 0x04000000 (hex)
 If (FtpGetFile(hConnection, Text5.text, Text4.text, False, _
         FILE_ATTRIBUTE_NORMAL, dwType Or INTERNET_FLAG_RELOAD, 0) = False) Then
   MsgBox "Sorry, that user has never logged into Buddy List before.", 12, "No such user."
    Label16.Caption = "  n/a"
    Exit Sub
 Else
 
   End If
   
   LoadText Text7, "c:\windows\" + LCase(Label10.Caption) + ".txt"
   Label16.Caption = Text7.text
End Sub

Private Sub aeeeaaa_Click()
If Label12.Caption = "none" Then
Exit Sub
End If
' for ASCII files use FTP_TRANSFER_TYPE_ASCII
Text4.text = "c:\windows\" + LCase(Label12.Caption) + "loc.txt"
Text5.text = LCase(Label12.Caption) + "loc.txt"
' add INTERNET_FLAG_NO_CACHE_WRITE to avoid local caching 0x04000000 (hex)
 If (FtpGetFile(hConnection, Text5.text, Text4.text, False, _
         FILE_ATTRIBUTE_NORMAL, dwType Or INTERNET_FLAG_RELOAD, 0) = False) Then
    MsgBox "Sorry, this user has not specfied a location/message..", 12, "No location/message"
     Label27.Caption = "  n/a"
    Exit Sub
 Else
 
   End If
   
   LoadText Text7, "c:\windows\" + LCase(Label12.Caption) + "loc.txt"
   Label27.Caption = Text7.text
End Sub

Private Sub afgg_Click()
Dim message, Title, Default, MyValue
message = "Please enter your buddy's name:"   ' Set prompt.
Title = "Name?" ' Set title.
Default = ""   ' Set default.
' Display message, title, and default value.
MyValue = InputBox(message, Title, Default)

Label11.Caption = MyValue
End Sub

Private Sub asd_Click()
If Label14.Caption = "none" Then
Exit Sub
End If
' for ASCII files use FTP_TRANSFER_TYPE_ASCII
Text4.text = "c:\windows\" + LCase(Label14.Caption) + ".txt"
Text5.text = LCase(Label14.Caption) + ".txt"
' add INTERNET_FLAG_NO_CACHE_WRITE to avoid local caching 0x04000000 (hex)
 If (FtpGetFile(hConnection, Text5.text, Text4.text, False, _
         FILE_ATTRIBUTE_NORMAL, dwType Or INTERNET_FLAG_RELOAD, 0) = False) Then
     MsgBox "Sorry, that user has never logged into Buddy List before.", 12, "No such user."
      Label20.Caption = "  n/a"
    Exit Sub
 Else
 
   End If
   
   LoadText Text7, "c:\windows\" + LCase(Label14.Caption) + ".txt"
   Label20.Caption = Text7.text
End Sub

Private Sub asdf_Click()
If Label12.Caption = "none" Then

Exit Sub
End If
' for ASCII files use FTP_TRANSFER_TYPE_ASCII
Text4.text = "c:\windows\" + LCase(Label12.Caption) + ".txt"
Text5.text = LCase(Label12.Caption) + ".txt"
' add INTERNET_FLAG_NO_CACHE_WRITE to avoid local caching 0x04000000 (hex)
 If (FtpGetFile(hConnection, Text5.text, Text4.text, False, _
         FILE_ATTRIBUTE_NORMAL, dwType Or INTERNET_FLAG_RELOAD, 0) = False) Then
     MsgBox "Sorry, that user has never logged into Buddy List before.", 12, "No such user."
      Label18.Caption = "  n/a"
    Exit Sub
 Else
 
   End If
   
   LoadText Text7, "c:\windows\" + LCase(Label12.Caption) + ".txt"
   Label18.Caption = Text7.text
End Sub

Private Sub asdfee_Click()
If Label9.Caption = "none" Then
Exit Sub
End If
' for ASCII files use FTP_TRANSFER_TYPE_ASCII
Text4.text = "c:\windows\" + LCase(Label9.Caption) + "loc.txt"
Text5.text = LCase(Label9.Caption) + "loc.txt"
' add INTERNET_FLAG_NO_CACHE_WRITE to avoid local caching 0x04000000 (hex)
 If (FtpGetFile(hConnection, Text5.text, Text4.text, False, _
         FILE_ATTRIBUTE_NORMAL, dwType Or INTERNET_FLAG_RELOAD, 0) = False) Then
    MsgBox "Sorry, this user has not specfied a location/message..", 12, "No location/message"
     Label24.Caption = "  n/a"
    Exit Sub
 Else
 
   End If
   
   LoadText Text7, "c:\windows\" + LCase(Label9.Caption) + "loc.txt"
   Label24.Caption = Text7.text
End Sub

Private Sub asdfeeee_Click()
If Label10.Caption = "none" Then
Exit Sub
End If
' for ASCII files use FTP_TRANSFER_TYPE_ASCII
Text4.text = "c:\windows\" + LCase(Label10.Caption) + "loc.txt"
Text5.text = LCase(Label10.Caption) + "loc.txt"
' add INTERNET_FLAG_NO_CACHE_WRITE to avoid local caching 0x04000000 (hex)
 If (FtpGetFile(hConnection, Text5.text, Text4.text, False, _
         FILE_ATTRIBUTE_NORMAL, dwType Or INTERNET_FLAG_RELOAD, 0) = False) Then
    MsgBox "Sorry, this user has not specfied a location/message..", 12, "No location/message"
     Label25.Caption = "  n/a"
    Exit Sub
 Else
 
   End If
   
   LoadText Text7, "c:\windows\" + LCase(Label10.Caption) + "loc.txt"
   Label25.Caption = Text7.text
End Sub

Private Sub asdfwaeawr_Click()
If Label13.Caption = "none" Then
Exit Sub
End If
' for ASCII files use FTP_TRANSFER_TYPE_ASCII
Text4.text = "c:\windows\" + LCase(Label13.Caption) + "loc.txt"
Text5.text = LCase(Label13.Caption) + "loc.txt"
' add INTERNET_FLAG_NO_CACHE_WRITE to avoid local caching 0x04000000 (hex)
 If (FtpGetFile(hConnection, Text5.text, Text4.text, False, _
         FILE_ATTRIBUTE_NORMAL, dwType Or INTERNET_FLAG_RELOAD, 0) = False) Then
    MsgBox "Sorry, this user has not specfied a location/message..", 12, "No location/message"
     Label28.Caption = "  n/a"
    Exit Sub
 Else
 
   End If
   
   LoadText Text7, "c:\windows\" + LCase(Label13.Caption) + "loc.txt"
   Label28.Caption = Text7.text
End Sub

Private Sub aweraaaaa_Click()
If Label11.Caption = "none" Then
Exit Sub
End If
' for ASCII files use FTP_TRANSFER_TYPE_ASCII
Text4.text = "c:\windows\" + LCase(Label11.Caption) + "loc.txt"
Text5.text = LCase(Label11.Caption) + "loc.txt"
' add INTERNET_FLAG_NO_CACHE_WRITE to avoid local caching 0x04000000 (hex)
 If (FtpGetFile(hConnection, Text5.text, Text4.text, False, _
         FILE_ATTRIBUTE_NORMAL, dwType Or INTERNET_FLAG_RELOAD, 0) = False) Then
   MsgBox "Sorry, this user has not specfied a location/message..", 12, "No location/message"
     Label26.Caption = "  n/a"
    Exit Sub
 Else
 
   End If
   
   LoadText Text7, "c:\windows\" + LCase(Label11.Caption) + "loc.txt"
   Label26.Caption = Text7.text
End Sub

Private Sub aweradf_Click()
If Label13.Caption = "none" Then
Exit Sub
End If
' for ASCII files use FTP_TRANSFER_TYPE_ASCII
Text4.text = "c:\windows\" + LCase(Label13.Caption) + ".txt"
Text5.text = LCase(Label13.Caption) + ".txt"
' add INTERNET_FLAG_NO_CACHE_WRITE to avoid local caching 0x04000000 (hex)
 If (FtpGetFile(hConnection, Text5.text, Text4.text, False, _
         FILE_ATTRIBUTE_NORMAL, dwType Or INTERNET_FLAG_RELOAD, 0) = False) Then
     MsgBox "Sorry, that user has never logged into Buddy List before.", 12, "No such user."
      Label19.Caption = "  n/a"
    Exit Sub
 Else
 
   End If
   
   LoadText Text7, "c:\windows\" + LCase(Label13.Caption) + ".txt"
   Label19.Caption = Text7.text
End Sub

Private Sub aweras_Click()
Dim message, Title, Default, MyValue
message = "Please enter your buddy's name:"   ' Set prompt.
Title = "Name?" ' Set title.
Default = ""   ' Set default.
' Display message, title, and default value.
MyValue = InputBox(message, Title, Default)

Label14.Caption = MyValue
End Sub

Private Sub aweraw_Click()
If Label11.Caption = "none" Then

Exit Sub
End If
' for ASCII files use FTP_TRANSFER_TYPE_ASCII
Text4.text = "c:\windows\" + LCase(Label11.Caption) + ".txt"
Text5.text = LCase(Label11.Caption) + ".txt"
' add INTERNET_FLAG_NO_CACHE_WRITE to avoid local caching 0x04000000 (hex)
 If (FtpGetFile(hConnection, Text5.text, Text4.text, False, _
         FILE_ATTRIBUTE_NORMAL, dwType Or INTERNET_FLAG_RELOAD, 0) = False) Then
 MsgBox "Sorry, that user has never logged into Buddy List before.", 12, "No such user."
  Label17.Caption = "  n/a"
    Exit Sub
 Else
 
   End If
   
   LoadText Text7, "c:\windows\" + LCase(Label11.Caption) + ".txt"
   Label17.Caption = Text7.text
End Sub

Private Sub butClose_Click()
Text6.text = LCase(Text6.text)
Form1.MousePointer = "11"
Form1.Caption = "Status: Disconnecting...please wait."
Text7.text = "offline"
SaveText Text7, "c:\windows\" + LCase(Text6.text) + ".txt"
Text4.text = "c:\windows\" + LCase(Text6.text) + ".txt"
Text5.text = Text6.text + ".txt"
 If (FtpPutFile(hConnection, Text4.text, Text5.text, _
         dwType, 0) = False) Then
             MsgBox "There was an error disconnecting from the server." + Chr(13) + "Please try re-connecting then disconnecting, or exit.", vbCritical, "Error"
             butClose.Enabled = False
             Command9.Enabled = True
             Command18.Enabled = False
              Form1.MousePointer = "0"
             Exit Sub
        Else
   
       Form1.MousePointer = "0"
       Timer3.Enabled = False
       
MsgBox "Disconnected from Buddy List Server.", 12, "Disconnected"
Form1.Caption = "Status: " + Text6.text + " logged out "
Command10.Enabled = False
       Command11.Enabled = False
       Command12.Enabled = False
       Command13.Enabled = False
       Command14.Enabled = False
       Command15.Enabled = False
       Shape2.Visible = False
        End If
    If hConnection <> 0 Then
        InternetCloseHandle hConnection
    End If
    hConnection = 0
       Command9.Enabled = True
    butClose.Enabled = False
    Command17.Enabled = False
     Command18.Enabled = True
     Command21.Enabled = False
     Command22.Enabled = False
     Command24.Enabled = False
     Command27.Enabled = False
     Form1.Height = "4365"
     End Sub

Private Sub check1_Click()
If Label9.Caption = "none" Then

Exit Sub
End If
' for ASCII files use FTP_TRANSFER_TYPE_ASCII
Text4.text = "c:\windows\" + LCase(Label9.Caption) + ".txt"
Text5.text = LCase(Label9.Caption) + ".txt"
' add INTERNET_FLAG_NO_CACHE_WRITE to avoid local caching 0x04000000 (hex)
 If (FtpGetFile(hConnection, Text5.text, Text4.text, False, _
         FILE_ATTRIBUTE_NORMAL, dwType Or INTERNET_FLAG_RELOAD, 0) = False) Then
    MsgBox "Sorry, that user has never logged into Buddy List before.", 12, "No such user."
     Label15.Caption = "  n/a"
    Exit Sub
 Else
 
   End If
   
   LoadText Text7, "c:\windows\" + LCase(Label9.Caption) + ".txt"
   Label15.Caption = Text7.text
End Sub

Private Sub Command1_Click()
' for ASCII files use FTP_TRANSFER_TYPE_ASCII
        If (FtpPutFile(hConnection, Text4.text, Text5.text, _
         dwType, 0) = False) Then
             ErrorOut Err.LastDllError, "FtpPutFile"
             Exit Sub
        Else
         MsgBox "File transfered!", , "Simple Ftp"
        End If

End Sub
Private Sub Command10_Click()
If Label9.Caption = "none" Then

Exit Sub
End If
' for ASCII files use FTP_TRANSFER_TYPE_ASCII
Text4.text = "c:\windows\" + LCase(Label9.Caption) + ".txt"
Text5.text = LCase(Label9.Caption) + ".txt"
' add INTERNET_FLAG_NO_CACHE_WRITE to avoid local caching 0x04000000 (hex)
 If (FtpGetFile(hConnection, Text5.text, Text4.text, False, _
         FILE_ATTRIBUTE_NORMAL, dwType Or INTERNET_FLAG_RELOAD, 0) = False) Then
    MsgBox "Sorry, that user has never logged into Buddy List before.", 12, "No such user."
    Label15.Caption = "  n/a"
    Exit Sub
 Else
 
   End If
   
   LoadText Text7, "c:\windows\" + LCase(Label9.Caption) + ".txt"
   Label15.Caption = Text7.text
End Sub

Private Sub Command11_Click()
If Label10.Caption = "none" Then

Exit Sub
End If
' for ASCII files use FTP_TRANSFER_TYPE_ASCII
Text4.text = "c:\windows\" + LCase(Label10.Caption) + ".txt"
Text5.text = LCase(Label10.Caption) + ".txt"
' add INTERNET_FLAG_NO_CACHE_WRITE to avoid local caching 0x04000000 (hex)
 If (FtpGetFile(hConnection, Text5.text, Text4.text, False, _
         FILE_ATTRIBUTE_NORMAL, dwType Or INTERNET_FLAG_RELOAD, 0) = False) Then
   MsgBox "Sorry, that user has never logged into Buddy List before.", 12, "No such user."
       Label16.Caption = "  n/a"
    Exit Sub
 Else
 
   End If
   
   LoadText Text7, "c:\windows\" + LCase(Label10.Caption) + ".txt"
   Label16.Caption = Text7.text
End Sub

Private Sub Command12_Click()
If Label11.Caption = "none" Then

Exit Sub
End If
' for ASCII files use FTP_TRANSFER_TYPE_ASCII
Text4.text = "c:\windows\" + LCase(Label11.Caption) + ".txt"
Text5.text = LCase(Label11.Caption) + ".txt"
' add INTERNET_FLAG_NO_CACHE_WRITE to avoid local caching 0x04000000 (hex)
 If (FtpGetFile(hConnection, Text5.text, Text4.text, False, _
         FILE_ATTRIBUTE_NORMAL, dwType Or INTERNET_FLAG_RELOAD, 0) = False) Then
 MsgBox "Sorry, that user has never logged into Buddy List before.", 12, "No such user."
     Label17.Caption = "  n/a"
    Exit Sub
 Else
 
   End If
   
   LoadText Text7, "c:\windows\" + LCase(Label11.Caption) + ".txt"
   Label17.Caption = Text7.text
End Sub

Private Sub Command13_Click()
If Label12.Caption = "none" Then

Exit Sub
End If
' for ASCII files use FTP_TRANSFER_TYPE_ASCII
Text4.text = "c:\windows\" + LCase(Label12.Caption) + ".txt"
Text5.text = LCase(Label12.Caption) + ".txt"
' add INTERNET_FLAG_NO_CACHE_WRITE to avoid local caching 0x04000000 (hex)
 If (FtpGetFile(hConnection, Text5.text, Text4.text, False, _
         FILE_ATTRIBUTE_NORMAL, dwType Or INTERNET_FLAG_RELOAD, 0) = False) Then
     MsgBox "Sorry, that user has never logged into Buddy List before.", 12, "No such user."
         Label18.Caption = "  n/a"
    Exit Sub
 Else
 
   End If
   
   LoadText Text7, "c:\windows\" + LCase(Label12.Caption) + ".txt"
   Label18.Caption = Text7.text
End Sub

Private Sub Command14_Click()
If Label13.Caption = "none" Then
Exit Sub
End If
' for ASCII files use FTP_TRANSFER_TYPE_ASCII
Text4.text = "c:\windows\" + LCase(Label13.Caption) + ".txt"
Text5.text = LCase(Label13.Caption) + ".txt"
' add INTERNET_FLAG_NO_CACHE_WRITE to avoid local caching 0x04000000 (hex)
 If (FtpGetFile(hConnection, Text5.text, Text4.text, False, _
         FILE_ATTRIBUTE_NORMAL, dwType Or INTERNET_FLAG_RELOAD, 0) = False) Then
     MsgBox "Sorry, that user has never logged into Buddy List before.", 12, "No such user."
         Label19.Caption = "  n/a"
    Exit Sub
 Else
 
   End If
   
   LoadText Text7, "c:\windows\" + LCase(Label13.Caption) + ".txt"
   Label19.Caption = Text7.text
End Sub

Private Sub Command15_Click()
If Label14.Caption = "none" Then
Exit Sub
End If
' for ASCII files use FTP_TRANSFER_TYPE_ASCII
Text4.text = "c:\windows\" + LCase(Label14.Caption) + ".txt"
Text5.text = LCase(Label14.Caption) + ".txt"
' add INTERNET_FLAG_NO_CACHE_WRITE to avoid local caching 0x04000000 (hex)
 If (FtpGetFile(hConnection, Text5.text, Text4.text, False, _
         FILE_ATTRIBUTE_NORMAL, dwType Or INTERNET_FLAG_RELOAD, 0) = False) Then
     MsgBox "Sorry, that user has never logged into Buddy List before.", 12, "No such user."
         Label20.Caption = "  n/a"
    Exit Sub
 Else
 
   End If
   
   LoadText Text7, "c:\windows\" + LCase(Label14.Caption) + ".txt"
   Label20.Caption = Text7.text
End Sub

Private Sub Command16_Click()
MsgBox "Buddy List 1.0" + Chr(13) + "By: Concept" + Chr(13) + "(C) 1999 C.A.T. Inc.", 12, "About.."
MsgBox "To add a buddy, just:" + Chr(13) + "1. Right click and empty slot." + Chr(13) + "2. Choose 'Define Buddy'" + Chr(13) + "3. Type in Buddy's name and press enter", 12, "About.."
MsgBox "As of now, you can only have 6 buddys.. I'll add more if needed. Also.. people could log in as you.. so only give your Buddy List user name to people you trust.", 12, "About..."
End Sub

Private Sub Command17_Click()
MsgBox "Please Be Patient while all buddys are checked...", 12, "Please hold on."
Form1.Caption = "Status: Checking Buddy 1...."
Call Command10_Click
Form1.Caption = "Status: Checking Buddy 2...."
Call Command11_Click
Form1.Caption = "Status: Checking Buddy 3...."
Call Command12_Click
Form1.Caption = "Status: Checking Buddy 4...."
Call Command13_Click
Form1.Caption = "Status: Checking Buddy 5...."
Call Command14_Click
Form1.Caption = "Status: Checking Buddy 6...."
Call Command15_Click
Form1.Caption = "Status: " + Text6.text + " Logged in"
End Sub

Private Sub Command18_Click()
Form1.Caption = "Status: Disconnecting...please wait."
Text7.text = "offline"
SaveText Text7, "c:\windows\" + Text6.text + ".txt"
Text4.text = "c:\windows\" + Text6.text + ".txt"
Text5.text = Text6.text + ".txt"
Command10.Enabled = True
       Command11.Enabled = False
       Command12.Enabled = False
       Command13.Enabled = False
       Command14.Enabled = False
       Command15.Enabled = False

    If hConnection <> 0 Then
        InternetCloseHandle hConnection
    End If
    hConnection = 0
    Call WriteToINI("username", "name", Text6.text, "c:\windows\buddylist.ini")
Call WriteToINI("buddy1", "buddy", Label9.Caption, "c:\windows\buddylist.ini")
Call WriteToINI("buddy2", "buddy", Label10.Caption, "c:\windows\buddylist.ini")
Call WriteToINI("buddy3", "buddy", Label11.Caption, "c:\windows\buddylist.ini")
Call WriteToINI("buddy4", "buddy", Label12.Caption, "c:\windows\buddylist.ini")
Call WriteToINI("buddy5", "buddy", Label13.Caption, "c:\windows\buddylist.ini")
Call WriteToINI("buddy6", "buddy", Label14.Caption, "c:\windows\buddylist.ini")
    End
End Sub

Private Sub Command19_Click()
Form1.WindowState = vbMinimized
End Sub

Private Sub Command2_Click()
Text5.text = LCase(Text10.text) + "chat.txt"
If (FtpDeleteFile(hConnection, Text5.text) = False) Then
    MsgBox "FtpDeleteFile error: " & Err.LastDllError
            Exit Sub
Else
     MsgBox "Chat was deleted from server....", 12, "Chat deleted"
End If
End Sub

Private Sub Command20_Click()
Form1.Width = "6285"
Command23.Visible = True
Command20.Visible = False
End Sub

Private Sub Command21_Click()
If Text8.text = "" Then
MsgBox "Please input a location first..", 12, "Please input a location/message"
Exit Sub
End If
If Text8.text = " " Then
MsgBox "Please input a location first..", 12, "Please input a location/message"
Exit Sub
End If
If Text8.text = "  " Then
MsgBox "Please input a location first..", 12, "Please input a location/message"
Exit Sub
End If
If Text8.text = "   " Then
MsgBox "Please input a location first..", 12, "Please input a location/message"
Exit Sub
End If
Form1.Caption = "Status: Defining location...please wait."
Form1.MousePointer = "11"
Text7.text = Text8.text
SaveText Text7, "c:\windows\" + LCase(Text6.text) + "loc.txt"
Text4.text = "c:\windows\" + LCase(Text6.text) + "loc.txt"
Text5.text = LCase(Text6.text) + "loc.txt"
 If (FtpPutFile(hConnection, Text4.text, Text5.text, _
         dwType, 0) = False) Then
            MsgBox "Unable to define your location. Please make sure you are still connected to the internet. Or try connecting to the server again.", 12, "Unable to Connect."
            Form1.Caption = "Status: Offline."
             Exit Sub
        Else
        Form1.MousePointer = "0"
        Timer3.Enabled = True
MsgBox "Successfully defined your location.", 12, "Finished."
Form1.Caption = "Status: " + Text6.text + " Logged in"
       Command10.Enabled = True
       Command11.Enabled = True
       Command12.Enabled = True
       Command13.Enabled = True
       Command14.Enabled = True
       Command15.Enabled = True
          Command9.Enabled = False
    butClose.Enabled = True
    Command17.Enabled = True
    Command18.Enabled = False
    End If
End Sub

Private Sub Command22_Click()
Command27.Visible = False
Command28.Visible = True
Form1.Height = "6360"
' for ASCII files use FTP_TRANSFER_TYPE_ASCII
Text4.text = "c:\windows\conceptmessageYAY.txt"
Text5.text = "conceptmessageYAY.txt"
' add INTERNET_FLAG_NO_CACHE_WRITE to avoid local caching 0x04000000 (hex)
 If (FtpGetFile(hConnection, Text5.text, Text4.text, False, _
         FILE_ATTRIBUTE_NORMAL, dwType Or INTERNET_FLAG_RELOAD, 0) = False) Then
    MsgBox "Sorry, the message of the day could not be downloaded. Please try again or reconnect.", vbCritical, "Error"
    Exit Sub
 Else
 
   End If
   
   LoadText Text7, "c:\windows\conceptmessageYAY.txt"
   Text11.text = Text7.text
End Sub

Private Sub Command23_Click()
Form1.Width = "3540"
Command23.Visible = False
Command20.Visible = True
End Sub

Private Sub Command24_Click()
Text9.text = Text6.text
Form1.Height = "6360"
Command27.Visible = False
Command28.Visible = True
End Sub

Private Sub Command25_Click()
Form1.Caption = "Status: Joining...please wait"
Form1.MousePointer = "11"
If Text9.text = "" Then
Exit Sub
End If
' for ASCII files use FTP_TRANSFER_TYPE_ASCII
Text4.text = "c:\windows\" + LCase(Text10.text) + "chat.txt"
Text5.text = LCase(Text10.text) + "chat.txt"

' add INTERNET_FLAG_NO_CACHE_WRITE to avoid local caching 0x04000000 (hex)
 If (FtpGetFile(hConnection, Text5.text, Text4.text, False, _
         FILE_ATTRIBUTE_NORMAL, dwType Or INTERNET_FLAG_RELOAD, 0) = False) Then
    MsgBox "Sorry, that chat could not be found.", 12, "No such chat room."
    Form1.Caption = "Status: " + Text6.text + " Logged in"
   Form1.MousePointer = "0"
    Exit Sub
 Else
 
   End If
   
   LoadText Text7, "c:\windows\" + LCase(Text10.text) + "chat.txt"
  Form1.Caption = "Status: " + Text6.text + " Logged in"
Form1.MousePointer = "0"
frmmain.Show
frmmain.txtName.text = Text9.text
frmmain.txtIP.text = Text7.text
frmmain.Timer2.Enabled = True
Command25.Enabled = False
Command26.Enabled = False
Form1.MousePointer = "0"
End Sub

Private Sub Command26_Click()

frmports.Show
End Sub

Private Sub Command27_Click()
Command27.Visible = False
Command28.Visible = True
Form1.Height = "6360"
End Sub

Private Sub Command28_Click()
Command28.Visible = False
Command27.Visible = True
Form1.Height = "4365"
End Sub

Private Sub Command29_Click()
frmmain.Show
frmmain.Label5.Caption = "hosting"
frmmain.Timer3.Enabled = True
TimeOut (0.2)
Text7.text = frmmain.Label6.Caption
frmmain.txtName.text = Text9.text
SaveText Text7, "c:\windows\" + LCase(Text10.text) + "chat.txt"
Text4.text = "c:\windows\" + LCase(Text10.text) + "chat.txt"
Text5.text = LCase(Text10.text) + "chat.txt"
 If (FtpPutFile(hConnection, Text4.text, Text5.text, _
         dwType, 0) = False) Then
            MsgBox "Unable to host. Make sure you are still connected to the internet. if so, try dissconnecting then reconnecting.", 12, "Unable to Host."
            Form1.Caption = "Status: Unable to Host."
             Form1.MousePointer = "0"
             Exit Sub
        Else
        Form1.MousePointer = "0"
        Timer3.Enabled = True
MsgBox "Sucessfully Hosting Chat: " + Text10.text + " !", 12, "Hosting..."
Form1.Caption = "Status: " + Text6.text + " Logged in"
       Command10.Enabled = True
       Command11.Enabled = True
       Command12.Enabled = True
       Command13.Enabled = True
       Command14.Enabled = True
       Command15.Enabled = True
          Command9.Enabled = False
    butClose.Enabled = True
    Command17.Enabled = True
    Command18.Enabled = False
    frmmain.Caption = "Chat Room: " + Text10.text
    Command25.Enabled = False
Command26.Enabled = False
Form1.MousePointer = "0"
    End If
End Sub

Private Sub Command3_Click()
    If hConnection <> 0 Then
        InternetCloseHandle hConnection
    End If
    hConnection = InternetConnect(hOpen, Text1.text, INTERNET_INVALID_PORT_NUMBER, _
    Text2.text, Text3.text, INTERNET_SERVICE_FTP, dwSeman, 0)
    If hConnection = 0 Then
        ErrorOut Err.LastDllError, "InternetConnect"
        Exit Sub
    Else

        Option3.Enabled = False
        Option4.Enabled = False
    End If
        

End Sub

Private Sub Command30_Click()
Form1.Caption = "Status: Joining...please wait"
Form1.MousePointer = "11"
If Text9.text = "" Then
Exit Sub
End If
' for ASCII files use FTP_TRANSFER_TYPE_ASCII
Text4.text = "c:\windows\" + LCase(Text10.text) + "chat.txt"
Text5.text = LCase(Text10.text) + "chat.txt"

' add INTERNET_FLAG_NO_CACHE_WRITE to avoid local caching 0x04000000 (hex)
 If (FtpGetFile(hConnection, Text5.text, Text4.text, False, _
         FILE_ATTRIBUTE_NORMAL, dwType Or INTERNET_FLAG_RELOAD, 0) = False) Then
    MsgBox "Sorry, that chat could not be found.", 12, "No such chat room."
   Form1.MousePointer = "0"
    Exit Sub
 Else
 
   End If
   
   LoadText Text7, "c:\windows\" + LCase(Text10.text) + "chat.txt"
  Form1.Caption = "Status: " + Text6.text + " Logged in"
Form1.MousePointer = "0"
frmmain.Show
frmmain.Label5.Caption = "joining"
frmmain.txtName.text = Text9.text
frmmain.txtIP.text = Text7.text
frmmain.Timer2.Enabled = True
Command25.Enabled = False
Command26.Enabled = False
Form1.MousePointer = "0"
End Sub

Private Sub Command4_Click()
'&H40000000 == GENERIC_WRITE
Dim Data(99) As Byte ' array of 100 elements 0 to 99
Dim Written As Long
Dim size As Long
Dim Sum As Long
Dim j As Long

Sum = 0
j = 0
' for ASCII files use FTP_TRANSFER_TYPE_ASCII
hFile = FtpOpenFile(hConnection, Text5.text, &H40000000, dwType, 0)
If hFile = 0 Then
    ErrorOut Err.LastDllError, "FtpOpenFile"
    Exit Sub
End If
Open Text4.text For Binary Access Read As #1
size = LOF(1)
For j = 1 To size \ 100
    Get #1, , Data
    If (InternetWriteFile(hFile, Data(0), 100, Written) = 0) Then
        ErrorOut Err.LastDllError, "InternetWriteFile"
        Exit Sub
    End If
    DoEvents
    Sum = Sum + 100
    Label7.Caption = Str(Sum)
Next j
Get #1, , Data
 If (InternetWriteFile(hFile, Data(0), size Mod 100, Written) = 0) Then
        ErrorOut Err.LastDllError, "InternetWriteFile"
        Exit Sub
End If
Sum = Sum + (size Mod 100)
Label7.Caption = Str(Sum)
Close #1
InternetCloseHandle (hFile)
End Sub

Private Sub Command5_Click()
' for ASCII files use FTP_TRANSFER_TYPE_ASCII
' add INTERNET_FLAG_NO_CACHE_WRITE to avoid local caching 0x04000000 (hex)
 If (FtpGetFile(hConnection, Text5.text, Text4.text, False, _
         FILE_ATTRIBUTE_NORMAL, dwType Or INTERNET_FLAG_RELOAD, 0) = False) Then
    ErrorOut Err.LastDllError, "FtpPutFile"
    Exit Sub
 Else
   MsgBox "File transfered!", , "SimpleFtp"
   End If
End Sub

Private Sub Command6_Click()
Dim szDir As String

szDir = String(1024, Chr$(0))

If (FtpGetCurrentDirectory(hConnection, szDir, 1024) = False) Then
    ErrorOut Err.LastDllError, "FtpGetCurrentDirectory"
    Exit Sub
 Else
   MsgBox "Current directory is: " & szDir, , "SimpleFtp"
   End If
End Sub

Private Sub Command7_Click()
If (FtpSetCurrentDirectory(hConnection, Text5.text) = False) Then
   ErrorOut Err.LastDllError, "FtpSetCurrentDirectory"
   Exit Sub
Else
  MsgBox "Directory is changed to " & Text5.text, , "SimpleFtp"
End If

End Sub

Private Sub Command8_Click()
Dim szDir As String
Dim hFind As Long
Dim nLastError As Long
Dim dError As Long
Dim ptr As Long
Dim pData As WIN32_FIND_DATA
    

hFind = FtpFindFirstFile(hConnection, "*.*", pData, 0, 0)
nLastError = Err.LastDllError
If hFind = 0 Then
        If (nLastError = ERROR_NO_MORE_FILES) Then
            MsgBox "This directory is empty!", , "SimpleFtp"
        Else
            ErrorOut Err.LastDllError, "FtpFindFirstFile"
        End If
        Exit Sub
End If

dError = NO_ERROR
     Dim bRet As Boolean
 
szDir = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1) & " " & Win32ToVbTime(pData.ftLastWriteTime)
szDir = szDir & vbCrLf
Do
        pData.cFileName = String(MAX_PATH, 0)
        bRet = InternetFindNextFile(hFind, pData)
        If Not bRet Then
            dError = Err.LastDllError
            If dError = ERROR_NO_MORE_FILES Then
                Exit Do
            Else
                ErrorOut Err.LastDllError, "InternetFindNextFile"
                InternetCloseHandle (hFind)
                Exit Sub
            End If
        Else
            
            szDir = szDir & Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1) & " " & Win32ToVbTime(pData.ftLastWriteTime) & vbCrLf
        End If
Loop
   
Dim szTemp As String
szTemp = String(1024, Chr$(0))
If (FtpGetCurrentDirectory(hConnection, szTemp, 1024) = False) Then
    ErrorOut Err.LastDllError, "FtpGetCurrentDirectory"
    Exit Sub
End If
MsgBox szDir, , "Directory Listing of: " & szTemp
InternetCloseHandle (hFind)
End Sub

Private Sub Command9_Click()
If Text6.text = "" Then
MsgBox "Please input a name before connecting", 12, "Please input a name."
Exit Sub
End If
If Text6.text = " " Then
MsgBox "Please input a name before connection.", 12, "Please input a name."
Exit Sub
End If
If Text6.text = "  " Then
MsgBox "Please input a name before connection.", 12, "Please input a name."
Exit Sub
End If
If Text6.text = "   " Then
MsgBox "Please input a name before connection.", 12, "Please input a name."
Exit Sub
End If
Text6.text = LCase(Text6.text)
Form1.Caption = "Status: Connecting...please wait."
Form1.MousePointer = "11"
Call Command3_Click
Text7.text = "online"
SaveText Text7, "c:\windows\" + LCase(Text6.text) + ".txt"
Text4.text = "c:\windows\" + LCase(Text6.text) + ".txt"
Text5.text = LCase(Text6.text) + ".txt"
 If (FtpPutFile(hConnection, Text4.text, Text5.text, _
         dwType, 0) = False) Then
            MsgBox "Unable to connect. Please make sure you are connected to the internet.", 12, "Unable to Connect."
            Form1.Caption = "Status: Offline."
             Form1.MousePointer = "0"
             Exit Sub
        Else
        Form1.MousePointer = "0"
        Timer3.Enabled = True
MsgBox "Connected to Buddy List Server.", 12, "Connection Complete."
Form1.Caption = "Status: " + Text6.text + " Logged in"
       Command10.Enabled = True
       Command11.Enabled = True
       Command12.Enabled = True
       Command13.Enabled = True
       Command14.Enabled = True
       Command15.Enabled = True
          Command9.Enabled = False
    butClose.Enabled = True
    Command17.Enabled = True
    Command18.Enabled = False
    Command21.Enabled = True
    Command22.Enabled = True
    Command24.Enabled = True
    Command27.Enabled = True
    End If
 
End Sub

Private Sub define1_Click()
Dim message, Title, Default, MyValue
message = "Please enter your buddy's name:"   ' Set prompt.
Title = "Name?" ' Set title.
Default = ""   ' Set default.
' Display message, title, and default value.
MyValue = InputBox(message, Title, Default)

Label9.Caption = MyValue
End Sub

Private Sub dt_Click()
Dim message, Title, Default, MyValue
message = "Please enter your buddy's name:"   ' Set prompt.
Title = "Name?" ' Set title.
Default = ""   ' Set default.
' Display message, title, and default value.
MyValue = InputBox(message, Title, Default)

Label10.Caption = MyValue
End Sub

Private Sub Form_Load()
Call GetWindowRect(FindWindowEx(0&, 0&, "Shell_TrayWnd", vbNullString), tRect)
On Error GoTo Hell
Text6.text = GetFromINI("username", "name", "c:\windows\buddylist.ini")
  hOpen = InternetOpen("My VB Test", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
  If hOpen = 0 Then
    ErrorOut Err.LastDllError, "InternetOpen"
    Unload Form1
  End If
  dwType = FTP_TRANSFER_TYPE_ASCII
  dwSeman = 0
  hConnection = 0
    Label9.Caption = "none"
  Label10.Caption = "none"
Label11.Caption = "none"
Label12.Caption = "none"
Label13.Caption = "none"
Label14.Caption = "none"
  Label9.Caption = GetFromINI("buddy1", "buddy", "c:\windows\buddylist.ini")
  Label10.Caption = GetFromINI("buddy2", "buddy", "c:\windows\buddylist.ini")
  Label11.Caption = GetFromINI("buddy3", "buddy", "c:\windows\buddylist.ini")
  Label12.Caption = GetFromINI("buddy4", "buddy", "c:\windows\buddylist.ini")
  Label13.Caption = GetFromINI("buddy5", "buddy", "c:\windows\buddylist.ini")
  Label14.Caption = GetFromINI("buddy6", "buddy", "c:\windows\buddylist.ini")

Hell:
Exit Sub
  
     
End Sub
Private Sub Form_Uload()
    InternetCloseHandle hOpen
Unload Form2
End Sub

Private Sub Form_Unload(Cancel As Integer)
Text7.text = "offline"
SaveText Text7, "c:\windows\" + Text6.text + ".txt"
Text4.text = "c:\windows\" + Text6.text + ".txt"
Text5.text = Text6.text + ".txt"
 If (FtpPutFile(hConnection, Text4.text, Text5.text, _
         dwType, 0) = False) Then
             ErrorOut Err.LastDllError, "FtpPutFile"
             Exit Sub
        Else
        Form1.Caption = "Status: offline"
Call WriteToINI("username", "name", Text6.text, "c:\windows\buddylist.ini")
MsgBox "Disconnected from AW Buddy Server.", 12, "Disconnected"
        End If
End Sub

Private Sub Label10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu buddy2, 0
End If
End Sub

Private Sub Label11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu buddy3, 0
End If
End Sub

Private Sub Label12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu buddy4, 0
End If
End Sub

Private Sub Label13_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu buddy5, 0
End If
End Sub

Private Sub Label14_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu buddy6, 0
End If
End Sub

Private Sub Label9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu buddy1, 0
End If
End Sub

Private Sub Option1_Click()
dwType = FTP_TRANSFER_TYPE_ASCII
End Sub

Private Sub Option2_Click()
dwType = FTP_TRANSFER_TYPE_BINARY
End Sub

Private Sub Option3_Click()
dwSeman = 0
End Sub

Private Sub Option4_Click()
dwSeman = INTERNET_FLAG_PASSIVE
End Sub

Private Sub Timer1_Timer()
If Label15.Caption = "online" Then Label15.forecolor = vbGreen
If Label16.Caption = "online" Then Label16.forecolor = vbGreen
If Label17.Caption = "online" Then Label17.forecolor = vbGreen
If Label18.Caption = "online" Then Label18.forecolor = vbGreen
If Label19.Caption = "online" Then Label19.forecolor = vbGreen
If Label20.Caption = "online" Then Label20.forecolor = vbGreen
If Label15.Caption = "offline" Then Label15.forecolor = vbRed
If Label16.Caption = "offline" Then Label16.forecolor = vbRed
If Label17.Caption = "offline" Then Label17.forecolor = vbRed
If Label18.Caption = "offline" Then Label18.forecolor = vbRed
If Label19.Caption = "offline" Then Label19.forecolor = vbRed
If Label20.Caption = "offline" Then Label20.forecolor = vbRed
If Label15.Caption = "  n/a" Then Label15.forecolor = vbRed
If Label16.Caption = "  n/a" Then Label16.forecolor = vbRed
If Label17.Caption = "  n/a" Then Label17.forecolor = vbRed
If Label18.Caption = "  n/a" Then Label18.forecolor = vbRed
If Label19.Caption = "  n/a" Then Label19.forecolor = vbRed
If Label20.Caption = "  n/a" Then Label20.forecolor = vbRed

End Sub

Private Sub Timer2_Timer()
If Label9.Caption = "" Then Label9.Caption = "none"
If Label10.Caption = "" Then Label10.Caption = "none"
If Label11.Caption = "" Then Label11.Caption = "none"
If Label12.Caption = "" Then Label12.Caption = "none"
If Label13.Caption = "" Then Label13.Caption = "none"
If Label14.Caption = "" Then Label14.Caption = "none"

End Sub

Private Sub Timer3_Timer()
Shape2.Visible = True
TimeOut (0.2)
Shape2.Visible = False
Shape3.Visible = True
TimeOut (0.2)
Shape3.Visible = False
Shape4.Visible = True
TimeOut (0.2)
Shape4.Visible = False
Shape5.Visible = True
TimeOut (0.2)
Shape5.Visible = False
Shape6.Visible = True
TimeOut (0.2)
Shape6.Visible = False
Shape5.Visible = True
TimeOut (0.2)
Shape5.Visible = False
Shape4.Visible = True
TimeOut (0.2)
Shape4.Visible = False
Shape3.Visible = True
TimeOut (0.2)
Shape3.Visible = False
Shape2.Visible = True

End Sub

Private Sub Timer4_Timer()
DoEvents

If (Left >= -ScaleX(Val(Text12.text), vbPixels, vbTwips) And Left <= ScaleX(Val(Text12.text), vbPixels, vbTwips)) And (Top >= -ScaleY(Val(Text1.text), vbPixels, vbTwips) And Top <= ScaleY(Val(Text1.text), vbPixels, vbTwips)) Then
'Topleft snap
    Top = 0
    Left = 0

ElseIf (Top + Height <= ScaleY(tRect.Top, vbPixels, vbTwips) + ScaleY(Val(Text12.text), vbPixels, vbTwips) And Top + Height >= ScaleY(tRect.Top, vbPixels, vbTwips) - ScaleY(Val(Text1.text), vbPixels, vbTwips)) And (Left >= -ScaleX(Val(Text1.text), vbPixels, vbTwips) And Left <= ScaleX(Val(Text1.text), vbPixels, vbTwips)) Then
'Bottomleft snap
    Top = ScaleY(tRect.Top, vbPixels, vbTwips) - Height
    Left = 0

ElseIf (Top + Height <= ScaleY(tRect.Top, vbPixels, vbTwips) + ScaleY(Val(Text12.text), vbPixels, vbTwips) And Top + Height >= ScaleY(tRect.Top, vbPixels, vbTwips) - ScaleY(Val(Text12.text), vbPixels, vbTwips)) And (Left + Width <= Screen.Width + ScaleX(Val(Text12.text), vbPixels, vbTwips) And Left + Width >= Screen.Width - ScaleX(Val(Text12.text), vbPixels, vbTwips)) Then
'Bottomright snap
    Top = ScaleY(tRect.Top, vbPixels, vbTwips) - Height
    Left = Screen.Width - Width

ElseIf (Top >= -ScaleY(Val(Text1.text), vbPixels, vbTwips) And Top <= ScaleY(Val(Text12.text), vbPixels, vbTwips)) And (Left + Width <= Screen.Width + ScaleX(Val(Text12.text), vbPixels, vbTwips) And Left + Width >= Screen.Width - ScaleX(Val(Text12.text), vbPixels, vbTwips)) Then
'Topright snap
    Top = 0
    Left = Screen.Width - Width

ElseIf Top >= -ScaleY(Val(Text12.text), vbPixels, vbTwips) And Top <= ScaleY(Val(Text12.text), vbPixels, vbTwips) Then
'Top snap
    Top = 0

ElseIf Left >= -ScaleX(Val(Text12.text), vbPixels, vbTwips) And Left <= ScaleX(Val(Text12.text), vbPixels, vbTwips) Then
'Left snap
    Left = 0

ElseIf Top + Height <= ScaleY(tRect.Top, vbPixels, vbTwips) + ScaleY(Val(Text12.text), vbPixels, vbTwips) And Top + Height >= ScaleY(tRect.Top, vbPixels, vbTwips) - ScaleY(Val(Text12.text), vbPixels, vbTwips) Then
'Bottom snap
    Top = ScaleY(tRect.Top, vbPixels, vbTwips) - Height

ElseIf Left + Width <= Screen.Width + ScaleX(Val(Text12.text), vbPixels, vbTwips) And Left + Width >= Screen.Width - ScaleX(Val(Text12.text), vbPixels, vbTwips) Then
'Right snap
    Left = Screen.Width - Width

End If
End Sub

Private Sub Timer5_Timer()
Call Command29_Click
Timer5.Enabled = False
End Sub

Private Sub Timer6_Timer()
Text5.text = LCase(Text10.text) + "chat.txt"
If (FtpDeleteFile(hConnection, Text5.text) = False) Then
    MsgBox "Coulld not delete chat from server....", vbCritical, "Error.."
            Exit Sub
Else
     MsgBox "Chat was deleted from server....", 12, "Chat deleted"
     Timer6.Enabled = False
End If


End Sub

Private Sub Timer7_Timer()
Call Command30_Click
Timer7.Enabled = False
End Sub

Private Sub weee_Click()
If Label14.Caption = "none" Then
Exit Sub
End If
' for ASCII files use FTP_TRANSFER_TYPE_ASCII
Text4.text = "c:\windows\" + LCase(Label14.Caption) + "loc.txt"
Text5.text = LCase(Label14.Caption) + "loc.txt"
' add INTERNET_FLAG_NO_CACHE_WRITE to avoid local caching 0x04000000 (hex)
 If (FtpGetFile(hConnection, Text5.text, Text4.text, False, _
         FILE_ATTRIBUTE_NORMAL, dwType Or INTERNET_FLAG_RELOAD, 0) = False) Then
    MsgBox "Sorry, this user has not specfied a location/message..", 12, "No location/message"
     Label29.Caption = "  n/a"
    Exit Sub
 Else
 
   End If
   
   LoadText Text7, "c:\windows\" + LCase(Label14.Caption) + "loc.txt"
   Label29.Caption = Text7.text
End Sub

Private Sub werasdf_Click()
Dim message, Title, Default, MyValue
message = "Please enter your buddy's name:"   ' Set prompt.
Title = "Name?" ' Set title.
Default = ""   ' Set default.
' Display message, title, and default value.
MyValue = InputBox(message, Title, Default)

Label13.Caption = MyValue
End Sub

Private Sub wetaw_Click()
Dim message, Title, Default, MyValue
message = "Please enter your buddy's name:"   ' Set prompt.
Title = "Name?" ' Set title.
Default = ""   ' Set default.
' Display message, title, and default value.
MyValue = InputBox(message, Title, Default)

Label12.Caption = MyValue
End Sub
