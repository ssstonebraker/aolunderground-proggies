VERSION 5.00
Object = "{A42B3179-DC51-11D0-BD85-7C6903C10627}#1.0#0"; "FlatBar32.ocx"
Begin VB.Form FormX 
   Caption         =   "FlatBar32 - IE 3.0 and above is required."
   ClientHeight    =   5385
   ClientLeft      =   3645
   ClientTop       =   2805
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   8835
   Begin FlatBar32.FlatBar FlatBar4 
      Height          =   795
      Left            =   90
      TabIndex        =   10
      Top             =   1890
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   1402
      ToolBarSize     =   32
      ButnCount       =   2
      Caption0        =   "New"
      ToolTips0       =   "New"
      BtnStyle0       =   1
      Caption1        =   "Open"
      ImageNumber1    =   1
      ToolTips1       =   "Open"
      BtnStyle1       =   1
      Caption2        =   "Save"
      ImageNumber2    =   2
      ToolTips2       =   "Save"
      BtnStyle2       =   1
      WithText        =   -1  'True
      Picture         =   "Test.frx":0000
      MaskColor       =   16776960
   End
   Begin FlatBar32.FlatBar FlatBar3 
      Height          =   345
      Left            =   75
      TabIndex        =   9
      Top             =   1335
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   609
      ButnCount       =   2
      Caption0        =   "Microsoft"
      ToolTips0       =   "Microsoft's Home Page"
      BtnStyle0       =   1
      Caption1        =   "Visual Basic"
      ToolTips1       =   "Visual Basic Home Page"
      BtnStyle1       =   1
      Caption2        =   "IconMenu32 DLL"
      ToolTips2       =   "Download IconMenu32 DLL to create menus with Icons or Bitmaps."
      BtnStyle2       =   1
      WithText        =   -1  'True
      Style           =   -1  'True
      Picture         =   "Test.frx":1647
      MaskColor       =   16776960
   End
   Begin FlatBar32.FlatBar FlatBar2 
      Height          =   570
      Left            =   120
      TabIndex        =   8
      Top             =   615
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   1005
      ToolBarSize     =   32
      ButnCount       =   3
      CheckGroup      =   -1  'True
      ImageNumber0    =   24
      ToolTips0       =   "Large Icons"
      BtnStyle0       =   1
      CheckGroup0     =   -1  'True
      ImageNumber1    =   25
      ToolTips1       =   "Small Icons"
      BtnStyle1       =   1
      CheckGroup1     =   -1  'True
      ImageNumber2    =   23
      ToolTips2       =   "List"
      BtnStyle2       =   1
      CheckGroup2     =   -1  'True
      ImageNumber3    =   22
      ToolTips3       =   "Details"
      BtnStyle3       =   1
      CheckGroup3     =   -1  'True
      Picture         =   "Test.frx":1A01
      MaskColor       =   16776960
   End
   Begin FlatBar32.FlatBar FlatBar1 
      Height          =   345
      Left            =   195
      TabIndex        =   7
      Top             =   120
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   609
      ButnCount       =   15
      ToolTips0       =   "New"
      BtnStyle0       =   1
      ImageNumber1    =   1
      ToolTips1       =   "Open"
      BtnStyle1       =   1
      ImageNumber2    =   2
      ToolTips2       =   "Save"
      BtnStyle2       =   1
      ImageNumber3    =   3
      ToolTips3       =   "Cut"
      BtnStyle3       =   1
      ImageNumber4    =   4
      ToolTips4       =   "Copy"
      BtnStyle4       =   1
      ImageNumber5    =   5
      ToolTips5       =   "Paste"
      BtnStyle5       =   1
      ImageNumber6    =   6
      ToolTips6       =   "Print"
      BtnStyle6       =   1
      ImageNumber7    =   7
      ToolTips7       =   "Help"
      BtnStyle7       =   1
      ImageNumber8    =   8
      ToolTips8       =   "What's This"
      BtnStyle8       =   1
      ImageNumber9    =   9
      BtnStyle9       =   1
      ImageNumber10   =   10
      BtnStyle10      =   1
      ImageNumber11   =   11
      BtnStyle11      =   1
      ImageNumber12   =   12
      BtnStyle12      =   1
      ImageNumber13   =   13
      BtnStyle13      =   1
      ImageNumber14   =   14
      BtnStyle14      =   1
      ImageNumber15   =   15
      BtnStyle15      =   1
      Picture         =   "Test.frx":3048
      MaskColor       =   16776960
      Wrappable       =   -1  'True
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   420
      Left            =   -990
      TabIndex        =   0
      Top             =   3945
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enable Button 1"
      Height          =   375
      Index           =   1
      Left            =   1845
      TabIndex        =   4
      Top             =   3300
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Disable Button 1"
      Height          =   375
      Index           =   0
      Left            =   255
      TabIndex        =   1
      Top             =   3300
      Width           =   1515
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ZoneCorp@Compuserve.com"
      Height          =   195
      Left            =   6570
      MouseIcon       =   "Test.frx":475A
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   5115
      Width           =   2115
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ZoneCorp@AOL.com"
      Height          =   195
      Left            =   3840
      MouseIcon       =   "Test.frx":4A64
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   5115
      Width           =   1530
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Ramon Guerrero ZoneCorp@dallas.net"
      Height          =   195
      Left            =   210
      MouseIcon       =   "Test.frx":4D6E
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   5115
      Width           =   2820
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"Test.frx":5078
      Height          =   1185
      Left            =   255
      TabIndex        =   2
      Top             =   3810
      Width           =   8340
   End
   Begin VB.Image RebarImage 
      Appearance      =   0  'Flat
      Height          =   1875
      Left            =   7065
      Picture         =   "Test.frx":52F7
      Top             =   915
      Visible         =   0   'False
      Width           =   9600
   End
End
Attribute VB_Name = "FormX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

 
 

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
FlatBar1.DisAbleButton 1
FlatBar2.DisAbleButton 1
FlatBar3.DisAbleButton 1
FlatBar4.DisAbleButton 1

Case 1
FlatBar1.EnableButton 1
FlatBar2.EnableButton 1
FlatBar3.EnableButton 1
FlatBar4.EnableButton 1
 
End Select
End Sub

Private Sub FlatBar1_Click(id As Integer)
MsgBox "Toolbar 1 Button " & id
End Sub

Private Sub FlatBar3_Click(id As Integer)
Select Case id
Case 0
FlatBar3.RunFile "http://www.microsoft.com"
Case 1
FlatBar3.RunFile "http://www.microsoft.com/vbasic"
Case 2
FlatBar3.RunFile "http://www.ZoneCorp.com/IconMenu32/IconMenu32.exe"
End Select
End Sub

Private Sub FlatBar4_Click(id As Integer)
MsgBox id
End Sub

Private Sub Form_Load()
On Error Resume Next
'Must initialize for RunTime
 FlatBar1.RunTime
 FlatBar2.RunTime
'FlatBar2's buttons are of the Check Group Style.
'Let's start off with the first button, (zero based), checked.
 FlatBar2.PressButton 0
 
 FlatBar3.RunTime
 FlatBar4.RunTime
'Add a separator bar in code
 FlatBar4.AddSeparator 3, 3
'Add a button in code
 FlatBar4.AddButtons 3, "Cut", 3, "Cut"
 
 
 'Set Rebar BackGround Picture
 FlatBar1.SetRebarPicture RebarImage

 'Create Rebar
 FlatBar1.CreateRebar

 'Add Bands to Rebar
 FlatBar1.AddBandsToRebar FlatBar1.GetToolbarHwnd, "Wrappable FlatBar"
 FlatBar1.AddBandsToRebar FlatBar2.GetToolbarHwnd, "Check Group Style"
 FlatBar1.AddBandsToRebar FlatBar3.GetToolbarHwnd, "Links"
 FlatBar1.AddBandsToRebar FlatBar4.GetToolbarHwnd, "Large Buttons With Text"
 
 'Don't need Design Time Containers since we are adding to the Rebar.
 'If your not then disregard the next four lines
 FlatBar1.Visible = False
 FlatBar2.Visible = False
 FlatBar3.Visible = False
 FlatBar4.Visible = False
 DoEvents
 Me.Show
 DoEvents
 DoEvents
 
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label2.ForeColor = vbBlue Then Label2.ForeColor = vbBlack
If Label3.ForeColor = vbBlue Then Label3.ForeColor = vbBlack
If Label4.ForeColor = vbBlue Then Label4.ForeColor = vbBlack
End Sub

Private Sub Form_Resize()
FlatBar1.ResizeRebar Me
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label2.ForeColor = vbBlue Then Label2.ForeColor = vbBlack
If Label3.ForeColor = vbBlue Then Label3.ForeColor = vbBlack
If Label4.ForeColor = vbBlue Then Label4.ForeColor = vbBlack
End Sub


Private Sub Label2_Click()
FlatBar1.RunFile "MailTo:ZoneCorp@dallas.net"


End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = vbBlue
End Sub


Private Sub Label3_Click()
FlatBar1.RunFile "MailTo:ZoneCorp@AOL.com"
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = vbBlue
End Sub


Private Sub Label4_Click()
FlatBar1.RunFile "MailTo:ZoneCorp@Compuserve.com"
End Sub


Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = vbBlue
End Sub


