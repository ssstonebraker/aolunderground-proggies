VERSION 5.00
Object = "{683364A1-B37D-11D1-ADC5-006008A5848C}#1.0#0"; "DHTMLED.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "html preview example by vsx"
   ClientHeight    =   2550
   ClientLeft      =   5820
   ClientTop       =   2310
   ClientWidth     =   3300
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   Begin DHTMLEDLibCtl.DHTMLEdit DHTMLEdit1 
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   3015
      ActivateApplets =   0   'False
      ActivateActiveXControls=   0   'False
      ActivateDTCs    =   -1  'True
      ShowDetails     =   0   'False
      ShowBorders     =   0   'False
      Appearance      =   1
      Scrollbars      =   -1  'True
      ScrollbarAppearance=   1
      SourceCodePreservation=   -1  'True
      AbsoluteDropMode=   0   'False
      SnapToGrid      =   0   'False
      SnapToGridX     =   50
      SnapToGridY     =   50
      BrowseMode      =   0   'False
      UseDivOnCarriageReturn=   0   'False
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C2811F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x"
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
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   10
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C2811F&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "html preview example by vsx"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   20
      TabIndex        =   0
      Top             =   20
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
form3d Me

End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
moveform Me
End Sub


Private Sub Label2_Click()
Unload Me
End

End Sub


Private Sub Text1_Change()
On Error GoTo err
DHTMLEdit1.DocumentHTML = Text1.Text
err:
End Sub


