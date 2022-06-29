VERSION 5.00
Begin VB.Form frmABOUT 
   BorderStyle     =   0  'None
   Caption         =   "About"
   ClientHeight    =   4725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraABOUT 
      Height          =   4665
      Left            =   70
      TabIndex        =   5
      Top             =   0
      Width           =   4215
      Begin VB.CommandButton cmdLISTFILES 
         Caption         =   "List of files"
         Height          =   285
         Left            =   3000
         TabIndex        =   11
         Top             =   3850
         Width           =   1095
      End
      Begin VB.CommandButton cmdCLOSE 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   285
         Left            =   3015
         TabIndex        =   7
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   $"frmABOUT.frx":0000
         Height          =   615
         Left            =   300
         TabIndex        =   10
         Top             =   2250
         Width           =   3615
      End
      Begin VB.Label Label2 
         Caption         =   "Please send any comments you have to this email address:"
         Height          =   465
         Left            =   300
         TabIndex        =   9
         Top             =   3000
         Width           =   3165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(Use the arrow keys for controls)"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   270
         TabIndex        =   8
         Top             =   4200
         Width           =   2280
      End
      Begin VB.Label urlHOMEPAGE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " http://come.to/magikcube"
         Height          =   195
         Left            =   225
         MouseIcon       =   "frmABOUT.frx":0091
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label urlEMAIL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " neomatrix@magikcube.freeserve.co.uk"
         Height          =   195
         Left            =   225
         TabIndex        =   3
         Top             =   3525
         Width           =   2820
      End
      Begin VB.Label lblABOUT2 
         Caption         =   $"frmABOUT.frx":2833
         Height          =   885
         Left            =   240
         TabIndex        =   2
         Top             =   1350
         Width           =   3855
      End
      Begin VB.Label lblABOUT1 
         Caption         =   "This program was brought to you by one of the best Visual Basic programmers around (maybe not)."
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label lblTITLEABOUT2 
         BackStyle       =   0  'Transparent
         Caption         =   "BOUT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   600
         TabIndex        =   0
         Top             =   465
         Width           =   735
      End
      Begin VB.Label lblTITLEABOUT1 
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmABOUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long  'Used for URL Coding

Private Sub cmdCLOSE_Click()
Unload Me
End Sub

Private Sub cmdLISTFILES_Click()
frmLISTFILES.Show
End Sub

Private Sub Form_Load()
frmMAIN.Show
End Sub

'Private Sub urlEMAIL_Click()
'ShellExecute 0, "Open", "mailto:neomatrix@magikcube.freeserve.co.uk?Subject=One of your programs", "", "", 0
'End Sub

Private Sub urlHOMEPAGE_Click()
ShellExecute 0, "Open", "http://come.to/magikcube", "", "", 0
End Sub
