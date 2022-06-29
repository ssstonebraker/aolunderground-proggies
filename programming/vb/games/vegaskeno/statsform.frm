VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form statsform 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Keno Ticket"
   ClientHeight    =   5025
   ClientLeft      =   4800
   ClientTop       =   4215
   ClientWidth     =   5325
   Icon            =   "statsform.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   3000
      TabIndex        =   10
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Print"
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Print Stats Form"
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Frame statsframe 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Keno Stats"
      ForeColor       =   &H000000FF&
      Height          =   2670
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   5025
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Bet this game.="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   195
         TabIndex        =   8
         Top             =   825
         Width           =   3615
      End
      Begin VB.Label bet_total 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   3885
         TabIndex        =   7
         Top             =   840
         Width           =   735
      End
      Begin VB.Label payrate 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   3120
         TabIndex        =   6
         Top             =   330
         Width           =   135
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Your Return"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   435
         TabIndex        =   5
         Top             =   315
         Width           =   1980
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Deals.="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   -30
         TabIndex        =   4
         Top             =   1995
         Width           =   3840
      End
      Begin VB.Label deals_label 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   3885
         TabIndex        =   3
         Top             =   2010
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Credits won this game.="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   -30
         TabIndex        =   2
         Top             =   1410
         Width           =   3840
      End
      Begin VB.Label credits_won 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   3885
         TabIndex        =   1
         Top             =   1425
         Width           =   735
      End
   End
   Begin VB.Label namelabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2280
      TabIndex        =   12
      Top             =   285
      Width           =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Players Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   240
      Width           =   2055
   End
   Begin VB.Menu newplayer 
      Caption         =   "New Player"
   End
End
Attribute VB_Name = "statsform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim BeginPage, EndPage, NumCopies, i
    ' Set Cancel to True
  CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Display the Print dialog box
   CommonDialog1.ShowPrinter
    ' Get user-selected values from the dialog box
     NumCopies = CommonDialog1.Copies
     For i = 1 To NumCopies
   statsform.PrintForm
    Next i
    Exit Sub
ErrHandler:
    ' User pressed the Cancel button
    Exit Sub
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub form_load()
playersname = GetSetting(App.Title, "Options", "playersname", "")
If playersname = "" Then
playersname = InputBox("Please enter Your Name.")
namelabel.Caption = playersname
SaveSetting App.Title, "options", "playersname", playersname
Else
namelabel.Caption = playersname
End If

bet_total.Caption = bettotal
credits_won.Caption = hopperempty
deals_label.Caption = totaldeals
Call Options.returnrate
payrate.Caption = percentagerate
End Sub

Private Sub newplayer_Click()
playersname = InputBox("Please enter Your Name.")
namelabel.Caption = playersname
SaveSetting App.Title, "options", "playersname", playersname
End Sub
