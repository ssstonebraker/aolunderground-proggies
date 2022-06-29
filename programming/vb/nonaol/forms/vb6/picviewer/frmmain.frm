VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Picture Viewer Example"
   ClientHeight    =   885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   2790
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2280
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Picture"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "example by: NiRVaNa"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CommonDialog1.DefaultExt = "(*.JPG)"
CommonDialog1.DialogTitle = "Select Picture"
CommonDialog1.Filter = "(*.JPG)"
CommonDialog1.ShowOpen
If x = vbCancel Then
Exit Sub
Else
Set f = New frmpicture
f.Image1.Picture = VB.LoadPicture(CommonDialog1.FileName)
f.Height = f.Image1.Height
f.Width = f.Image1.Width
f.Caption = "File - [" + CommonDialog1.FileName + "]"
f.Show
End If
End Sub
