VERSION 5.00
Object = "{2F0E6D73-498C-11D3-9D69-0008C701AAC8}#1.0#0"; "VIRIISCAN.OCX"
Begin VB.Form Form1 
   Caption         =   "Virii Scan Example Using ViriiScan.ocx"
   ClientHeight    =   2055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Type Of Virii To Look For"
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   5055
      Begin VB.OptionButton Option6 
         Caption         =   "WFBD"
         Height          =   255
         Left            =   3960
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option5 
         Caption         =   "PWS"
         Height          =   255
         Left            =   2880
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option4 
         Caption         =   "VCL"
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Deltree"
         Height          =   195
         Left            =   480
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "If Detected..."
      Height          =   1095
      Left            =   2880
      TabIndex        =   4
      Top             =   0
      Width           =   2295
      Begin VB.OptionButton Option2 
         Caption         =   "Notify Me But Do Not Delete"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Delete File"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Scan File"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin Eroki_Virii_Scan.VScan VScan1 
      Left            =   1920
      Top             =   3360
      _ExtentX        =   6773
      _ExtentY        =   5080
   End
   Begin VB.Label Label1 
      Caption         =   "File To Scan:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1 = "" Then
MsgBox "Please enter a filename to scan!", vbCritical, "Error"
Exit Sub
End If
VScan1.VScan_StartScan
End Sub

Private Sub Command2_Click()
Dim b
b = MsgBox("Are you sure you want to exit?", 36, "Exit")
Select Case b
Case 6: End
End Select
End Sub

Private Sub VScan1_Infected()
If Option1.Value = True Then
Kill (Text1)
End If
If Option2.Value = ture Then
Beep
MsgBox "Eroki Virii Scan has found a virii!", vbExclamation, "Virii Scan"
End If
End Sub

Private Sub VScan1_NotInfected()
MsgBox "That file is not infected.", vbInformation, "Virii Scan"
End Sub

Private Sub VScan1_ScanMsg()
If Option3.Value = True Then
VScan1.VScan_Deltree (Text1)
End If
If Option4.Value = True Then
VScan1.VScan_VCL (Text1)
End If
If Option4.Value = True Then
VScan1.VScan_PWS (Text1)
End If
If Option5.Value = True Then
VScan1.VScan_WFBD (Text1)
End If
End Sub
