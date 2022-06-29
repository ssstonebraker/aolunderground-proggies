VERSION 5.00
Object = "{40B4C461-FE90-11D2-B1AA-E02862C10000}#1.0#0"; "VFCDWR~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame Frame4 
      Caption         =   "ISO"
      Height          =   2925
      Left            =   3585
      TabIndex        =   15
      Top             =   1335
      Width           =   1950
      Begin VB.CommandButton Command10 
         Caption         =   "CrtMultiSessISO"
         Height          =   345
         Left            =   285
         TabIndex        =   18
         Top             =   1695
         Width           =   1350
      End
      Begin VB.CommandButton Command9 
         Caption         =   "WriteISOtoCDR"
         Height          =   345
         Left            =   285
         TabIndex        =   17
         Top             =   1245
         Width           =   1350
      End
      Begin VB.CommandButton Command7 
         Caption         =   "ISO from Path"
         Height          =   345
         Left            =   285
         TabIndex        =   16
         Top             =   825
         Width           =   1350
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Audio"
      Height          =   1155
      Left            =   3585
      TabIndex        =   11
      Top             =   165
      Width           =   1950
      Begin VB.CommandButton Command4 
         Caption         =   "AddWave"
         Height          =   345
         Left            =   390
         TabIndex        =   14
         Top             =   255
         Width           =   1155
      End
      Begin VB.CommandButton Command5 
         Caption         =   "WriteWaves"
         Height          =   345
         Left            =   405
         TabIndex        =   13
         Top             =   690
         Width           =   1170
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "SCSI Commands"
      Height          =   2925
      Left            =   1275
      TabIndex        =   7
      Top             =   1335
      Width           =   2265
      Begin VB.CommandButton Command8 
         Caption         =   "SCSI Contr."
         Height          =   345
         Left            =   540
         TabIndex        =   20
         Top             =   1518
         Width           =   1185
      End
      Begin VB.CommandButton Command11 
         Caption         =   "SCSI Name"
         Height          =   345
         Left            =   540
         TabIndex        =   19
         Top             =   1920
         Width           =   1185
      End
      Begin VB.CheckBox Check1 
         Caption         =   "TestMode"
         Height          =   345
         Left            =   555
         TabIndex        =   12
         Top             =   2400
         Value           =   1  'Aktiviert
         Width           =   1185
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Scan SCSI"
         Height          =   345
         Left            =   540
         TabIndex        =   10
         Top             =   1117
         Width           =   1185
      End
      Begin VB.CommandButton Command2 
         Caption         =   "BlankDisk"
         Height          =   345
         Left            =   540
         TabIndex        =   9
         Top             =   716
         Width           =   1185
      End
      Begin VB.CommandButton Command3 
         Caption         =   "ResetDevice"
         Height          =   345
         Left            =   540
         TabIndex        =   8
         Top             =   315
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "SetDevice"
      Height          =   1155
      Left            =   1275
      TabIndex        =   2
      Top             =   165
      Width           =   2265
      Begin VB.CommandButton Command6 
         Caption         =   "SetDevice"
         Height          =   315
         Left            =   570
         TabIndex        =   6
         Top             =   765
         Width           =   1140
      End
      Begin VB.TextBox Text3 
         Height          =   345
         Left            =   1500
         TabIndex        =   5
         Text            =   "0"
         Top             =   330
         Width           =   450
      End
      Begin VB.TextBox Text2 
         Height          =   345
         Left            =   922
         TabIndex        =   4
         Text            =   "0"
         Top             =   330
         Width           =   450
      End
      Begin VB.TextBox Text1 
         Height          =   345
         Left            =   345
         TabIndex        =   3
         Text            =   "0"
         Top             =   330
         Width           =   450
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2595
      Top             =   2925
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "wav"
      DialogTitle     =   "Open Wave File"
      Filter          =   "Wave|*.wav"
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   150
      TabIndex        =   0
      Top             =   4440
      Width           =   6870
   End
   Begin VFCDWRITERLib.VFcdwriter VFcdwriter1 
      Left            =   3450
      Top             =   3105
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Status :"
      Height          =   225
      Left            =   180
      TabIndex        =   1
      Top             =   4140
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim writererror As Integer
Private Sub Check1_Click()
VFcdwriter1.TestMode = Check1.Value
End Sub

Private Sub Command1_Click()
Form2.Show
End Sub

Private Sub Command10_Click()
Dim ans As Integer
VFcdwriter1.ClearISOCue
a = VFcdwriter1.AddFileOrPathToISOcue("d:\crypt", True, "\test\")
VFcdwriter1.Joliet = True
a = VFcdwriter1.CreateMultiISOFromCue("d:\test2.iso")
If writererror = 1 Then
    a = MsgBox("There were Errors !! Aborting Command ...", vbCritical)
    Exit Sub
End If
    
a = MsgBox("Create CD ?", vbYesNo, "CDWriter Test")

If a = vbNo Then
    Exit Sub
End If
VFcdwriter1.SessionType = MultiSession
a = VFcdwriter1.WriteISOtoCDR("d:\test2.iso")
End Sub

Private Sub Command11_Click()
MsgBox "The Name of your Adapter is " & VFcdwriter1.GetHostAdapterName(0)
End Sub

Private Sub Command12_Click()
VFcdwriter1.TestTest
End Sub

Private Sub Command2_Click()
List1.Clear
a = VFcdwriter1.BlankDisk()
End Sub

Private Sub Command3_Click()
List1.Clear
VFcdwriter1.ResetDrive
End Sub

Private Sub Command4_Click()
Dim filename As String
CommonDialog1.ShowOpen
filename = CommonDialog1.filename
If filename <> "" Then
    
    VFcdwriter1.AddWaveToCue (filename)
End If
End Sub

Private Sub Command5_Click()
VFcdwriter1.WriteWavecueToCDR
End Sub

Private Sub Command6_Click()
VFcdwriter1.SCSIController = Text1.Text
VFcdwriter1.Target = Text2.Text
VFcdwriter1.Lun = Text3.Text
End Sub

Private Sub Command7_Click()
VFcdwriter1.ClearISOCue
an = InputBox("Please select a valid path :")
If an = "" Then
    Exit Sub
End If
a = VFcdwriter1.AddFileOrPathToISOcue(an, True, "\mycdrom\")
VFcdwriter1.Joliet = True
an = InputBox("Enter the name of the new iso file")
If an = "" Then
    Exit Sub
End If
a = VFcdwriter1.CreateISOFromCue(an)

End Sub

Private Sub command8_Click()
MsgBox "You have " & VFcdwriter1.GetNumOfAdapters & " SCSI-Adapter(s)"

End Sub

Private Sub Command9_Click()
VFcdwriter1.SessionType = MultiSession
VFcdwriter1.Speed = 2
an = InputBox("Enter a iso image (with path) for writing to cdr")
If an = "" Then
Exit Sub
End If
a = VFcdwriter1.WriteISOtoCDR(an)
End Sub

Private Sub Form_Load()
MsgBox "Make sure that wnaspi32.dll Version 4.57 (1008) is installed on your system ! Otherwise you may get problems writing a cd !"
writererror = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
If VFcdwriter1.IsTaskRunning = True Then
    Cancel = -1
    MsgBox "SCSI Command still running ... Can't quit !"
End If
End Sub

Private Sub VFcdwriter1_CommandDone()
MsgBox "done ...."
End Sub

Private Sub VFcdwriter1_OnError(ByVal errorcode As Integer, ByVal errmsg As String)
List1.AddItem (errorcode & " " & errmsg)
If errorcode = 800 Or errorcode = 900 Then
    writererror = 1
End If
End Sub

Private Sub VFcdwriter1_OnSCSIScan(ByVal Adapter As Integer, ByVal ID As Integer, ByVal Lun As Integer, ByVal Product As String, ByVal ProductIdent As String, ByVal ProductRevision As String, ByVal devType As Integer)
Dim scanstring
scanstring = "Adapter " & Adapter & " ID = " & ID & " Lun = " & Lun & " Name : " & Product & " ProductIdent : " & ProductIdent & " ProductRevision : " & ProductRevision & " Typ : " & devType
Form2.List1.AddItem scanstring
End Sub
