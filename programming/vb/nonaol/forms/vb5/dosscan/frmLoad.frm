VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLoad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "loading a bas file"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5430
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdgBas 
      Left            =   5160
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      MaxFileSize     =   32000
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.TextBox txtInfo 
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Text            =   "frmLoad.frx":0000
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   5415
      Begin VB.TextBox txtBasTitle 
         Height          =   315
         Left            =   4440
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtProcedureCount 
         Height          =   315
         Left            =   3480
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtFileName 
         Height          =   315
         Left            =   480
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtCode 
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   960
         Width           =   5175
      End
      Begin VB.ComboBox cmbProcedure 
         Height          =   330
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   5175
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "procedures:"
         Height          =   210
         Left            =   2520
         TabIndex        =   9
         Top             =   330
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "title:"
         Height          =   210
         Left            =   4080
         TabIndex        =   8
         Top             =   330
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "file:"
         Height          =   210
         Left            =   120
         TabIndex        =   4
         Top             =   330
         Width           =   255
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&file"
      Begin VB.Menu mnuLoad 
         Caption         =   "&load bas"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&exit"
      End
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim BasStr As String, Code() As String

Private Sub cmbProcedure_Click()
    If cmbProcedure.ListIndex <> -1 Then
        txtCode.Text = Code(cmbProcedure.ListIndex)
    End If
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuLoad_Click()
    Dim Files As String
    cdgBas.CancelError = True
    cdgBas.FileName = ""
    On Error GoTo ErrHandler
    cdgBas.ShowOpen
    cmbProcedure.Clear
    txtCode.Text = ""
    txtFileName.Text = ""
    txtProcedureCount.Text = ""
    txtBasTitle.Text = ""
    Call LoadString(BasStr$, cdgBas.FileName)
    Call LoadBasFile
    txtFileName.Text = cdgBas.FileName
    txtBasTitle.Text = GetBasTitle(BasStr$)
    txtProcedureCount.Text = CStr(cmbProcedure.ListCount)
    cmbProcedure.Text = cmbProcedure.List(0)
ErrHandler:
    Exit Sub
End Sub

Private Sub LoadBasFile()
    Dim StartStr1 As String, StartStr2 As String, StartStr3 As String
    Dim StartStr4 As String, StartStr5 As String, StartStr6 As String
    Dim TheEnd As Long, EndChar As Long, StartSpot As Long
    Dim Spot1 As Long, Spot2 As Long, Spot3 As Long
    Dim Spot4 As Long, Spot5 As Long, Spot6 As Long
    Dim SubStr As String, sEnd As Long, fEnd As Long
    Dim sTitle As String, SubCounter As Long
    StartStr1$ = vbCrLf & "Sub"
    StartStr2$ = vbCrLf & "Public Sub"
    StartStr3$ = vbCrLf & "Private Sub"
    StartStr4$ = vbCrLf & "Function"
    StartStr5$ = vbCrLf & "Public Function"
    StartStr6$ = vbCrLf & "Private Function"
    TheEnd& = 1&
    EndChar& = 1&
    Spot1& = InStr(BasStr$, StartStr1$)
    Spot2& = InStr(BasStr$, StartStr2$)
    Spot3& = InStr(BasStr$, StartStr3$)
    Spot4& = InStr(BasStr$, StartStr4$)
    Spot5& = InStr(BasStr$, StartStr5$)
    Spot6& = InStr(BasStr$, StartStr6$)
    SubCounter& = 0&
    Do
        DoEvents
        StartSpot& = 0&
        StartSpot& = Spot1& + Spot2& + Spot3& + Spot4& + Spot5& + Spot6&
        If Spot1& <> 0& And Spot1& < StartSpot& Then
            StartSpot& = Spot1&
        End If
        If Spot2& <> 0& And Spot2& < StartSpot& Then
            StartSpot& = Spot2&
        End If
        If Spot3& <> 0& And Spot3& < StartSpot& Then
            StartSpot& = Spot3&
        End If
        If Spot4& <> 0& And Spot4& < StartSpot& Then
            StartSpot& = Spot4&
        End If
        If Spot5& <> 0& And Spot5& < StartSpot& Then
            StartSpot& = Spot5&
        End If
        If Spot6& <> 0& And Spot6& < StartSpot& Then
            StartSpot& = Spot6&
        End If
        sEnd& = InStr(StartSpot&, BasStr$, Chr(10) & "End Sub")
        fEnd& = InStr(StartSpot&, BasStr$, Chr(10) & "End Function")
        If sEnd& = 0& And fEnd& = 0& Then
            Exit Sub
        End If
        If sEnd& = 0& Then
            sEnd& = 2000000000
        End If
        If fEnd& = 0& Then
            fEnd& = 2000000000
        End If
        If sEnd& < fEnd& Then
            TheEnd& = sEnd&
        Else
            TheEnd& = fEnd&
        End If
        EndChar& = InStr(TheEnd&, BasStr$, Chr(13))
        SubStr$ = Mid(BasStr$, StartSpot& + 2&, EndChar& - StartSpot& - 2&)
        sTitle$ = GetSubTitle(SubStr$)
        cmbProcedure.AddItem sTitle$
        ReDim Preserve Code(cmbProcedure.ListCount)
        Code(SubCounter&) = SubStr$
        Spot1& = InStr(EndChar&, BasStr$, StartStr1$)
        Spot2& = InStr(EndChar&, BasStr$, StartStr2$)
        Spot3& = InStr(EndChar&, BasStr$, StartStr3$)
        Spot4& = InStr(EndChar&, BasStr$, StartStr4$)
        Spot5& = InStr(EndChar&, BasStr$, StartStr5$)
        Spot6& = InStr(EndChar&, BasStr$, StartStr6$)
        SubCounter& = SubCounter& + 1&
    Loop Until Spot1& = 0& And Spot2& = 0& And Spot3& = 0& And Spot4& = 0& And Spot5& = 0& And Spot6& = 0&
End Sub
