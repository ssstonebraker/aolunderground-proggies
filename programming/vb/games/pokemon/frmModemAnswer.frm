VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmModemAnswer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modem Connection - Answer"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   3240
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstBuffer 
      Height          =   60
      Left            =   -15
      TabIndex        =   5
      Top             =   1155
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.CommandButton cmdAnswer 
      Caption         =   "Answer"
      Enabled         =   0   'False
      Height          =   345
      Left            =   960
      TabIndex        =   4
      Top             =   795
      Width           =   1065
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   2115
      TabIndex        =   3
      Top             =   795
      Width           =   1065
   End
   Begin MSCommLib.MSComm Modem 
      Left            =   2400
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
      BaudRate        =   57600
   End
   Begin VB.Frame Frame1 
      Caption         =   "Modem"
      Height          =   660
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3135
      Begin VB.ComboBox lstPort 
         Height          =   315
         ItemData        =   "frmModemAnswer.frx":0000
         Left            =   705
         List            =   "frmModemAnswer.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   210
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Port:"
         Height          =   180
         Left            =   60
         TabIndex        =   2
         Top             =   255
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmModemAnswer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flgCancel As Boolean, flgAnswer As Boolean
Private Sub Answer()
    Dim strCommand As String, strReturn As String, strDummy As String
    strCommand = "ATA" & vbCr
    Modem.InBufferCount = 0
    Modem.Output = strCommand
    Do
        strDummy = DoEvents()
        If Modem.InBufferCount Then
            strReturn = strReturn + Modem.Input
            If InStr(strReturn, "CONNECT") Then
                Exit Do
            End If
        End If
        If flgCancel = True Then
            cmdAnswer.Enabled = False
            Modem.PortOpen = False
            flgCancel = False
            Exit Do
        End If
    Loop
    Do
        strDummy = DoEvents()
        If Modem.InBufferCount And Left(Modem.Input, 4) = "SET-" Then
            strData = Modem.Input
            Pokemon2 = Right(Left(strData, InStr(strData, ",") - 1), Len(Left(strData, InStr(strData, ",") - 1)) - 4)
            strData = Right(strData, Len(strData) - InStr(strData, ","))
            HP2 = Left(strData, InStr(strData, ",") - 1)
            strData = Right(strData, Len(strData) - InStr(strData, ","))
            ForeignID = strData
            MsgBox Pokemon2 & vbNewLine & HP2 & vbNewLine & ForeignID
            lstBuffer.AddItem "Bulbasaur"
            lstBuffer.ItemData(0) = 1
            Modem.Output = "SET-" & lstBuffer.ItemData(0) & "," & GetHealth(lstBuffer.ItemData(0)) & "," & frmMain.Player
            Exit Do
        ElseIf Modem.InBufferCount And Not Left(Modem.Input, 4) = "SET-" Then
            Modem.InBufferCount = 0
        End If
        If flgCancel = True Then
            cmdAnswer.Enabled = False
            Modem.PortOpen = False
            flgCancel = False
            Exit Do
        End If
    Loop
End Sub
Private Sub cmdAnswer_Click()
    flgAnswer = True
    Modem.PortOpen = True
End Sub
Private Sub lstPort_Click()
    Modem.CommPort = Right(lstPort.Text, 1)
    cmdAnswer.Enabled = True
End Sub
Private Sub cmdCancel_Click()
    If Modem.PortOpen = True Then
        flgCancel = True
        flgAnswer = False
    Else
        Unload Me
    End If
End Sub
Private Sub Modem_OnComm()
    If Modem.CommEvent = comEvRing And flgAnswer = True Then
        Answer
    End If
End Sub
Function GetHealth(num As Integer)
    GetHealth = 611
End Function
