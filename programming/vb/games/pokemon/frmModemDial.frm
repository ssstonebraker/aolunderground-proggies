VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmModemDial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modem Connection - Dial"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3210
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   3210
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstBuffer 
      Height          =   255
      Left            =   30
      TabIndex        =   8
      Top             =   1950
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   2115
      TabIndex        =   7
      Top             =   1620
      Width           =   1065
   End
   Begin VB.CommandButton cmdDial 
      Caption         =   "Dial"
      Enabled         =   0   'False
      Height          =   345
      Left            =   960
      TabIndex        =   6
      Top             =   1620
      Width           =   1065
   End
   Begin MSCommLib.MSComm Modem 
      Left            =   2070
      Top             =   735
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
      TabIndex        =   3
      Top             =   900
      Width           =   3135
      Begin VB.ComboBox lstPort 
         Height          =   315
         ItemData        =   "frmModemDial.frx":0000
         Left            =   705
         List            =   "frmModemDial.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   210
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Port:"
         Height          =   180
         Left            =   60
         TabIndex        =   4
         Top             =   255
         Width           =   600
      End
   End
   Begin VB.Frame fraDial 
      Caption         =   "Dial Properties"
      Height          =   795
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   3135
      Begin VB.TextBox txtNumber 
         Height          =   285
         Left            =   1425
         TabIndex        =   2
         Top             =   270
         Width           =   1650
      End
      Begin VB.Label lblNumber 
         Alignment       =   1  'Right Justify
         Caption         =   "Number:"
         Height          =   225
         Left            =   720
         TabIndex        =   1
         Top             =   300
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmModemDial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flgCancel As Boolean
Private Sub cmdCancel_Click()
    If Modem.PortOpen = True Then
        flgCancel = True
        Modem.PortOpen = False
    Else
        Unload Me
    End If
End Sub
Private Sub cmdDial_Click()
    Dim strCommand As String, strReturn As String, strDummy As String
    strCommand = "ATDT " & txtNumber.Text & vbCr
    On Error Resume Next
    Modem.PortOpen = True
    If Err Then
        MsgBox "The Communications port is in use or not available!"
        Exit Sub
    End If
    Modem.InBufferCount = 0
    Modem.Output = strCommand
    Do
        strDummy = DoEvents()
        If Modem.InBufferCount Then
            strReturn = strReturn + Modem.Input
            If InStr(strReturn, "CONNECT") Then
                TimeOut 5
                lstBuffer.AddItem "Bulbasaur"
                lstBuffer.ItemData(0) = 1
                Modem.Output = "SET-" & lstBuffer.ItemData(0) & "," & GetHealth(lstBuffer.ItemData(0)) & "," & "Andrew" 'frmMain.Player
                Exit Do
            End If
        End If
        If flgCancel = True Then
            cmdDial.Enabled = False
            Modem.PortOpen = False
            flgCancel = False
            Exit Sub
        End If
    Loop
    Do
        strDummy = DoEvents()
        If Modem.InBufferCount Then
            strReturn = Modem.Input
            If Left(strReturn, 4) = "SET-" Then
                strData = Modem.Input
                Pokemon2 = Right(Left(strData, InStr(strData, ",") - 1), Len(Left(strData, InStr(strData, ",") - 1)) - 4)
                strData = Right(strData, Len(strData) - InStr(strData, ","))
                HP2 = Left(strData, InStr(strData, ",") - 1)
                strData = Right(strData, Len(strData) - InStr(strData, ","))
                ForeignID = strData
                MsgBox Pokemon2 & vbNewLine & HP2 & vbNewLine & ForeignID
                Exit Do
            End If
        End If
        If flgCancel = True Then
            cmdAnswer.Enabled = False
            Modem.PortOpen = False
            flgCancel = False
            Exit Sub
        End If
    Loop
End Sub
Private Sub lstPort_Change()
    Modem.CommPort = Right(lstPort.Text, 1)
End Sub
Private Sub txtNumber_Change()
    If txtNumber.Text = Empty Then
        cmdDial.Enabled = False
    Else
        cmdDial.Enabled = True
    End If
End Sub
Function GetHealth(num As Integer)
    GetHealth = 611
End Function
