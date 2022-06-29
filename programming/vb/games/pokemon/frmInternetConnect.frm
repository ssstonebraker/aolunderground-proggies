VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmInternetConnect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trainer Battle : Connect"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   2865
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTimer 
      Interval        =   1
      Left            =   -375
      Top             =   1020
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   8
      Top             =   1440
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   423
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   5001
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   930
      TabIndex        =   5
      Top             =   585
      Width           =   1860
   End
   Begin VB.ListBox lstBuffer 
      Height          =   255
      Left            =   -90
      TabIndex        =   4
      Top             =   1020
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Frame fraIP 
      Caption         =   "IP Connect"
      Height          =   990
      Left            =   15
      TabIndex        =   2
      Top             =   15
      Width           =   2835
      Begin VB.ComboBox txtIP 
         Height          =   315
         ItemData        =   "frmInternetConnect.frx":0000
         Left            =   915
         List            =   "frmInternetConnect.frx":0002
         TabIndex        =   7
         Top             =   225
         Width           =   1860
      End
      Begin MSWinsockLib.Winsock sckConnect 
         Left            =   2775
         Top             =   1020
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   7000
      End
      Begin VB.Label lblPassword 
         Alignment       =   1  'Right Justify
         Caption         =   "Code:"
         Height          =   255
         Left            =   75
         TabIndex        =   6
         Top             =   585
         Width           =   825
      End
      Begin VB.Label lblHN 
         Alignment       =   1  'Right Justify
         Caption         =   "Remote IP:"
         Height          =   255
         Left            =   75
         TabIndex        =   3
         Top             =   270
         Width           =   825
      End
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Enabled         =   0   'False
      Height          =   345
      Left            =   330
      TabIndex        =   1
      Top             =   1050
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   1485
      TabIndex        =   0
      Top             =   1050
      Width           =   1035
   End
End
Attribute VB_Name = "frmInternetConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intAction As Integer, strType As String
Private Sub cmdCancel_Click()
    If Not sckConnect.State = 0 Then
        sckConnect.Close
        txtIP.Enabled = True
        txtPassword.Enabled = True
        status.Tag = Empty
    Else
        Unload Me
    End If
End Sub
Private Sub Form_Activate()
    FormOnTop Me
End Sub
Private Sub cmdConnect_Click()
    If sckConnect.State > 0 Then
        MsgBox "You cannot connect if the port is not closed.", vbExclamation
    Else
        sckConnect.Connect txtIP.Text, 7000
        txtIP.Enabled = False
        txtPassword.Enabled = False
    End If
End Sub
Private Sub Form_Load()
    txtIP.AddItem "cimstudios.dynodns.net"
    LoadHosts txtIP
    txtPassword.Text = frmMain.Player + "_battle_001"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    txtIP.RemoveItem 0
    SaveHosts txtIP
    frmMain.Show
End Sub
Private Sub sckConnect_Connect()
    If intAction = 1 Then
        sckConnect.SendData "200 hello, i'm ready to battle"
    End If
    If intAction = 2 Then
        sckConnect.SendData "200, hello, i'm readt to trade"
    End If
End Sub
Private Sub tmrTimer_Timer()
    If status.Tag = "W" Then
        status.SimpleText = "Waiting for opponent to connect..."
        Exit Sub
    End If
    If sckConnect.State = 8 Then
        status.SimpleText = "Connection Terminated by Server."
        tmrTimer.Enabled = False
        Unload frmBattle
        Unload frmChatroom
        Unload Me
        frmMain.Hide
        MsgBox "Lost connection to remote computer."
        frmMain.Show
    End If
    If sckConnect.State = 9 Then
        status.SimpleText = "Error!"
        If Me.Tag = "battle" Then
            Unload frmBattle
            Unload frmChatroom
            Unload Me
            frmMain.Hide
            MsgBox "Lost connection to remote computer."
            frmMain.Show
            tmrTimer.Enabled = False
        End If
    End If
    If sckConnect.State = 0 Then
        status.SimpleText = "Not Connected."
    End If
    If sckConnect.State = 1 Then
        status.SimpleText = "Socket Open."
    End If
    If sckConnect.State = 3 Then
        status.SimpleText = "Connection Pending..."
    End If
    If sckConnect.State = 4 Then
        status.SimpleText = "Resolving Host..."
    End If
    If sckConnect.State = 5 Then
        status.SimpleText = "Host Resolved."
    End If
    If sckConnect.State = 6 Then
        status.SimpleText = "Connecting..."
    End If
    If sckConnect.State = 7 Then
        status.SimpleText = "Connected."
    End If
End Sub
Private Sub txtIP_Change()
    If Not txtIP = Empty Then
        cmdConnect.Enabled = True
    Else
        cmdConnect.Enabled = False
    End If
End Sub
Private Sub sckConnect_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    sckConnect.GetData strData
    If strData = "datServer" Then
        frmBattle.TurnA = 2
        sckConnect.SendData "datAccept" + txtPassword.Text
        strType = "Server"
    End If
    If strData = "datSession" Then
        Me.Hide
        MsgBox "There is already a session in use by that name!"
        Unload Me
    End If
    If strData = "datNone" Then
        status.Tag = "W"
    End If
    If strData = "datConnected" Then
        status.Tag = Empty
        lstBuffer.Clear
        LoadBench lstBuffer
        sckConnect.SendData "SET-" & lstBuffer.ItemData(0) & "," & GetHealth(lstBuffer.ItemData(0)) & "," & frmMain.Player
    End If
    If strData = "datDisconnected" Then
        Me.Hide
        MsgBox "Lost connection to remote computer through server!"
        Unload Me
        Unload frmBattle
        Unload frmChatroom
    End If
    If strData = "datSendData" Then
        lstBuffer.Clear
        LoadBench lstBuffer
        TimeOut 1
        sckConnect.SendData "re-SET1-" & lstBuffer.ItemData(0) & "," & GetHealth(lstBuffer.ItemData(0)) & "," & frmMain.Player
    End If
    If Left(strData, 5) = "SET1-" Then
        lstBuffer.Clear
        LoadBench lstBuffer
        sckConnect.SendData "re-SET-" & lstBuffer.ItemData(0) & "," & GetHealth(lstBuffer.ItemData(0)) & "," & frmMain.Player
        TimeOut 1
        strA = Left(strData, InStr(strData, ",") - 1)
        frmBattle.Pokemon2 = Right(Left(strData, InStr(strData, ",") - 1), Len(strA) - 4)
        strData = Right(strData, Len(strData) - InStr(strData, ","))
        frmBattle.HP2 = Left(strData, InStr(strData, ",") - 1)
        strData = Right(strData, Len(strData) - InStr(strData, ","))
        frmBattle.ForeignID = strData
        frmBattle.Pokemon1 = lstBuffer.ItemData(0)
        frmBattle.HP1 = GetHealth(frmBattle.Pokemon1)
        lstBuffer.Clear
        movelist LCase(N2N(frmBattle.Pokemon1)), lstBuffer
        frmBattle.Move1 = lstBuffer.List(0)
        frmBattle.Move2 = lstBuffer.List(1)
        frmBattle.Move3 = lstBuffer.List(2)
        frmBattle.Turn = 1
        Me.Hide
        frmBattle.Show
    End If
    If Left(strData, 4) = "SET-" Then
        strA = Left(strData, InStr(strData, ",") - 1)
        frmBattle.Pokemon2 = Right(Left(strData, InStr(strData, ",") - 1), Len(strA) - 4)
        strData = Right(strData, Len(strData) - InStr(strData, ","))
        frmBattle.HP2 = Left(strData, InStr(strData, ",") - 1)
        strData = Right(strData, Len(strData) - InStr(strData, ","))
        frmBattle.ForeignID = strData
        lstBuffer.Clear
        LoadBench lstBuffer
        frmBattle.Pokemon1 = lstBuffer.ItemData(0)
        frmBattle.HP1 = GetHealth(frmBattle.Pokemon1)
        lstBuffer.Clear
        movelist LCase(N2N(frmBattle.Pokemon1)), lstBuffer
        frmBattle.Move1 = lstBuffer.List(0)
        frmBattle.Move2 = lstBuffer.List(1)
        frmBattle.Move3 = lstBuffer.List(2)
        frmBattle.Turn = 0
        Me.Hide
        frmBattle.Show
    End If
    If Left(strData, 4) = "STE-" Then
        strA = Left(strData, InStr(strData, ",") - 1)
        frmBattle.Pokemon2 = Right(Left(strData, InStr(strData, ",") - 1), Len(strA) - 4)
        frmBattle.HP2 = Right(strData, Len(strData) - InStr(strData, ","))
        frmBattle.lblPokemon2.Caption = N2N(frmBattle.Pokemon2)
        frmBattle.imgPokemon2.Picture = frmPokedex.imgList.ListImages.Item(frmBattle.Pokemon2).Picture
        frmBattle.AddHP2 0
        frmBattle.Turn = 1
        frmBattle.SetStatus frmBattle.ForeignID & " sent out " & N2N(frmBattle.Pokemon2)
    End If
    If Left(strData, 4) = "PLU-" Then
        intRestore = Right(strData, Len(strData) - 4)
        frmBattle.AddHP2 intRestore
        frmBattle.Turn = 1
        frmBattle.SetStatus N2N(frmBattle.Pokemon2) & "'s health rose by " & intRestore
    End If
    If Left(strData, 4) = "MIN-" Then
        intAttack = Right(strData, Len(strData) - 4)
        frmBattle.MinusHP1 intAttack
        frmBattle.Turn = 1
        frmBattle.SetStatus frmBattle.ForeignID & "'s " & N2N(frmBattle.Pokemon2) & " hit for " & intAttack & " HP!"
    End If
    If Left(strData, 4) = "EFF-" Then
        strEffect = Right(strData, Len(strData) - 4)
        If strEffect = 1 Then
            frmBattle.SetStatus N2N(frmBattle.Pokemon1) + " is now poisoned!"
            SetEffect frmBattle.Pokemon1, 1
        End If
        If strEffect = 4 Then
            frmBattle.SetStatus N2N(frmBattle.Pokemon1) + " is paralyzed!"
            SetEffect frmBattle.Pokemon1, 2
        End If
    End If
    If strData = "220 fine with me!" Then
        lstBuffer.Clear
        LoadBench lstBuffer
        sckConnect.SendData "SET-" & lstBuffer.ItemData(0) & "," & GetHealth(lstBuffer.ItemData(0)) & "," & frmMain.Player
    End If
    If strData = "220 sorry, wrong password!" Then
        Me.Hide
        MsgBox "Incorrect password for access to " & txtIP.Text
        sckConnect.Close
        txtPassword.Enabled = True
        txtIP.Enabled = True
        Me.Show
    End If
    If strData = "LOSE" Then
        frmBattle.Hide
        frmWin.Show
    End If
    If strData = "SLOSE" Then
        frmBattle.SetStatus frmBattle.ForeignID & "'s " & N2N(frmBattle.Pokemon2) & " fainted!"
        frmBattle.Turn = 0
    End If
    If strData = "MISS" Then
        frmBattle.SetStatus frmBattle.ForeignID & " attacked and missed!"
        frmBattle.Turn = 1
    End If
    If Left(strData, 4) = "CHT-" Then
        frmChatroom.txtChat.Text = frmChatroom.txtChat.Text & vbNewLine & frmBattle.ForeignID & ":" & Chr(9) & Right(strData, Len(strData) - 4)
        frmChatroom.txtChat.SelStart = Len(frmChatroom.txtChat.Text)
        frmBattle.lblChat.ForeColor = &HFF&
    End If
    If Left(strData, 4) = "ITM-" Then
        strItem = Right(strData, Len(strData) - 4)
        frmBattle.SetStatus frmBattle.ForeignID & " used " & stritme & "!"
        frmBattle.Turn = 1
    End If
    If Left(strData, 4) = "CAS-" Then
        strMoney% = Right(strData, Len(strData) - 4)
        SaveMoney strMoney%
        frmWin.lblCaption = "Recieved $" & strMoney% & " For Winning"
        frmBattle.SetStatus frmBattle.ForeignID & " ran from the battle!"
        frmBattle.Hide
        frmChatroom.Hide
        frmWin.Show
    End If
    If Left(strData, 4) = "PKM-" Then
        strPokemon% = Right(strData, Len(strData) - 4)
        SavePokemon strPokemon%
    End If
    If strData = "invTYPEa" Then
        MsgBox "The host you are connecting to is not Battling.", vbExclamation
        sckConnect.Close
    End If
    If strData = "invTYPEb" Then
        MsgBox "The host you are connecting to is not Trading.", vbExclamation
        sckConnect.Close
    End If
    If strData = "typeOK" Then
        If intAction = 1 Then
            If StringInList(sckConnect.RemoteHostIP, txtIP) = False Then
                txtIP.AddItem sckConnect.RemoteHostIP
            End If
            sckConnect.SendData "200 is this password ok? " & txtPassword.Text
        End If
        If intAction = 2 Then
            SendBench
        End If
    End If
    If Left(strData, 6) = "bench-" Then
        strBuffer = Right(strData, Len(strData) - 6)
        strPokemon = Left(strBuffer, InStr(strBuffer, "|"))
        strHealth = Right(strBuffer, Len(strBuffer) - InStr(strBuffer, "|"))
        frmTrade.lstBench2.AddItem N2N(strPokemon)
        frmTrade.lstBench2.ItemData(frmTrade.lstBench2.ListCount) = strPokemon
        frmTrade.lstBench2HP.AddItem strHealth
    End If
    If strData = "actTRADEOPEN" Then
        frmTrade.Show
        Me.Hide
    End If
    If Left(strData, 9) = "datOFFER-" Then
        strBuffer = Right(strData, Len(strData) - 9)
        strPokemon1% = Left(strBuffer, InStr(strBuffer, "|"))
        strPokemon2 = Right(strBuffer, Len(strBuffer) - InStr(strBuffer, "|"))
        If MsgBox("Trade Offered!" & vbNewLine & vbNewLine & N2N(strPokemon1%) & vbNewLine & "for" & vbNewLine & N2N(strPokemon2), vbQuestion) = vbYes Then
            sckListen.SendData "datTRADEACCEPT-" & strPokemon2
            TimeOut 0.5
            sckListen.SendData "PKM-" & strPokemon1%
            DeletePokemon strPokemon1%
        Else
            sckListen.SendData "datTRADEDENIED"
        End If
    End If
    If Left(strData, 14) = "datTRADEACCEPT" Then
        frmTrade.Hide
        MsgBox "The Trade Was Accepted!", vbInformation
        frmTrade.Show
        sckConnect.SendData "PKM-" & Right(strData, Len(strData) - 14)
        DeletePokemon Right(strData, Len(strData) - 14)
    End If
    If strData = "datTRADEDENIED" Then
        MsgBox "The Trade Was Declined!"
        frmTrade.lblTrade.Visible = True
        frmTrade.lblCancel.Visible = True
    End If
End Sub
Function SendBench()
    lstBuffer.Clear
    Dim strBuffer As String
    Dim intBuffer As Integer
    intBuffer = 0
loopA:
    If intBuffer = lstBuffer.ListCount Then
        sckConnect.SendData "reqBENCH"
        Exit Function
    Else
        sckConnect.SendData "bench-" & lstBuffer.ItemData(intBuffer) & "|" & GetHealth(lstBuffer.ItemData(intBuffer))
        MsgBox "bench-" & lstBuffer.ItemData(intBuffer) & "|" & GetHealth(lstBuffer.ItemData(intBuffer))
        TimeOut 0.5
        GoTo loopA
    End If
End Function
Private Sub txtIP_Click()
    If Not txtIP = Empty Then
        cmdConnect.Enabled = True
    Else
        cmdConnect.Enabled = False
    End If
End Sub
Private Sub txtPassword_Change()
    If InStr(txtPassword.Text, "~") Then
        MsgBox "Password can NOT contain a tilda!", vbExclamation
    End If
    txtPassword.Text = ReplaceString(txtPassword.Text, "~", "")
End Sub
