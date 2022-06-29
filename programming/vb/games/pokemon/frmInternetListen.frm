VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmInternetListen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trainer Battle : Listen"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   2910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTimer 
      Interval        =   1
      Left            =   -375
      Top             =   720
   End
   Begin VB.ListBox lstBuffer 
      Height          =   255
      Left            =   -75
      TabIndex        =   5
      Top             =   855
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   1500
      TabIndex        =   4
      Top             =   765
      Width           =   1035
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "Listen"
      Height          =   345
      Left            =   315
      TabIndex        =   3
      Top             =   765
      Width           =   1035
   End
   Begin VB.Frame fraIP 
      Caption         =   "IP Listen"
      Height          =   690
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   2835
      Begin MSWinsockLib.Winsock sckListen 
         Left            =   -375
         Top             =   -375
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   7000
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   915
         TabIndex        =   2
         Top             =   240
         Width           =   1860
      End
      Begin VB.Label lblPassword 
         Alignment       =   1  'Right Justify
         Caption         =   "Code:"
         Height          =   255
         Left            =   90
         TabIndex        =   1
         Top             =   270
         Width           =   825
      End
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   6
      Top             =   1140
      Width           =   2910
      _ExtentX        =   5133
      _ExtentY        =   423
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmInternetListen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intAction As Integer
Private Sub cmdCancel_Click()
    If sckListen.State > 0 Then
        sckListen.Close
        txtPassword.Enabled = True
    Else
        Unload Me
    End If
End Sub
Private Sub cmdListen_Click()
    If sckListen.State > 0 Then
        MsgBox "You cannot listen while the port is still open!"
    Else
        sckListen.Listen
        If Err Then
            Me.Hide
            MsgBox "The TCP/IP protocol is unavailable or port 7000 is in use!"
            Me.Show
        End If
        txtPassword.Enabled = False
    End If
End Sub
Private Sub Form_Activate()
    FormOnTop Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmMain.Show
End Sub
Private Sub sckListen_ConnectionRequest(ByVal requestID As Long)
    If sckListen.State > 0 Then sckListen.Close
    sckListen.Accept requestID
End Sub
Private Sub sckListen_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    sckListen.GetData strData
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
        sckListen.SendData "SET-" & lstBuffer.ItemData(0) & "," & GetHealth(lstBuffer.ItemData(0)) & "," & frmMain.Player
        lstBuffer.Clear
        movelist LCase(N2N(frmBattle.Pokemon1)), lstBuffer
        frmBattle.Move1 = lstBuffer.List(0)
        frmBattle.Move2 = lstBuffer.List(1)
        frmBattle.Move3 = lstBuffer.List(2)
        frmBattle.Turn = 1
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
    If Left(strData, 25) = "200 is this password ok? " Then
        strPassword = Right(strData, Len(strData) - 25)
        If strPassword = txtPassword.Text Then
            sckListen.SendData "220 fine with me!"
        Else
            sckListen.SendData "220 sorry, wrong password!"
            TimeOut 1
            sckListen.Close
            sckListen.Listen
        End If
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
    If strData = "200 hello, i'm ready to trade" Then
        If intAction = 1 Then
            sckListen.SendData "invTYPEa"
            TimeOut 0.5
            sckListen.Close
        Else
            sckListen.SendData "typeOK"
        End If
    End If
    If strData = "200 hello, i'm ready to battle" Then
        If intAction = 2 Then
            sckListen.SendData "invTYPEb"
            TimeOut 0.5
            sckListen.Close
        Else
            sckListen.SendData "typeOK"
        End If
    End If
    If Left(strData, 6) = "bench-" Then
        strBuffer = Right(strData, Len(strData) - 6)
        strPokemon = Left(strBuffer, InStr(strBuffer, "|"))
        strHealth = Right(strBuffer, Len(strBuffer) - InStr(strBuffer, "|"))
        frmTrade.lstBench2.AddItem N2N(strPokemon)
        frmTrade.lstBench2.ItemData(frmTrade.lstBench2.ListCount) = strHealth
        frmTrade.lstBench2HP.AddItem strHealth
    End If
    If strData = "reqBENCH" Then
        SendBench
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
        sckListen.SendData "actTRADEOPEN"
        Exit Function
    Else
        sckListen.SendData "bench-" & lstBuffer.ItemData(intBuffer) & "|" & GetHealth(lstBuffer.ItemData(intBuffer))
        MsgBox "bench-" & lstBuffer.ItemData(intBuffer) & "|" & GetHealth(lstBuffer.ItemData(intBuffer))
        TimeOut 0.5
        GoTo loopA
    End If
End Function
Private Sub tmrTimer_Timer()
    If sckListen.State = 0 Then
        status.SimpleText = "Not Connected."
    End If
    If sckListen.State = 2 Then
        status.SimpleText = "Waiting For Connection..."
    End If
    If sckListen.State = 1 Then
        status.SimpleText = "Socket Open."
    End If
    If sckListen.State = 3 Then
        status.SimpleText = "Connection Pending..."
    End If
    If sckListen.State = 4 Then
        status.SimpleText = "Resolving Host..."
    End If
    If sckListen.State = 5 Then
        status.SimpleText = "Host Resolved."
    End If
    If sckListen.State = 6 Then
        status.SimpleText = "Connecting..."
    End If
    If sckListen.State = 7 Then
        status.SimpleText = "Connected."
    End If
    If sckListen.State = 8 Then
        status.SimpleText = "Connection Terminated by Server."
        tmrTimer.Enabled = False
        Unload frmBattle
        Unload frmChatroom
        Unload Me
        frmMain.Hide
        MsgBox "Lost connection to remote computer."
        frmMain.Show
    End If
    If sckListen.State = 9 Then
        status.SimpleText = "Error!"
        tmrTimer.Enabled = False
        Unload frmBattle
        Unload frmChatroom
        Unload Me
        frmMain.Hide
        MsgBox "Lost connection to remote computer."
        frmMain.Show
    End If
End Sub
