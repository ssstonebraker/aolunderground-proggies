VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Multi-User Rich Text IP Chat"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstUsers 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1890
      Left            =   4890
      Sorted          =   -1  'True
      TabIndex        =   19
      Top             =   195
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock sckConnect 
      Index           =   0
      Left            =   6360
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtbChat 
      Height          =   1890
      Left            =   75
      TabIndex        =   18
      Top             =   195
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3334
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":0000
   End
   Begin MSComDlg.CommonDialog dlgColors 
      Left            =   6720
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdColors 
      Height          =   320
      Left            =   3120
      Picture         =   "frmMain.frx":00C1
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2175
      Width           =   315
   End
   Begin RichTextLib.RichTextBox rtbText 
      Height          =   360
      Left            =   0
      TabIndex        =   16
      Top             =   2475
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   635
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":0403
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox chkUnderline 
      Height          =   320
      Left            =   4200
      Picture         =   "frmMain.frx":04C4
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2175
      Width           =   315
   End
   Begin VB.CheckBox chkItalic 
      Height          =   320
      Left            =   3900
      Picture         =   "frmMain.frx":0806
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2175
      Width           =   315
   End
   Begin VB.ComboBox cmbFonts 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2175
      Width           =   3000
   End
   Begin VB.CheckBox chkBold 
      Height          =   320
      Left            =   3600
      Picture         =   "frmMain.frx":0B48
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2175
      UseMaskColor    =   -1  'True
      Width           =   315
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   405
      Left            =   5070
      TabIndex        =   11
      Top             =   2895
      Width           =   975
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   2760
      TabIndex        =   10
      Text            =   "localhost"
      Top             =   3360
      Width           =   2250
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   360
      TabIndex        =   8
      Text            =   "400"
      Top             =   3360
      Width           =   2055
   End
   Begin VB.OptionButton optServerClient 
      Caption         =   "Client"
      Height          =   195
      Index           =   1
      Left            =   4200
      TabIndex        =   6
      Top             =   2995
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.OptionButton optServerClient 
      Caption         =   "Server"
      Height          =   195
      Index           =   0
      Left            =   3360
      TabIndex        =   5
      Top             =   2995
      Width           =   855
   End
   Begin VB.TextBox txtNick 
      Height          =   285
      Left            =   440
      TabIndex        =   4
      Text            =   "NickName"
      Top             =   2940
      Width           =   2775
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   335
      Left            =   5400
      TabIndex        =   2
      Top             =   2490
      Width           =   625
   End
   Begin VB.Frame frmeSep 
      Height          =   135
      Left            =   0
      TabIndex        =   1
      Top             =   2760
      Width           =   6060
   End
   Begin VB.Frame frmeChatWindow 
      Caption         =   "Chat Window"
      Height          =   2165
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6045
   End
   Begin VB.Shape shpGreen 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008000&
      Height          =   255
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   255
   End
   Begin VB.Shape shpRed 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      FillColor       =   &H00000080&
      Height          =   255
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label lblIP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP:"
      Height          =   195
      Left            =   2520
      TabIndex        =   9
      Top             =   3390
      Width           =   195
   End
   Begin VB.Label lblPort 
      BackStyle       =   0  'Transparent
      Caption         =   "Port:"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   3390
      Width           =   375
   End
   Begin VB.Line lneSep3 
      BorderColor     =   &H80000003&
      Index           =   1
      X1              =   5025
      X2              =   5025
      Y1              =   2880
      Y2              =   3290
   End
   Begin VB.Line lneSep3 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   5040
      X2              =   5040
      Y1              =   2880
      Y2              =   3300
   End
   Begin VB.Line lneSep2 
      BorderColor     =   &H80000003&
      Index           =   1
      X1              =   3240
      X2              =   3240
      Y1              =   3270
      Y2              =   2880
   End
   Begin VB.Line lneSep 
      BorderColor     =   &H80000003&
      Index           =   1
      X1              =   0
      X2              =   5040
      Y1              =   3270
      Y2              =   3270
   End
   Begin VB.Line lneSep2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   3255
      X2              =   3255
      Y1              =   3290
      Y2              =   2880
   End
   Begin VB.Line lneSep 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   0
      X2              =   5040
      Y1              =   3285
      Y2              =   3285
   End
   Begin VB.Label lblNick 
      BackStyle       =   0  'Transparent
      Caption         =   "Nick: "
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   2970
      Width           =   495
   End
   Begin VB.Menu mnuList 
      Caption         =   "mnuList"
      Visible         =   0   'False
      Begin VB.Menu mnuKickUser 
         Caption         =   "Kick User"
         Index           =   0
      End
      Begin VB.Menu mnuKickUser 
         Caption         =   "Kick User (why?)"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************
'*     Multi User Rich Text IP Chat by Joseph Huntley     *
'*               joseph_huntley@email.com                 *
'*                http://joseph.vr9.com                   *
'*                                                        *
'*  Made:  October 21, 1999                                *
'*  Level: Intermediate/Advanced                          *
'**********************************************************
'* Notes: I might make a regular text IP chat without all *
'*        this fancy stuff, so it can be understood how   *
'*        it's done better.                               *
'**********************************************************


Private Const vbDarkRed = &H80&
Private Const vbDarkGreen = &H8000&

Private strUsers() As String 'Array to hold the nick of the person connecting by index

Sub AddChat(strNick As String, strRTF As String)

''Adds someone's nick and what they said to rtbChat

  Dim lngLastLen As Long
  
  ''set selected position to length of text
  rtbChat.SelStart = Len(rtbChat.Text)
  rtbChat.SelLength = 0
  
  ''set the seltext to a new line plus "Nick:" and tab character
  rtbChat.SelText = vbCrLf & strNick$ & ":" & vbTab
  
  ''change color, size, font name, and font styles
  rtbChat.SelStart = Len(rtbChat.Text) - (Len(strNick$) + 4) '4 = Length of vbCrLf + ':' + vbTab
  rtbChat.SelLength = Len(strNick$) + 4
  rtbChat.SelColor = vbBlue
  rtbChat.SelFontSize = 8
  rtbChat.SelFontName = "Arial"
  rtbChat.SelBold = True
  rtbChat.SelUnderline = False
  rtbChat.SelItalic = False
  
  ''store length of text so we can have a hangingindent later
  lngLastLen& = Len(rtbChat.Text)
  
  ''set selstart & sellength then add the rtf string
  rtbChat.SelStart = lngLastLen&
  rtbChat.SelLength = 0
  rtbChat.SelRTF = strRTF$
  
  ''now set the hanging indent
  rtbChat.SelStart = lngLastLen&
  rtbChat.SelLength = Len(rtbChat.Text) - lngLastLen&
  rtbChat.SelHangingIndent = 1400
  
  ''scroll textbox down
  rtbChat.SelStart = Len(rtbChat.Text)
  rtbChat.SelLength = 0
  
  ''set focus to rtbText
  rtbText.SetFocus

End Sub
Sub ParseData(strData As String, Index As Integer)

  ''used to parse data then propery use the arguments
  
  Dim strCommand As String, strArgument As String, strBuf1 As String
  Dim strBuf2 As String, lngBuf As Long, lngBuf2 As Long, lngPos As Long, lngIndex As Long

  ''store command and argument. Syntax is: '[Command] Argument'
  strCommand$ = Left$(strData$, InStr(strData$, " ") - 1)
  strArgument$ = Right$(strData$, Len(strData$) - InStr(strData$, " "))

      Select Case UCase$(Mid$(strCommand$, 2, Len(strCommand$) - 2))
         
         Case "MESSAGE":
            ''store nick and rtf message
            strBuf1$ = Left$(strArgument$, InStr(strArgument$, ":") - 1)
            strBuf2$ = Right(strArgument$, Len(strArgument$) - InStr(strArgument$, ":"))
             
            ''add message to rtbChat
            Call AddChat(strBuf1$, strBuf2$)
            
         Case "SYSMSG":
            ''store color and system message
            strBuf1$ = Left$(strArgument$, InStr(strArgument$, ":") - 1)
            strBuf2$ = Right(strArgument$, Len(strArgument$) - InStr(strArgument$, ":"))
            
            ''print system message
            Call AddSysMessage(strBuf2$, CLng(strBuf1$))
         
         Case "JOIN": ''if someone new joined.
         
                 ''loop through listbox and check if
                 ''someone is using that nick - ONLY if server
                 If optServerClient(0).Value Then
                     For lngIndex& = 0 To lstUsers.ListCount - 1
                         If Trim(LCase$(lstUsers.List(lngIndex&))) = Trim(LCase$(strArgument$)) Then
                            Call sckConnect(Index).SendData("[ERR_NICKINUSE] ")
                            DoEvents
                            Exit Sub
                         End If
                     Next lngIndex&
                 End If
                
             ''print "*** [Nick] has joined the chat."
             Call AddSysMessage(vbCrLf & "*** " & strArgument$ & " has joined the chat.", RGB(15, 181, 0))
         
             ''add nick to list
             Call lstUsers.AddItem(strArgument$)
             strUsers(Index) = strArgument$
             
                ''add all the nicks' of users in the chat, then send them to the
                ''newly connected user. - ONLY if server.
                If optServerClient(0).Value Then
                     For lngIndex& = LBound(strUsers()) To UBound(strUsers())
                        If strUsers(lngIndex&) <> strArgument$ And strUsers(lngIndex&) <> "" Then strBuf1$ = strBuf1$ & Chr$(1) & strUsers(lngIndex)
                     Next lngIndex&
                   
                  ''get rid of the extra chr$(1)
                  If Len(strBuf1$) Then strBuf1$ = Right$(strBuf1$, Len(strBuf1$) - 1)
                  
    
                  ''send data to user.
                  Call sckConnect(Index).SendData("[User] " & strBuf1$)
                  DoEvents
                End If
                   
         Case "LEAVE": ''user has left the chat
                               
                ''get listindex of nick in listbox
                ''then remove it.
                For lngIndex& = 0 To lstUsers.ListCount - 1
                   If lstUsers.List(lngIndex&) = strArgument$ Then
                       Call lstUsers.RemoveItem(lngIndex&)
                       Exit For
                   End If
                Next lngIndex&
                
             ''remove nick from strUsers()
             strUsers(Index) = ""
             
             ''print "*** [Nick] has left the chat."
             Call AddSysMessage(vbCrLf & "*** " & strArgument$ & " has left the chat.", RGB(15, 181, 0))
         
         Case "USER": ''receiving a user that's in the room
                      ''to add to lstUsers.
                 
                  ''parse string so you get the users,
                  ''then add then to the listbox
            lngIndex& = 1
            
            Call lstUsers.AddItem(txtNick.Text)

                  Do
                    lngBuf2& = 1
                    lngBuf& = InStr(lngBuf& + 1, strArgument$, Chr$(1))
                    If lngBuf& = 0 Then lngBuf& = Len(strArgument$): lngBuf2& = 0
                    If lngPos& = Len(strArgument$) Then Exit Do
                    strBuf1$ = Mid$(strArgument$, lngPos& + 1, lngBuf& - lngPos& - lngBuf2&)
                    
                    ''add nick to list and strUsers()
                    Call lstUsers.AddItem(strBuf1$)
                    ReDim Preserve strUsers(lngIndex&) As String
                    strUsers(lngIndex&) = strBuf1$
                    lngIndex& = lngIndex& + 1
                    
                    lngPos& = lngBuf&
                  Loop
         
         Case "KICK": ''someone was kicked by the server
            ''store nick and reason
            strBuf1$ = Left$(strArgument$, InStr(strArgument$, Chr$(1)) - 1)
            strBuf2$ = Mid$(strArgument$, InStr(strArgument$, Chr$(1)) + 1, InStr(InStr(strArgument$, Chr$(1)) + 1, strArgument$, Chr$(1)) - InStr(strArgument$, Chr$(1)) - 1)
            strbuf3$ = Right$(strArgument$, Len(strArgument$) - InStr(InStr(strArgument$, Chr$(1)) + 1, strArgument$, Chr$(1)))
           
            ''print "*** user has been kicked by [server] (reason)"
            Call AddSysMessage(vbCrLf & "*** " & strBuf1$ & " was kicked by " & strBuf2$ & " (" & strbuf3$ & ")", RGB(15, 181, 0))
         
              ''delete nick from list
              For lngIndex& = 0 To lstUsers.ListCount - 1
                 If lstUsers.List(lngIndex&) = strBuf1$ Then
                    Call lstUsers.RemoveItem(lngIndex&)
                    Exit For
                 End If
              Next lngIndex&
              
              ''delete nick from strUsers
              For lngIndex& = 0 To UBound(strUsers())
                 If strUsers(lngIndex&) = strBuf1$ Then
                    strUsers(lngIndex&) = ""
                    Exit For
                 End If
              Next lngIndex&
              
         Case "ERR_NICKINUSE": ''nick is being used.
              ''prompt for new nick
              Do
                strBuf1$ = InputBox("The nickname '" & txtNick.Text & "' is currently in use by someone else in the chat. Please choose another nickname: ", "New nick")
              Loop Until Trim(strBuf1$) <> ""
              
            txtNick.Text = strBuf1$
            
            ''change index 0 of strUsers to the person's nick.
            strUsers(0) = strBuf1$
            
            ''resend data and exit
            Call sckConnect(0).SendData("[Join] " & strBuf1$)
            Exit Sub
            
      End Select
  
  ''send data to everyone in the chat - ONLY if server.
  If optServerClient(0).Value = True Then
      For lngIndex& = 1 To sckConnect().Count - 1
          If sckConnect(lngIndex&).State = sckConnected And lngIndex& <> Index Then
              Call sckConnect(lngIndex&).SendData(strData$)
              DoEvents
          End If
      Next lngIndex&
  End If

End Sub
Sub AddSysMessage(strText As String, Optional lngColor As Long = vbRed)

  'A system message is something like '*** Disconnected'

  rtbChat.SelStart = Len(rtbChat.Text)
  rtbChat.SelLength = 0
  rtbChat.SelText = strText$
  rtbChat.SelStart = Len(rtbChat.Text) - Len(strText$)
  rtbChat.SelLength = Len(strText$)
  rtbChat.SelColor = lngColor&
  rtbChat.SelBold = False
  rtbChat.SelFontName = "Courier New"
  rtbChat.SelFontSize = 10
  rtbChat.SelStart = Len(rtbChat.Text)
  rtbChat.SelLength = 0


End Sub


Private Sub chkBold_Click()

   'toggle bold
   rtbText.SelBold = Not rtbText.SelBold
   rtbText.SetFocus
   
End Sub

Private Sub chkItalic_Click()

   'toggle italic
   rtbText.SelItalic = Not rtbText.SelItalic
   rtbText.SetFocus
   
End Sub

Private Sub chkUnderline_Click()

   'toggle underline
   rtbText.SelUnderline = Not rtbText.SelUnderline
   rtbText.SetFocus
   
End Sub

Private Sub cmbFonts_Click()
  
  On Error Resume Next

  'set the font
  rtbText.SelFontName = cmbFonts.List(cmbFonts.ListIndex)
  rtbText.SetFocus

End Sub

Private Sub cmdColors_Click()

On Error GoTo ErrorHandler

dlgColors.CancelError = True
dlgColors.ShowColor
rtbText.SelColor = dlgColors.Color
rtbText.SetFocus


ErrorHandler: 'user click 'Cancel'
End Sub

Private Sub cmdSend_Click()
    
  Dim lngIndex As Long
  
    ''check if connected to someone
    If sckConnect(0).State = sckClosed Then
       MsgBox "Error: You must be connected to someone.", vbCritical, "Error"
       Exit Sub
    End If
  
  ''send to rtbChat
  Call AddChat(txtNick.Text, rtbText.TextRTF)

  ''send text to server to process. - ONLY if guest
  If sckConnect(0).State = sckConnected And optServerClient(1).Value Then Call sckConnect(0).SendData("[Message] " & txtNick.Text & ":" & rtbText.TextRTF)

     ''send data to everyone in the chat - ONLY if server.
     If optServerClient(0).Value = True Then
         For lngIndex& = 1 To sckConnect().Count - 1
            If sckConnect(lngIndex&).State = sckConnected Then
              Call sckConnect(lngIndex&).SendData("[Message] " & txtNick.Text & ":" & rtbText.TextRTF)
              DoEvents ''tell processor to finish sending
                       ''data before proceeding
            End If
         Next lngIndex&
     End If


  ''clear textbox
  rtbText.Text = ""


End Sub

Private Sub Command1_Click()
   For i = 0 To UBound(strUsers())
      MsgBox strUsers(i)
   Next i
End Sub

Private Sub Form_Load()
 
 Dim intBuffer As Integer, strFont As String
 
   'load printer fonts to combobox
   If Dir$(App.Path & "\fonts.dat") = "" Then
        'font file doesnt exist. Create it.
        Open App.Path & "\fonts.dat" For Output As #1
             For intBuffer% = 0 To Printer.FontCount - 1
                Call cmbFonts.AddItem(Printer.Fonts(intBuffer%))
                Print #1, Printer.Fonts(intBuffer%)
             Next intBuffer%
        Close #1
   Else
        'load fonts from file
        Open App.Path & "\fonts.dat" For Input As #1
             While Not EOF(1)
                Input #1, strFont$
                Call cmbFonts.AddItem(strFont$)
             Wend
        Close #1
   End If
   
 cmbFonts.ListIndex = 0
 ''cmbFonts.Sorted = True 'Alphabetize list


  'set combobox to "Arial"
  For intBuffer% = 0 To cmbFonts.ListCount - 1
    If cmbFonts.List(intBuffer%) = "Arial" Then cmbFonts.ListIndex = intBuffer%: Exit For
  Next intBuffer%
  

  'set rtbText's font-styles
  rtbText.SelBold = False
  rtbText.SelUnderline = False
  rtbText.SelItalic = False
  rtbText.SelColor = vbBlack
  rtbText.SelFontName = cmbFonts.List(cmbFonts.ListIndex)
  rtbText.SelFontSize = 10

End Sub

Private Sub Form_Resize()

  ''resize the controls on the form.
  
    ''resize by height
    
  rtbChat.Height = Me.Height - 2170
  lstUsers.Height = Me.Height - 2170
  
  ''set form height so it fits height of listbox
  If rtbChat.Height <> lstUsers.Height And Me.WindowState = vbNormal Then Me.Height = lstUsers.Height + 2170
    
  txtIP.Top = Me.Height - 705
  txtPort.Top = Me.Height - 705
  lblIP.Top = Me.Height - 675
  lblPort.Top = Me.Height - 675
  
  lneSep(0).Y1 = Me.Height - 780
  lneSep(0).Y2 = Me.Height - 780
  lneSep(1).Y1 = Me.Height - 790
  lneSep(1).Y2 = Me.Height - 790
  lneSep2(0).Y1 = Me.Height - 1185
  lneSep2(0).Y2 = Me.Height - 765
  lneSep2(1).Y1 = Me.Height - 1185
  lneSep2(1).Y2 = Me.Height - 775
  lneSep3(0).Y1 = Me.Height - 1185
  lneSep3(0).Y2 = Me.Height - 765
  lneSep3(1).Y1 = Me.Height - 1185
  lneSep3(1).Y2 = Me.Height - 775
  
  shpRed.Top = Me.Height - 705
  shpGreen.Top = Me.Height - 705
  
  cmdConnect.Top = Me.Height - 1170
  
  optServerClient(0).Top = Me.Height - 1070
  optServerClient(1).Top = Me.Height - 1070
  
  lblNick.Top = Me.Height - 1095
  txtNick.Top = Me.Height - 1125
  
  frmeSep.Top = Me.Height - 1305
  
  rtbText.Top = Me.Height - 1590
  cmdSend.Top = Me.Height - 1575
  
  cmbFonts.Top = Me.Height - 1890
  cmdColors.Top = Me.Height - 1890
  chkBold.Top = Me.Height - 1890
  chkItalic.Top = Me.Height - 1890
  chkUnderline.Top = Me.Height - 1890
  
  frmeChatWindow.Height = Me.Height - 1920
  
  
  ''do width
  frmeSep.Width = Me.Width - 115
  cmdConnect.Left = Me.Width - 1090
  
  lneSep(0).X2 = Me.Width - 1125
  lneSep(1).X2 = Me.Width - 1125
  lneSep2(0).X1 = Me.Width - 2910
  lneSep2(0).X2 = Me.Width - 2910
  lneSep2(1).X1 = Me.Width - 2920
  lneSep2(1).X2 = Me.Width - 2920
  lneSep3(0).X1 = Me.Width - 1140
  lneSep3(0).X2 = Me.Width - 1140
  lneSep3(1).X1 = Me.Width - 1150
  lneSep3(1).X2 = Me.Width - 1150
  
  optServerClient(1).Left = Me.Width - 1975
  optServerClient(0).Left = Me.Width - 2815
  
  txtNick.Width = Me.Width - 3400
  
  cmdSend.Left = Me.Width - 750
  rtbText.Width = Me.Width - 750
  
  shpGreen.Left = Me.Width - 525
  shpRed.Left = Me.Width - 1005
  
  lblIP.Left = (Me.Width - 1160) / 2
  txtIP.Left = (Me.Width - 1160) / 2 + 250
  txtIP.Width = (Me.Width - 1130) - ((Me.Width - 1160) / 2 + 250)
  txtPort.Width = (Me.Width - 1160) / 2 - 410
  
  frmeChatWindow.Width = Me.Width - 115
  lstUsers.Left = Me.Width - 1275
  rtbChat.Width = Me.Width - 1350
  
  chkUnderline.Left = Me.Width - 1965
  chkItalic.Left = Me.Width - 1965 - 315
  chkBold.Left = Me.Width - 1965 - 315 * 2
  cmdColors.Left = Me.Width - 3045
  
  cmbFonts.Width = Me.Width - 3045 - 140
  
End Sub

Private Sub Form_Unload(Cancel As Integer)

   ''if connected tell everyone in the chat you left.
   If sckConnect(0).State = sckConnected Then Call cmdConnect_Click
     
End Sub

Private Sub lstUsers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  ''If not right-clicking then exit
  If Button <> 2 Then Exit Sub

    ''Enable menu if connected
    If cmdConnect.Caption = "Disconnect" Then
       mnuKickUser(0).Enabled = True
       mnuKickUser(1).Enabled = True
    Else
       mnuKickUser(0).Enabled = False
       mnuKickUser(1).Enabled = False
    End If
  
  ''Pop up the menu
  Call Me.PopupMenu(mnuList, vbAlignNone)
  
  
End Sub

Private Sub mnuKickUser_Click(Index As Integer)
  
  Dim lngIndex As Long, lngWinsock As Long, strReason As String

    ''If client then deny access to kick
    If optServerClient(1).Value Then
       Call AddSysMessage(vbCrLf & "*** Permission Denied")
       Exit Sub
    End If
  
    ''Find winsock for user and close that connection
    For lngIndex& = 0 To UBound(strUsers())
         If strUsers(lngIndex&) = lstUsers.List(lstUsers.ListIndex) Then
              ''Get reason
              If Index Then strReason$ = InputBox("Please enter a reason why you want to kick '" & lstUsers.List(lstUsers.ListIndex) & "':", "Kick User", "<none>")
              If strReason$ = "" Or strReason$ = "<none>" Then strReason$ = txtNick.Text
              
              Call sckConnect(lngIndex&).SendData("[SysMsg] " & RGB(15, 181, 0) & ":" & vbCrLf & "*** You were kicked by " & txtNick.Text & " (" & strReason$ & ")")
              DoEvents
              Call sckConnect(lngIndex&).Close
              
              ''clear entry in strUsers
              strUsers(lngIndex&) = ""
              
                 ''Tell everyone that person was kicked
                 For lngWinsock& = 1 To sckConnect().UBound
                    If sckConnect(lngWinsock&).State = sckConnected Then
                       Call sckConnect(lngWinsock&).SendData("[Kick] " & lstUsers.List(lstUsers.ListIndex) & Chr$(1) & txtNick.Text & Chr$(1) & strReason$)
                       DoEvents
                    End If
                 Next lngWinsock&
                 
             ''write that nick was kicked.
              Call AddSysMessage(vbCrLf & "*** " & lstUsers.List(lstUsers.ListIndex) & " was kicked by " & txtNick.Text & " (" & strReason$ & ")", RGB(15, 181, 0))
              
              ''remove him from list
              Call lstUsers.RemoveItem(lstUsers.ListIndex)
                
          End If
    Next lngIndex&

End Sub

Private Sub rtbText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'check if text has a certain font-style and set checkboxes
  'to the current font-style

   If rtbText.SelBold = True Then
      chkBold.Value = vbChecked
   Else
      chkBold.Value = vbUnchecked
   End If
   
   If rtbText.SelItalic = True Then
      chkItalic.Value = vbChecked
   Else
      chkItalic.Value = vbUnchecked
   End If
   
   If rtbText.SelUnderline = True Then
      chkUnderline.Value = vbChecked
   Else
      chkUnderline.Value = vbUnchecked
   End If

End Sub

Private Sub sckConnect_Close(Index As Integer)

   'if user is guest then display "Disconnected".
   If optServerClient(1).Value = True Then
      cmdConnect_Click
   End If

End Sub

Private Sub sckConnect_Connect(Index As Integer)

        ''turn from red to green light
        shpRed.BackColor = vbDarkRed
        shpRed.BorderColor = vbDarkRed
        shpGreen.BackColor = vbGreen
        shpGreen.BorderColor = vbGreen
        
        
           ''write "*** Connected". ONLY if it's the first user to connect
           ''and your the server OR you just a client
           If (sckConnect().UBound = 1 And optServerClient(0).Value) Or optServerClient(1).Value Then
              Call AddSysMessage(vbCrLf & "*** Connected")
                
                ''if guest, then tell everyone you
                ''have joined the chat.
                If optServerClient(1).Value Then Call sckConnect(0).SendData("[Join] " & txtNick.Text)
           End If
        
End Sub

Private Sub sckConnect_ConnectionRequest(Index As Integer, ByVal requestID As Long)
   
   Dim lngIndex As Long, blnFlag As Boolean
   
      ''loop through winsocks and see if there is a
      ''winsock that is not in use.
      For lngIndex& = 1 To sckConnect().UBound
         If sckConnect(lngIndex&).State = sckClosed Then
             blnFlag = True
             Exit For
         End If
      Next lngIndex&
      
      ''if all winsocks is in use then assign lngIndex to
      ''UBound + 1, and load a new winsock.
      If blnFlag = False Then
         lngIndex& = sckConnect().UBound + 1
         Load sckConnect(lngIndex&)
         ReDim Preserve strUsers(lngIndex&) As String
      End If
   
   ''accept connection
   Call sckConnect(lngIndex&).Accept(requestID&)
   
   Call sckConnect_Connect(Index) 'raise connect event

End Sub

Private Sub sckConnect_DataArrival(Index As Integer, ByVal bytesTotal As Long)

   Dim strData As String
   
   ''get data
   Call sckConnect(Index).GetData(strData$, vbString)
         
   ''parse data
   Call ParseData(strData$, Index)
   
End Sub

Private Sub cmdConnect_Click()

  Dim strIP As String, lngIndex As Long, strNewLine As String

  
    If cmdConnect.Caption = "Connect" Or cmdConnect.Caption = "Listen" Then
       txtPort.Enabled = False
       txtNick.Enabled = False
       optServerClient(0).Enabled = False
       optServerClient(1).Enabled = False
       cmdConnect.Caption = "Disconnect"
       
       ''set strUsers index 0 the nick of the person.
       ReDim strUsers(0) As String
       strUsers(0) = txtNick.Text
    Else
       txtPort.Enabled = True
       txtNick.Enabled = True
       optServerClient(0).Enabled = True
       optServerClient(1).Enabled = True
       
          If optServerClient(0).Value = True Then
             cmdConnect.Caption = "Listen"
          Else
             cmdConnect.Caption = "Connect"
          End If
          
          ''if guest AND your connected then tell server you left the chat
          If optServerClient(1).Value And sckConnect(0).State = sckConnected Then
             Call sckConnect(0).SendData("[Leave] " & txtNick.Text)
             DoEvents
          End If
        
           ''loop through all winsocks and close all
           ''connections.
           For lngIndex& = 0 To sckConnect().UBound
              If sckConnect(lngIndex&).State <> sckClosed Then Call sckConnect(lngIndex&).Close
           Next lngIndex&
        
        ''turn from green to red light
        shpRed.BackColor = vbRed
        shpRed.BorderColor = vbRed
        shpGreen.BackColor = vbDarkGreen
        shpGreen.BorderColor = vbDarkGreen
        
        ''write "*** Disconnected".
        Call AddSysMessage(vbCrLf & "*** Disconnected")
        
        ''clear user list
        Call lstUsers.Clear
        
        ''clear user array
        Erase strUsers()
        Exit Sub
    End If
  
 ''if there's something in rtbChat then add
 ''a new line to strNewLine - which is to be
 ''sent to rtbChat.
 strNewLine$ = vbCrLf
 If rtbChat.Text = "" Then strNewLine$ = ""
  
    Select Case optServerClient(0).Value
        
        Case True:  ''Host
        
           ''listen for connections
           sckConnect(0).LocalPort = CLng(txtPort.Text)
           sckConnect(0).Listen
           
           ''write "*** Waiting for Connection.", and add a new line
           ''if there's something in rtbChat.
           Call AddSysMessage(strNewLine$ & "*** Waiting for connection...")
        
           ''add nick to list
           Call lstUsers.AddItem(txtNick.Text)
        
        Case False: ''Guest
        
           ''try to connect
           strIP$ = txtIP.Text
           If LCase$(strIP$) = "localhost" Then strIP$ = sckConnect(0).LocalIP
           Call sckConnect(0).Connect(strIP$, txtPort.Text)
           
           ''write "*** Connecting", and add a new line
           ''if there's something in rtbChat.
           Call AddSysMessage(strNewLine$ & "*** Connecting...")
           
    End Select
            


End Sub
Private Sub optServerClient_Click(Index As Integer)

    Select Case Index
       Case 0: 'Server
          txtIP.BackColor = &H8000000F  'Grey
          txtIP.Locked = True
          cmdConnect.Caption = "Listen"
          txtIP.Text = sckConnect(0).LocalIP
       Case 1: 'Client
          txtIP.Text = "localhost"
          txtIP.BackColor = vbWhite
          txtIP.Locked = False
          cmdConnect.Caption = "Connect"
          
    End Select

End Sub


Private Sub rtbText_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 13 Then 'If user pressed 'Enter'
      cmdSend_Click 'click 'Send' button
      KeyAscii = 0 'Make sure it doesnt write enter to rtbText
   End If
    
End Sub



