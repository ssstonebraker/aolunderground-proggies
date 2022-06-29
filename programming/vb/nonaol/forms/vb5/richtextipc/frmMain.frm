VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rich Text Chat"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   4440
      TabIndex        =   19
      Top             =   1920
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock sckListen 
      Left            =   4440
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckConnect 
      Left            =   3960
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtbChat 
      Height          =   1660
      Left            =   80
      TabIndex        =   18
      Top             =   190
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2910
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":0000
   End
   Begin MSComDlg.CommonDialog dlgColors 
      Left            =   3480
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdColors 
      Height          =   320
      Left            =   2040
      Picture         =   "frmMain.frx":00EF
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1930
      Width           =   315
   End
   Begin RichTextLib.RichTextBox rtbText 
      Height          =   360
      Left            =   -5
      TabIndex        =   16
      Top             =   2240
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   635
      _Version        =   393217
      TextRTF         =   $"frmMain.frx":0431
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
      Left            =   3120
      Picture         =   "frmMain.frx":0520
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1930
      Width           =   315
   End
   Begin VB.CheckBox chkItalic 
      Height          =   320
      Left            =   2815
      Picture         =   "frmMain.frx":0862
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1930
      Width           =   315
   End
   Begin VB.ComboBox cmbFonts 
      Height          =   315
      Left            =   10
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1930
      Width           =   1935
   End
   Begin VB.CheckBox chkBold 
      Height          =   320
      Left            =   2520
      Picture         =   "frmMain.frx":0BA4
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1930
      UseMaskColor    =   -1  'True
      Width           =   315
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   405
      Left            =   3390
      TabIndex        =   11
      Top             =   2655
      Width           =   975
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   1960
      TabIndex        =   10
      Text            =   "localhost"
      Top             =   3120
      Width           =   1415
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   440
      TabIndex        =   8
      Text            =   "400"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.OptionButton optHostGuest 
      Caption         =   "Guest"
      Height          =   195
      Index           =   1
      Left            =   2520
      TabIndex        =   6
      Top             =   2760
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.OptionButton optHostGuest 
      Caption         =   "Host"
      Height          =   195
      Index           =   0
      Left            =   1800
      TabIndex        =   5
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox txtNick 
      Height          =   285
      Left            =   440
      TabIndex        =   4
      Text            =   "NickName"
      Top             =   2700
      Width           =   1215
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   335
      Left            =   3720
      TabIndex        =   2
      Top             =   2245
      Width           =   625
   End
   Begin VB.Frame frmeSep 
      Height          =   135
      Left            =   0
      TabIndex        =   1
      Top             =   2520
      Width           =   4385
   End
   Begin VB.Frame frmeChatWindow 
      Caption         =   "Chat Window"
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4365
   End
   Begin VB.Shape shpGreen 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008000&
      Height          =   255
      Left            =   3960
      Shape           =   3  'Circle
      Top             =   3120
      Width           =   255
   End
   Begin VB.Shape shpRed 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      FillColor       =   &H00000080&
      Height          =   255
      Left            =   3480
      Shape           =   3  'Circle
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label lblIP 
      BackStyle       =   0  'Transparent
      Caption         =   "IP:"
      Height          =   255
      Left            =   1710
      TabIndex        =   9
      Top             =   3150
      Width           =   255
   End
   Begin VB.Label lblPort 
      BackStyle       =   0  'Transparent
      Caption         =   "Port:"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   3150
      Width           =   375
   End
   Begin VB.Line lneSep3 
      BorderColor     =   &H80000003&
      Index           =   1
      X1              =   3345
      X2              =   3345
      Y1              =   2640
      Y2              =   3050
   End
   Begin VB.Line lneSep3 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   3360
      X2              =   3360
      Y1              =   2640
      Y2              =   3060
   End
   Begin VB.Line lneSep2 
      BorderColor     =   &H80000003&
      Index           =   1
      X1              =   1680
      X2              =   1680
      Y1              =   3025
      Y2              =   2625
   End
   Begin VB.Line lneSep 
      BorderColor     =   &H80000003&
      Index           =   1
      X1              =   0
      X2              =   3360
      Y1              =   3030
      Y2              =   3030
   End
   Begin VB.Line lneSep2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   1695
      X2              =   1695
      Y1              =   3045
      Y2              =   2640
   End
   Begin VB.Line lneSep 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   0
      X2              =   3360
      Y1              =   3045
      Y2              =   3045
   End
   Begin VB.Label lblNick 
      BackStyle       =   0  'Transparent
      Caption         =   "Nick: "
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   2730
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************
'*          Rich Text IP Chat by Joseph Huntley           *
'*               joseph_huntley@email.com                 *
'*                http://joseph.vr9.com                   *
'*                                                        *
'*  Made:  October 6, 1999                                *
'*  Level: Intermediate/Advanced                          *
'**********************************************************
'* Notes: This is an expanded version of my original IP   *
'*        chat example. The only difference is that this  *
'*        version uses a richtextbox for colorful chat.   *
'**********************************************************

Private Const vbDarkRed = &H80&
Private Const vbDarkGreen = &H8000&


Private Sub chkBold_Click()

   'toggle bold
   If chkBold.Value = vbChecked Then
      rtbText.SelBold = True
   Else
      rtbText.SelBold = False
   End If
   
 rtbText.SetFocus
   
End Sub

Private Sub chkItalic_Click()

   'toggle italic
   If chkItalic.Value = vbChecked Then
      rtbText.SelItalic = True
   Else
      rtbText.SelItalic = False
   End If
   
 rtbText.SetFocus
   
End Sub

Private Sub chkUnderline_Click()

   'toggle underline
   If chkUnderline.Value = vbChecked Then
      rtbText.SelUnderline = True
   Else
      rtbText.SelUnderline = False
   End If
   
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
  
  Dim strRTF As String, strFontName As String, lngFontColor As Long, lngFontSize As Long
  Dim blnBold As Boolean, blnUnderline As Boolean, blnItalic As Boolean
  
    ''check if connected to someone
    If sckConnect.State = sckClosed Then
       MsgBox "You must be connected to someone."
       Exit Sub
    End If
   
  ''get font-styles so we can reset them later
  blnBold = rtbText.SelBold
  blnUnderline = rtbText.SelUnderline
  blnItalic = rtbText.SelItalic
  lngFontColor& = rtbText.SelColor
  lngFontSize& = rtbText.SelFontSize
  strFontName$ = rtbText.SelFontName
  
  ''format text w/ nick and assign rtf to strRTF$
  rtbText.SelStart = 0
  rtbText.SelLength = 0
  rtbText.SelText = vbCrLf & txtNick.Text & ":" & vbTab
  rtbText.SelStart = 0
  rtbText.SelLength = Len(txtNick.Text) + 4 '4 = Length of vbCrLf + ':' + vbTab
  rtbText.SelColor = vbBlue
  rtbText.SelFontSize = 8
  rtbText.SelFontName = "Arial"
  rtbText.SelBold = True
  rtbText.SelUnderline = False
  rtbText.SelItalic = False
  rtbText.SelStart = 0
  rtbText.SelLength = 0
  
  strRTF$ = rtbText.TextRTF
  
  ''clear textbox
  rtbText.Text = ""
   
  ''reset font-styles
  rtbText.SelBold = blnBold
  rtbText.SelUnderline = blnUnderline
  rtbText.SelItalic = blnItalic
  rtbText.SelColor = lngFontColor&
  rtbText.SelFontSize = lngFontSize&
  rtbText.SelFontName = strFontName$
   
  ''show bottom half of textbox
  rtbChat.SelStart = Len(rtbChat.Text)
  rtbChat.SelLength = 0
  
  'print text in our rtbChat
  rtbChat.SelRTF = strRTF$
  
  
  'scroll rtbChat down
  rtbChat.SelStart = Len(rtbChat.Text)
  rtbChat.SelLength = 0
  
  'set focus
  rtbText.SetFocus
  
  ''Send text to other person
  Call sckConnect.SendData(strRTF$)
 
  
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

Private Sub rtbText_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
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

Private Sub sckConnect_Close()
   cmdConnect_Click 'Reset everything
End Sub

Private Sub sckConnect_Connect()

        ''turn from red to green light
        shpRed.BackColor = vbDarkRed
        shpRed.BorderColor = vbDarkRed
        shpGreen.BackColor = vbGreen
        shpGreen.BorderColor = vbGreen
        
        ''write "*** Connected".
        rtbChat.SelStart = Len(rtbChat.Text)
        rtbChat.SelLength = 0
        rtbChat.SelText = "*** Connected"
        rtbChat.SelStart = Len(rtbChat.Text)
        rtbChat.SelLength = Len("*** Connected")
        rtbChat.SelColor = vbRed
        rtbChat.SelBold = False
        rtbChat.SelFontName = "Courier New"
        rtbChat.SelFontSize = 10
        rtbChat.SelStart = Len(rtbChat.Text)
        rtbChat.SelLength = 0
        
End Sub

Private Sub sckConnect_ConnectionRequest(ByVal requestID As Long)
   ''check if connected to something else. If so, close that connection
   If sckConnect.State <> sckClosed Then sckConnect.Close
   
   'accept the connection
   Call sckConnect.Accept(requestID&)
   
   sckConnect_Connect 'fire 'Connect' event
End Sub

Private Sub sckConnect_DataArrival(ByVal bytesTotal As Long)
   Dim strData As String, strBuf As String
   Dim strNewNick As String
   
   'get RTF string to add
   Call sckConnect.GetData(strData$, vbString)
         
   'set selpostioning
   rtbChat.SelStart = Len(rtbChat.Text)
   rtbChat.SelLength = 0
   
   'print text
   rtbChat.SelRTF = strData$
   
   'make textbox scroll
   rtbChat.SelStart = Len(rtbChat.Text)
   rtbChat.SelLength = 0
   
End Sub

Private Sub cmdConnect_Click()

  Dim strIP As String

  
    If cmdConnect.Caption = "Connect" Or cmdConnect.Caption = "Listen" Then
       txtPort.Enabled = False
       txtNick.Enabled = False
       optHostGuest(0).Enabled = False
       optHostGuest(1).Enabled = False
       cmdConnect.Caption = "Disconnect"
    Else
       txtPort.Enabled = True
       txtNick.Enabled = True
       optHostGuest(0).Enabled = True
       optHostGuest(1).Enabled = True
       
          If optHostGuest(0).Value = True Then
             cmdConnect.Caption = "Listen"
          Else
             cmdConnect.Caption = "Connect"
          End If
          
        sckConnect.Close
        
        ''turn from green to red light
        shpRed.BackColor = vbRed
        shpRed.BorderColor = vbRed
        shpGreen.BackColor = vbDarkGreen
        shpGreen.BorderColor = vbDarkGreen
        ''write "*** Disconnected".
        rtbChat.SelStart = Len(rtbChat.Text)
        rtbChat.SelLength = 0
        rtbChat.SelText = vbCrLf & "*** Disconnected"
        rtbChat.SelStart = Len(rtbChat.Text) - Len(vbCrLf & "*** Disconnected")
        rtbChat.SelLength = Len(vbCrLf & "*** Disconnected")
        rtbChat.SelColor = vbRed
        rtbChat.SelBold = False
        rtbChat.SelFontName = "Courier New"
        rtbChat.SelFontSize = 10
        rtbChat.SelStart = Len(rtbChat.Text)
        rtbChat.SelLength = 0
        Exit Sub
    End If
  
    Select Case optHostGuest(0).Value
        Case True:  'Host
           ''listen for connections
           sckConnect.LocalPort = CLng(txtPort.Text)
           sckConnect.Listen
           
           ''write "*** Waiting for Connection."
           rtbChat.Text = ""
           rtbChat.SelStart = 0
           rtbChat.SelLength = 0
           rtbChat.SelText = "*** Waiting for connection..." & vbCrLf
           rtbChat.SelStart = 0
           rtbChat.SelLength = Len("*** Waiting for connection..." & vbCrLf)
           rtbChat.SelColor = vbRed
           rtbChat.SelBold = False
           rtbChat.SelFontName = "Courier New"
           rtbChat.SelFontSize = 10
           rtbChat.SelStart = Len(rtbChat.Text)
           rtbChat.SelLength = 0
        Case False: 'Guest
           ''try to connect
           strIP$ = txtIP.Text
           If LCase$(strIP$) = "localhost" Then strIP$ = sckConnect.LocalIP
           sckConnect.Connect txtIP.Text, txtPort.Text
           
           ''write "*** Connecting".
           rtbChat.Text = ""
           rtbChat.SelStart = 0
           rtbChat.SelLength = 0
           rtbChat.SelText = "*** Connecting..." & vbCrLf
           rtbChat.SelStart = 0
           rtbChat.SelLength = Len("*** Connecting..." & vbCrLf)
           rtbChat.SelColor = vbRed
           rtbChat.SelBold = False
           rtbChat.SelFontName = "Courier New"
           rtbChat.SelFontSize = 10
           rtbChat.SelStart = Len(rtbChat.Text)
           rtbChat.SelLength = 0

    End Select
            


End Sub
Private Sub optHostGuest_Click(Index As Integer)

    Select Case Index
       Case 0: 'Host
          txtIP.Text = sckConnect.LocalIP
          txtIP.BackColor = &H8000000F  'Grey
          txtIP.Locked = True
          cmdConnect.Caption = "Listen"
       Case 1: 'Guest
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
