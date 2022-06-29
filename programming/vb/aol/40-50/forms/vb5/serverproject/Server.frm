VERSION 5.00
Object = "{30ACFC93-545D-11D2-A11E-20AE06C10000}#1.0#0"; "VB5CHAT2.OCX"
Begin VB.Form Server 
   BackColor       =   &H00000040&
   BorderStyle     =   0  'None
   Caption         =   "X-Treme Server '98"
   ClientHeight    =   2925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5655
   ControlBox      =   0   'False
   FillColor       =   &H00808080&
   Icon            =   "Server.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin VB5Chat2.Chat Chat1 
      Left            =   120
      Top             =   3720
      _ExtentX        =   3969
      _ExtentY        =   2170
   End
   Begin VB.Timer Timer7 
      Interval        =   1
      Left            =   6120
      Top             =   1800
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5640
      Top             =   1200
   End
   Begin VB.ListBox List7 
      Height          =   840
      Left            =   4920
      TabIndex        =   30
      Top             =   3360
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000008&
      Height          =   255
      Left            =   4080
      Picture         =   "Server.frx":030A
      ScaleHeight     =   195
      ScaleWidth      =   1395
      TabIndex        =   24
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4920
      TabIndex        =   23
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Add SN's"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4080
      TabIndex        =   22
      Top             =   1440
      Width           =   855
   End
   Begin VB.ListBox List6 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   480
      ItemData        =   "Server.frx":08D5
      Left            =   4080
      List            =   "Server.frx":08D7
      TabIndex        =   21
      Top             =   600
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   1680
      ScaleHeight     =   195
      ScaleWidth      =   3795
      TabIndex        =   20
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Timer Timer5 
      Interval        =   1000
      Left            =   1560
      Top             =   3000
   End
   Begin VB.Timer Timer3 
      Interval        =   500
      Left            =   840
      Top             =   3000
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   480
      Top             =   3000
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   3000
   End
   Begin VB.ListBox List5 
      Height          =   255
      Left            =   3480
      TabIndex        =   18
      Top             =   3000
      Width           =   1335
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   480
      Left            =   240
      TabIndex        =   17
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton Command7 
      Caption         =   "X"
      Height          =   195
      Left            =   3720
      TabIndex        =   12
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Misc"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3240
      TabIndex        =   11
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Prefs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2640
      TabIndex        =   10
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Refresh Mails"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1440
      TabIndex        =   9
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   720
      TabIndex        =   8
      Tag             =   "0"
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   1920
      Width           =   615
   End
   Begin VB.ListBox List2 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   1680
      TabIndex        =   4
      Top             =   2280
      Width           =   3855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Text            =   "X-Treme Server '99 Rules"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "  Unknown"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.ListBox List4 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   480
      Left            =   2400
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.ListBox List3 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   480
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
   Begin VB.Image Image11 
      Height          =   255
      Left            =   0
      Picture         =   "Server.frx":08D9
      Top             =   360
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Picture         =   "Server.frx":0D82
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Label Label13 
      Caption         =   "0"
      Height          =   255
      Left            =   4920
      TabIndex        =   32
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label14 
      Caption         =   "1"
      Height          =   255
      Left            =   4560
      TabIndex        =   31
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label label11 
      BackColor       =   &H8000000E&
      Caption         =   "0"
      Height          =   255
      Left            =   5280
      TabIndex        =   29
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1080
      TabIndex        =   28
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " - •´)•–  X-Treme Server '99 By Tito •´)•–"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   5655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   5160
      TabIndex        =   26
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   5160
      TabIndex        =   25
      Top             =   360
      Width           =   375
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000C&
      X1              =   4080
      X2              =   5520
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Image Image10 
      Height          =   255
      Left            =   120
      Picture         =   "Server.frx":1395
      Top             =   2280
      Width           =   1500
   End
   Begin VB.Image Image9 
      Height          =   255
      Left            =   4080
      Picture         =   "Server.frx":18F3
      Top             =   360
      Width           =   1500
   End
   Begin VB.Image Image8 
      Height          =   255
      Left            =   4080
      Picture         =   "Server.frx":1E68
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   4080
      X2              =   5520
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Image Image7 
      Height          =   240
      Left            =   1080
      Picture         =   "Server.frx":23AB
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   3600
      TabIndex        =   19
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   3480
      TabIndex        =   16
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   2640
      TabIndex        =   15
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   1440
      TabIndex        =   14
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Off"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3480
      TabIndex        =   13
      Top             =   1200
      Width           =   495
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   3960
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Status 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ready"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   2640
      Width           =   5655
   End
   Begin VB.Image Image6 
      Height          =   255
      Left            =   2760
      Picture         =   "Server.frx":2A66
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Image Image5 
      Height          =   255
      Left            =   1800
      Picture         =   "Server.frx":2EC3
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   240
      Picture         =   "Server.frx":33CB
      Top             =   1440
      Width           =   1500
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   120
      Picture         =   "Server.frx":3874
      Top             =   1680
      Width           =   1500
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   2280
      Picture         =   "Server.frx":3DB5
      Top             =   360
      Width           =   1500
   End
End
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Dim WithEvents Chat1 As CHAT_CLASS
Private Sub sChat1_newMsg(What_Said As String)
Dim sNum As String
If MChatBot = True Then
    frmChat.Text1 = frmChat.Text1 & Screen_Name & " :    " & What_Said & Chr(13) & Chr(10)
    frmChat.Text1.SelStart = Len(frmChat.Text1)
End If
If ServerBot = False Then Exit Sub
For DoThis = 0 To List6.ListCount - 1
      If InStr(LCase$(Screen_Name), LCase$(List6.List(DoThis))) Then Exit Sub
Next DoThis
If InStr(UCase$(What_Said), "/" & UCase$(AOLUserSN)) Then
'If UCase$(What_Said) Like "/" & UCase$(AOLUserSN) & "*" Then
    Srv$ = UCase$(Mid$(What_Said, Len(AOLUserSN) + 3))
    If Srv$ Like "FIND *" Then
        fPhr$ = UCase$(Mid$(What_Said, Len(AOLUserSN) + 8))
        For DoThis = 0 To List5.ListCount - 1
           If UCase$(List5.List(DoThis)) Like UCase$(Screen_Name) & ";" & fPhr$ Then List5.RemoveItem a
        Next DoThis
        List5.AddItem (Screen_Name & ";" & fPhr$)
        Label3.Caption = List5.ListCount
    End If
    If Srv$ Like "SEND *" Then
       sNum$ = UCase$(Trim(Mid$(What_Said, Len(AOLUserSN()) + 8)))
       If sNum$ = "LIST" Then
           For DoIt = 0 To List4.ListCount - 1
               If UCase$(List4.List(DoIt)) Like UCase$(Screen_Name) Then Exit Sub
           Next DoIt
           List4.AddItem (Screen_Name)
           Label16.Caption = List4.ListCount '& " Waiting For List"
           Exit Sub
        End If
        If InStr(sNum$, "-") Then
        'If IsNumeric(sNum$) = True Then
           For DoThis = 0 To List3.ListCount - 1
              If UCase$(List3.List(DoThis)) Like UCase$(Screen_Name) & " - SEND " & sNum$ Then Exit Sub
           Next DoThis
        List3.AddItem (Screen_Name & " - SEND " & sNum$)
        Label9.Caption = List3.ListCount
        End If
        If sNum$ = "STATUS" Then
          Pen = 0
          For i = 0 To List3.ListCount - 1
              SN$ = Left(List3.List(i), InStr(1, List3.List(i), "-") - 2)
              If Screen_Name Like SN$ Then
                 Pen = Pen + 1
              End If
          Next i
          SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•–Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & Screen_Name & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "] You Have " & Pen & " Requests Pending"
          SendChat ("<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•–There Is No Mail Limits,So Request More")
          If List3.ListCount Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•–There Is Now " & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & Trim(Str(List3.ListCount)) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " Commands Pending"
        Exit Sub
       End If
     End If
End If
End Sub

Private Sub Chat1_ChatMsg(Screen_Name As String, Chat_Said As String)


If MChatBot = True Then
    frmChat.Text1 = frmChat.Text1 & Screen_Name & ":" & Chr(9) & Chat_Said & Chr(13) & Chr(10)
    frmChat.Text1.SelStart = Len(frmChat.Text1)
End If

If ServerBot = False Then Exit Sub
For DoThis = 0 To List6.ListCount - 1
      If UCase$(Screen_Name) Like UCase$(List6.List(DoThis)) Then Exit Sub
Next DoThis

If InStr(1, UCase$(Chat_Said), UCase$("<aolpromo>")) <> 0 Then
   Chat_Said = Mid$(Chat_Said, (InStr(Chat_Said, ">") + 1), Len(Chat_Said))
End If

If InStr(UCase$(Chat_Said), "/" & UCase$(AOLUserSN)) Then
    Srv$ = UCase$(Mid$(Chat_Said, Len(AOLUserSN()) + 3))
    If Srv$ Like "FIND *" Then
        fPhr$ = UCase$(Mid$(Chat_Said, Len(AOLUserSN) + 8))
        For DoThis = 0 To List5.ListCount - 1
           If List5.List(DoThis) Like Screen_Name & ";" & fPhr$ Then List5.RemoveItem a
        Next DoThis
        List5.AddItem (Screen_Name & ";" & fPhr$)
        Label3.Caption = List5.ListCount
    End If
    If Srv$ Like "SEND *" Then
       sNum$ = UCase$(Trim(Mid$(Chat_Said, Len(AOLUserSN) + 8)))
       If sNum$ = "LIST" Then
           For DoThis = 0 To List4.ListCount - 1
               If UCase$(List4.List(DoThis)) Like UCase$(Screen_Name) Then Exit Sub
           Next DoThis
           List4.AddItem (Screen_Name)
           Label16.Caption = List4.ListCount '& " Waiting For List"
           Exit Sub
        End If
       If InStr(sNum$, "-") <> 0& Then
          Dim First As String, Last As String
          First = Left$(sNum$, (InStr(sNum$, "-") - 1))
          Last = Mid$(sNum$, (InStr(sNum$, "-") + 1), Len(sNum$))
          
          If Last - First > 500 Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & Screen_Name & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]Stop Requesting Cuz You Will Be Banned": Exit Sub
          If Last - First > 25 Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & Screen_Name & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]Only 25 Request Are Allowed With The Quick Command": Exit Sub
            
            For DoThis = First To Last
               List3.AddItem (Screen_Name & " - SEND " & DoThis)
               Label9.Caption = List3.ListCount
               For DoIt = 0 To List3.ListCount - 1
                  If UCase$(List3.List(DoIt)) Like UCase$(Screen_Name) & " - SEND " & sNum$ Then Exit Sub
               Next DoIt
             Next DoThis
       End If
       If IsNumeric(sNum$) = True Then
           For DoThis = 0 To List3.ListCount - 1
              If UCase$(List3.List(DoThis)) Like UCase$(Screen_Name) & " - SEND " & sNum$ Then Exit Sub
           Next DoThis
        List3.AddItem (Screen_Name & " - SEND " & sNum$)
        Label9.Caption = List3.ListCount
        End If
        If sNum$ = "STATUS" Then
          Pen = 0
          For i = 0 To List3.ListCount - 1
              SN$ = Left(List3.List(i), InStr(1, List3.List(i), "-") - 2)
              If Screen_Name Like SN$ Then
                 Pen = Pen + 1
              End If
          Next i
          SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•–Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & Screen_Name & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "] You Have " & Pen & " Requests Pending"
          SendChat ("<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•–There Is No Mail Limits,So Request More")
          If List3.ListCount Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•–There Is Now " & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & Trim(Str(List3.ListCount)) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " Commands Pending"
        Exit Sub
       End If
     End If
End If

End Sub

Private Sub Command1_Click()

List5.SetFocus
AOL = FindWindow("AOL Frame25", vbNullString)
If AOL = 0 Then
  MsgBox "America Online Is Not Loaded", 16
  Status = "Ready"
  Exit Sub
End If
  Wel = FindChildByTitle(AOLMDI(), "Welcome,")
  Welc$ = String(255, 0)
  WhichWel = GetWindowText(Wel, Welc$, 250)
 If WhichWel < 8 Then
  MsgBox "You Need To Sign On First To Use The Server", 16
  Status = "Ready"
  Exit Sub
 End If
If Text3 = "" Then Text3.SetFocus: Exit Sub
If AOLFindChatRoom Then
  SendChat ("<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & Text3.text)
  Text3 = ""
  Text3.SetFocus
Else
Status = AOLUserSN() & ",You Gotta Be In A Chat Room."
Timeout 0.5
List2.SetFocus
Status = "Ready"
Exit Sub
End If

End Sub











Private Sub Command10_Click()
SendChat ("<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•–" & App.Title): DoEvents
   SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•–" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">Type /" & AOLUserSN() & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " Send List - ( " & Trim(Str$(Server.List2.ListCount)) & " ) Mails":  DoEvents
   SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•–" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">Type /" & AOLUserSN() & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " Send X - X Is Index":  DoEvents:
   SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•–" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">Type /" & AOLUserSN() & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " Find X - X Is Search Query"
   Timeout 0.4
  ' SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "<B>Download The Latest Copy <A HREF=""http://members.tripod.com/TiTo_Vb5/index.htm"">here!</A></B>"
   
End Sub

Private Sub Command2_Click()
Dim AOL As Long
List5.SetFocus
If label11.Caption = "0" Then
AOL& = FindWindow("AOL Frame25", vbNullString)
If AOL& = 0 Then
  MsgBox "America Online Is Not Loaded", 16
  Command2.Caption = "Start"
  Exit Sub
End If
  Wel = FindChildByTitle(AOLMDI, "Welcome,")
  Welc$ = String(255, 0)
  WhichWel = GetWindowText(Wel, Welc$, 250)
 If WhichWel < 8 Then
  MsgBox "You Need To Sign On First To Use The Server", 16
  Command2.Caption = "Start"
  Exit Sub
 End If
 If List2.ListCount = 0 Then
  X = MsgBox("You Need To Create A Mail List First!" & Chr(13) & "Would You Like To Create a List Now?", vbYesNo + 32)
  If X = 6 Then
  Call Command4_Click
   
 If List2.ListCount <> 0 Then
      MsgBox "You Are All Set Now Click Start To Start Serving.", 64
      Status = "Ready"
 Else
     MsgBox "The Server Could'nt Complete It's Task Please" & Chr(13) & Chr(10) & " Try Again Using The Refresh Mails Buttons!!", 16
     Command2.Caption = "Start"
     Exit Sub
 End If
 Else
   Status = "Ready"
   Command2.Caption = "Start"
   Exit Sub
   End If
 End If
  Room = AOLFindChatRoom()
  If Room = 0 Then
    MsgBox "You Must Be In a Chat Room To Start The Server.", 64
    Exit Sub
  End If
  Status = "Setting Server Prefrences."
  Pause 0.5
  Call SetMailPrefs
  Status = "Ready"
  SendKeys "{NumLock}"
  Status = "Turning Off Your IM's."
  Call InstantMessage("$IM_Off", App.Title): DoEvents
  Status = "Preparing Chat To Start Server"
  ClearChatText
  Status = "Ready"
  lblStatus = "On"
  Command2.Caption = "Stop"
  
   SendChat ("<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•–" & App.Title): DoEvents
   SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•–" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">Type /" & AOLUserSN() & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " Send List - ( " & Trim(Str$(Server.List2.ListCount)) & " ) Mails":  DoEvents
   SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•–" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">Type /" & AOLUserSN() & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " Send X - X Is Index":  DoEvents:
   SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•–" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">Type /" & AOLUserSN() & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " Find X - X Is Search Query"
   
   Call WriteINI("Server", "Cur", "0", App.Path & "\Server.ini")
    
   List1.Clear
   List3.Clear
   List4.Clear
   List5.Clear
    
   Label10.Caption = List1.ListCount
   Label9.Caption = List3.ListCount
    
   ServerBot = True
    
   label11.Caption = "1"
   Timer1.Enabled = True
   Timer2.Enabled = True
   Chat1.ScanOn
   'Timer6.Enabled = True                     'tell the ocx to start reading the chat
   
   Label14 = "0"
   Label13 = "1"
    
   MenuForm.itemNewMail.Enabled = False
   MenuForm.itemOld.Enabled = False
   MenuForm.itemSent.Enabled = False
   MenuForm.itemFlashMail.Enabled = False

    On Error Resume Next
    Kill "c:\Server.log"
    LogLine "Server started at " & Time & "."
    Timeout (2)
    
  Else
   'If List3.ListCount Then
    '    Prompt.Show
     '   Do
      '      DoEvents
       ' Loop Until PromptValue
        'Select Case PromptValue
         '   Case 1
          '      Timer2.Enabled = False
           '     Server.Chat1.ScanOff
            '    SendChat ("<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•–Server Has Stopped Logging Commands.")
             '   Timeout 0.1
              '  SendChat ("<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•–Commands Waiting Will Be Completed.")
               ' ServerBot = False
                'Command2.Caption = "Stop"
                'Do
                 'DoEvents
                 'Loop While List3.ListCount
           ' Case 3
            '    Command2.Caption = "Stop"
             '   Exit Sub
        'End Select
    'End If
  Call WriteINI("Server", "Cur", "0", App.Path & "\Server.ini")
  SendChat ("<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•–" & App.Title): DoEvents
  SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•–Server is [" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & "OFF" & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "],Please Stop Requesting!":  DoEvents
  If List1.ListCount Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•–Total Commands Completed:[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & Trim(Str$(List1.ListCount)) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]":  DoEvents
  label11.Caption = "0"
  ServerBot = False
  Timer1.Enabled = False
  Timer2.Enabled = False
  Command3.Caption = "Pause"
  'Timer6.Enabled = False
  Chat1.ScanOff
  Label13 = "0"
  Label14 = "1"
  
  Command2.Caption = "Start"
  Command3.Caption = "Pause"
  Command3.Tag = "0"
  List1.Clear
  List2.Clear
  List3.Clear
  List4.Clear
  List5.Clear
        
  Label5 = List2.ListCount
  Label10.Caption = List1.ListCount
  Label9.Caption = List3.ListCount
  Label16.Caption = List4.ListCount
  
  MenuForm.itemNewMail.Enabled = True
  MenuForm.itemOld.Enabled = True
  MenuForm.itemSent.Enabled = True
  MenuForm.itemFlashMail.Enabled = True
        
  lblStatus = "Off"
  Status = "Ready"
  
  LogLine "Server halted at " & Time & "."
  LogLine "Total commands completed: " & Trim(Str$(List1.ListCount))
  
  AOL& = FindWindow("AOL Frame25", vbNullString)
  If AOL& = 0& Then
   Exit Sub
  End If
  If AOLUserSN = "" Then
     Exit Sub
  End If
  Msg = MsgBox("Do You Want To Turn Your On Im", 32 + vbYesNo)
  If Msg = 6 Then Call InstantMessage("$IM_On", "•´)" & App.Title)
  X = MsgBox("Do You Want To Close Your MailBox?", 32 + vbYesNo)
  If X = 6 Then
     Call SendMessage(FindMailBox, WM_CLOSE, 0, 0&)
     Call SendMessage(FindFlashMail, WM_CLOSE, 0, 0&)
  End If
  LogLine "Server halted at " & Time & "."
  LogLine "Total commands completed: " & Trim(Str$(List1.ListCount))
 End If
End Sub
Private Sub Command3_Click()
List5.SetFocus
If Label13 = "0" Then MsgBox "The Server Is (Off) !.You Cant Pause" & Chr(13) & Chr(10) & " It While The Server Is Not Activated. ", 16: Exit Sub
If Command3.Tag = "0" Then
    Command3.Tag = "1"
    Timer1.Enabled = False
    Timer2.Enabled = False
    lblStatus.Caption = "Pause"
    Command3.Caption = "Resume"
    SendChat ("<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•–" & App.Title): DoEvents
    Timeout 0.2
    SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•–Server is Now [" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & "Pause" & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "],Still Taking Requests!":  DoEvents
Else
    Command3.Tag = "0"
    Timer1.Enabled = True
    Timer2.Enabled = True
    lblStatus.Caption = "On"
    Command3.Caption = "Pause"
    SendChat ("<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•–" & App.Title): DoEvents
    SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•–Server is Now [" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & "UnPause" & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "],Start Requesting!":  DoEvents
End If
End Sub

Private Sub Command4_Click()
List5.SetFocus
List2.Clear
AOL = FindWindow("AOL Frame25", vbNullString)
If AOL = 0 Then
    MsgBox "America Online Is Not Open.", 16
    Exit Sub
End If
Wel = FindChildByTitle(AOLMDI(), "Welcome,")
Welc$ = String(255, 0)
WhichWel = GetWindowText(Wel, Welc$, 250)
If WhichWel < 8 Then
    MsgBox "You Need To Sign On First To Use The Server.", 16
    Exit Sub
End If
 If MenuForm.itemFlashMail.Checked Then
     SendMessage FindFlashMail, WM_CLOSE, 0, 0
 Else
     SendMessage FindMailBox, WM_CLOSE, 0, 0
 End If
Picture1.Visible = True
If MenuForm.itemNewMail.Checked Then
Status = "Waiting For All Mails"
Call FindMailBox
Call MailWaitForLoadNew
Call MailToListNew(Server.List2)
Label5 = List2.ListCount
End If
If MenuForm.itemOld.Checked Then
Status = "Waiting For All Mails"
Call FindMailBox
Call MailWaitForLoadOld
Call MailToListOld(Server.List2)
Label5 = List2.ListCount
End If
If MenuForm.itemSent.Checked Then
Status = "Waiting For All Mails"
Call FindMailBox
Call MailWaitForLoadSent
Call MailToListSent(Server.List2)
Label5 = List2.ListCount
End If
If MenuForm.itemFlashMail.Checked Then
Status = "Waiting For All Mails"
Call FindFlashMail&
Call MailWaitForLoadFlash
Call MailToListFlash(Server.List2)
Label5 = List2.ListCount
End If
End Sub
Private Sub Command5_Click()
List5.SetFocus
PopupMenu MenuForm.menuPrefs, 0, Command5.Left, Command5.Top + Command5.Height
End Sub
Private Sub Command6_Click()
List5.SetFocus
PopupMenu MenuForm.menuOther, 0, Command6.Left, Command6.Top + Command6.Height
End Sub
Private Sub Command7_Click()
List5.SetFocus
SERVER_FILENAME = App.Path & "\Server.Dat"
SERVER_FINDFILE = App.Path & "\ServerFind.Dat"
If Label14 = "0" Then Exit Sub
X = MsgBox(AOLUserSN() & " Are You Sure You Want To Exit?", 64 + vbYesNo)
If X = 6 Then
On Error Resume Next
Kill SERVER_FILENAME
Kill SERVER_FINDFILE
Call QuitHelp
If MenuForm.itemSaveExit.Checked Then
If MenuForm.itemNewMail.Checked Then a$ = "0"
If MenuForm.itemOld.Checked Then a$ = "1"
If MenuForm.itemSent.Checked Then a$ = "2"
If MenuForm.itemFlashMail.Checked Then a$ = "3"
Call WriteINI("Preferences", "Mail", a$, App.Path & "\Server.ini")
If MenuForm.itemIm.Checked Then a$ = "1"
If MenuForm.ItemChat.Checked Then a$ = "2"
Call WriteINI("Preferences", "Notify", a$, App.Path & "\Server.ini")
Call WriteINI("Preferences", "IM", a$, App.Path & "\Server.ini")
If MenuForm.itemIdleBot.Checked Then a$ = "1" Else a$ = "0"
Call WriteINI("Preferences", "Idle", a$, App.Path & "\Server.ini")
If MenuForm.ItemWrite.Checked Then a$ = "1"
Call WriteINI("Preferences", "Write", a$, App.Path & "\Server.ini")
If MenuForm.ItemKill.Checked Then a$ = "1"
Call WriteINI("Preferences", "Kill", a$, App.Path & "\Server.ini")
If MenuForm.itemStatus.Checked Then a$ = "1"
Call WriteINI("Preferences", "Status", a$, App.Path & "\Server.ini")
Call WriteINI("Preferences", "Comment", mComm$, App.Path & "\Server.ini")
End If

AOL = FindWindow("AOL Frame25", vbNullString)
If AOL = 0 Then
 End
End If
Wel = FindChildByTitle(AOLMDI(), "Welcome,")
Welc$ = String(255, 0)
WhichWel = GetWindowText(Wel, Welc$, 250)
If WhichWel < 8 Then
End
End If
Timeout (1)
 SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "       «–=•(·•· " & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & "X-Treme Server '99" & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ·•·)•=–»·":  DoEvents:
 Timeout 0.4
 SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "        «–=•(·•·  " & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & " Now Unloaded * " & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "·•·)•=–»":  DoEvents:
Timeout (1)
End
End If

End Sub

Private Sub Command8_Click()
List6.SetFocus
MenuForm.itemSN.Visible = True
MenuForm.ItemName0.Visible = False
MenuForm.ItemName1.Visible = False
MenuForm.ItemName2.Visible = False
MenuForm.ItemName3.Visible = False
MenuForm.ItemName4.Visible = False
MenuForm.ItemName5.Visible = False
MenuForm.ItemName6.Visible = False
MenuForm.ItemName7.Visible = False
MenuForm.ItemName8.Visible = False
MenuForm.ItemName9.Visible = False
MenuForm.ItemName10.Visible = False
MenuForm.ItemName11.Visible = False
MenuForm.ItemName12.Visible = False
MenuForm.ItemName13.Visible = False
MenuForm.ItemName14.Visible = False
MenuForm.ItemName15.Visible = False
MenuForm.ItemName16.Visible = False
MenuForm.ItemName17.Visible = False
MenuForm.ItemName18.Visible = False
MenuForm.ItemName19.Visible = False
MenuForm.ItemName20.Visible = False
MenuForm.ItemName21.Visible = False
If AOLFindChatRoom Then
DoEvents
List7.Clear
Call AddRoomToListbox(List7)
For index% = 0 To List7.ListCount - 1
    DoEvents
    StringSpace$ = List7.List(index%)
    If StringSpace$ = "" Then GoTo EndRoomGather
    Select Case index%
        Case 0
            MenuForm.ItemName0.Caption = StringSpace$
            MenuForm.ItemName0.Visible = True
        Case 1
            MenuForm.ItemName1.Caption = StringSpace$
            MenuForm.ItemName1.Visible = True
        Case 2
            MenuForm.ItemName2.Caption = StringSpace$
            MenuForm.ItemName2.Visible = True
        Case 3
            MenuForm.ItemName3.Caption = StringSpace$
            MenuForm.ItemName3.Visible = True
        Case 4
            MenuForm.ItemName4.Caption = StringSpace$
            MenuForm.ItemName4.Visible = True
        Case 5
            MenuForm.ItemName5.Caption = StringSpace$
            MenuForm.ItemName5.Visible = True
        Case 6
            MenuForm.ItemName6.Caption = StringSpace$
            MenuForm.ItemName6.Visible = True
        Case 7
            MenuForm.ItemName7.Caption = StringSpace$
            MenuForm.ItemName7.Visible = True
        Case 8
            MenuForm.ItemName8.Caption = StringSpace$
            MenuForm.ItemName8.Visible = True
        Case 9
            MenuForm.ItemName9.Caption = StringSpace$
            MenuForm.ItemName9.Visible = True
        Case 10
            MenuForm.ItemName10.Caption = StringSpace$
            MenuForm.ItemName10.Visible = True
        Case 11
            MenuForm.ItemName11.Caption = StringSpace$
            MenuForm.ItemName11.Visible = True
        Case 12
            MenuForm.ItemName12.Caption = StringSpace$
            MenuForm.ItemName12.Visible = True
        Case 13
            MenuForm.ItemName13.Caption = StringSpace$
            MenuForm.ItemName13.Visible = True
        Case 14
            MenuForm.ItemName14.Caption = StringSpace$
            MenuForm.ItemName14.Visible = True
        Case 15
            MenuForm.ItemName15.Caption = StringSpace$
            MenuForm.ItemName15.Visible = True
        Case 16
            MenuForm.ItemName16.Caption = StringSpace$
            MenuForm.ItemName16.Visible = True
        Case 17
            MenuForm.ItemName17.Caption = StringSpace$
            MenuForm.ItemName17.Visible = True
        Case 18
            MenuForm.ItemName18.Caption = StringSpace$
            MenuForm.ItemName18.Visible = True
        Case 19
            MenuForm.ItemName19.Caption = StringSpace$
            MenuForm.ItemName19.Visible = True
        Case 20
            MenuForm.ItemName20.Caption = StringSpace$
            MenuForm.ItemName20.Visible = True
        Case 21
            MenuForm.ItemName21.Caption = StringSpace$
            MenuForm.ItemName21.Visible = True
        Case 22
            MenuForm.ItemName22.Caption = StringSpace$
            MenuForm.ItemName22.Visible = True
    End Select
Next index%
EndRoomGather:
If MenuForm.ItemName0.Visible = True Or MenuForm.ItemName1.Visible = True Then MenuForm.itemSN.Visible = False
PopupMenu MenuForm.menuIgnore, 0, Command8.Left, Command8.Top + Command8.Height
Else
MsgBox "You Gotta Be In A Chat Room", 16
Exit Sub
End If
End Sub

Private Sub Command9_Click()
List6.SetFocus
List6.Clear
End Sub
Private Sub Form_Load()
'Set Chat1 = New CHAT_CLASS
Call SetAppHelp(Me.hwnd)
Call PreVent
CenterFormTop Me
StayOnTop Me
Show
ReadPreferences:
Status = "Reading Setting Please Wait..."
If IFileExists(App.Path & "\Server.ini") Then
a$ = ReadINI("Server", "TOTAL", App.Path & "\Server.ini")
If a$ = "" Then Call WriteINI("Server", "Total", "0", App.Path & "\Server.ini")
a$ = ReadINI("Preferences", "Write", App.Path & "\Server.ini")
If a$ = "1" Then MenuForm.ItemWrite.Checked = True
a$ = ReadINI("Preferences", "Kill", App.Path & "\Server.ini")
If a$ = "1" Then MenuForm.ItemKill.Checked = True
a$ = ReadINI("Preferences", "Status", App.Path & "\Server.ini")
If a$ = "1" Then MenuForm.itemStatus.Checked = True
a$ = ReadINI("Preferences", "Mail", App.Path & "\Server.ini")
If a$ = "0" Or a$ = "" Then MenuForm.itemNewMail.Checked = True
If a$ = "1" Then MenuForm.itemOld.Checked = True
If a$ = "2" Then MenuForm.itemSent.Checked = True
If a$ = "3" Then MenuForm.itemFlashMail.Checked = True
a$ = ReadINI("Preferences", "Notify", App.Path & "\Server.ini")
If a$ = "1" Then MenuForm.itemIm.Checked = True
If a$ = "2" Or a$ = "0" Or a$ = "" Then MenuForm.ItemChat.Checked = True
a$ = ReadINI("Preferences", "Remove", App.Path & "\Server.ini")
If a$ = "1" Then MenuForm.itemRemove.Checked = True
If a$ <> "1" Then MenuForm.itemRemove.Checked = False
a$ = ReadINI("Preferences", "Idle", App.Path & "\Server.ini")
If a$ = "1" Then MenuForm.itemIdleBot.Checked = True Else MenuForm.itemIdleBot.Checked = False
mComm$ = ReadINI("Preferences", "Comment", App.Path & "\Server.ini")
If Len(mComm$) Then MenuForm.itemComments.Checked = True
Call WriteINI("Properties", "App.Title", (App.Title), App.Path & "\Server.ini")
a$ = ReadINI("Properties", "App.Title", App.Path & "\Server.ini")
If a$ <> App.Title Then Call WriteINI("Properties", "App.Title", (App.Title), App.Path & "\Server.ini")
App.Title = a$
a$ = ReadINI("Properties", "App.Path", App.Path & "\Server.ini")
If a$ <> App.Path Then Call WriteINI("Properties", "App.Path", App.Path, App.Path & "\Server.ini")
a$ = ReadINI("Preferences", "User", App.Path & "\Server.ini")
If a$ <> "1" Then
    X = MsgBox("Freely Distruibuted Free Usage Granted For All.This Program Was Made For Educational Purposes Only.The Maker Of This Program Is Not Responsible For He/She Actions." & Chr(13) & Chr(10) & "This Program Can Not Directly Violate AOL's Terms Of Services On It's Own.By Clicking On Agree You Assume Full Responsibility For Your Own Actions." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Do You Agree This Disclaimer?", 64 + vbYesNo)
    If X = 6 Then
      Call WriteINI("Preferences", "User", "1", App.Path & "\Server.ini")
    Else
      End
   End If
End If
Else
 If MenuForm.itemNewMail.Checked Then a$ = "0"
 If MenuForm.itemOld.Checked Then a$ = "1"
 If MenuForm.itemSent.Checked Then a$ = "2"
 If MenuForm.itemFlashMail.Checked Then a$ = "3"
 Call WriteINI("Preferences", "Mail", a$, App.Path & "\Server.ini")
 If MenuForm.itemIm.Checked Then a$ = "1"
 If MenuForm.ItemChat.Checked Then a$ = "2"
 Call WriteINI("Preferences", "Notify", a$, App.Path & "\Server.ini")
 Call WriteINI("Preferences", "IM", a$, App.Path & "\Server.ini")
 If MenuForm.itemIdleBot.Checked Then a$ = "1" Else a$ = "0"
 Call WriteINI("Preferences", "Idle", a$, App.Path & "\Server.ini")
 If MenuForm.ItemWrite.Checked Then a$ = "1"
 Call WriteINI("Preferences", "Write", a$, App.Path & "\Server.ini")
 If MenuForm.ItemKill.Checked Then a$ = "1"
 Call WriteINI("Preferences", "Kill", a$, App.Path & "\Server.ini")
 If MenuForm.itemStatus.Checked Then a$ = "1"
 Call WriteINI("Preferences", "Status", a$, App.Path & "\Server.ini")
 Call WriteINI("Preferences", "Comment", mComm$, App.Path & "\Server.ini")
 GoTo ReadPreferences
 End If
 Status = "Ready."
 Timeout (1)
  SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "       «–=•(·•·" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & " X-Treme Server '99 " & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "·•·)•=–»·":  DoEvents:
  SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "          «–=•(·•·  " & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & "Now Loaded *" & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "   ·•·)•=–»":  DoEvents:
 Timeout (1)
 CenterFormTop Me
 StayOnTop Me

End Sub

Private Sub List3_Click()
List3.RemoveItem (List3.ListIndex)
End Sub
Private Sub List4_Click()
List4.RemoveItem (List4.ListIndex)
End Sub
Private Sub List6_Click()
SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & List6 & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ]You Are Now UnBanned From The Server!": DoEvents
List6.RemoveItem (List6.ListIndex)
Label1.Caption = List6.ListCount
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then PopupMenu MenuForm.menuPrefs, 0, Picture2.Left, Picture2.Top + Picture2.Height
End Sub

Private Sub Text2_GotFocus()
Text3.SetFocus
End Sub

Private Sub Timer1_Timer()
Dim DoThis As Long, DoIt As Long
Dim b As String, a As String
Dim oWindow As Long, oButton As Long
Dim oStatic As Long, oString As String
Dim thelist As String, TheFind As String
Dim AOLErrorBox As Long, AOLErrorView As Long, ErrorString As String
Dim NameCount As Long, TempString As String, AOL As Long, mdi As Long, Error As Long
Dim OpenForward As Long, OpenSend As Long, SendButton As Long
Dim EditTo As Long, EditCC As Long
Dim EditSubject As Long, Rich As Long, fCombo As Long
Dim Combo As Long, Button1 As Long, Button2 As Long
Dim TempSubject As String
Dim ErrorHandle As Long, View As Long, ErrorIcon As Long, ErrorIcon2 As Long
Dim ErrIcon As Long, StatWindow As Long, Wel As Long
Dim MyString As String
Dim NoError As Long, NoErrorButton As Long
Dim index As Integer
Static Counter As Integer

Counter = Counter + 1

AOL& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
If AOL& = 0& Then
  Status = "Not Signed On.Please Sign On"
  Status = ""
  Exit Sub
End If
  Wel = FindChildByTitle(mdi&, "Welcome,")
  Welc$ = String(255, 0)
  WhichWel = GetWindowText(Wel&, Welc$, 250)
 If WhichWel < 8 Then
  Status = "Not Signed On.Please Sign On"
  Status = ""
  Exit Sub
 End If
 INI_FILENAME = App.Path + "\Settings\Server.ini"
 SERVER_FILENAME = App.Path + "\Server.dat"
 SERVER_FIND_FILENAME = App.Path + "\ServerFind.dat"
 AOL& = FindWindow("AOL Frame25", vbNullString)
 mdi& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
If List5.ListCount Then
      If List6.ListCount Then
         For DoThis = 0 To List5.ListCount - 1
          SN$ = Left(List5.List(DoThis), InStr(1, List5.List(DoThis), ";") - 1)
            For DoIt = 0 To List6.ListCount - 1
                 If InStr(Trim(UCase$(SN$)), Trim(UCase$(List6.List(DoIt)))) Then
                   List5.RemoveItem DoThis
                   Label3.Caption = List5.ListCount
                   Exit Sub
                 End If
            Next DoIt
         Next DoThis
      End If
    Counter = Counter + 1
    SN$ = Left(List5.List(0), InStr(1, List5.List(0), ";") - 1)
    If Trim(UCase$(AOLUserSN)) Like Trim(UCase$(SN$)) Then List5.RemoveItem (0)
    fPhr$ = UCase$(Mid$(List5.List(0), InStr(1, List5.List(0), ";") + 1))
    FindRes$ = "<P ALIGN=LEFT><FONT  COLOR=""#000000"" SIZE=3><U><B>Find Results for " & Chr$(34) + fPhr$ + Chr$(34) + ":</B></U></FONT><FONT  COLOR=""#FF0000"" SIZE=3>" + Chr$(13) + Chr$(10)
    On Error Resume Next
    Kill SERVER_FIND_FILENAME
    Open SERVER_FIND_FILENAME For Binary Access Write As #1
        Status = "Finding Requested..."
        Found = False
     For i = 0 To List2.ListCount - 1
        If UCase$(List2.List(i)) Like "*" & fPhr$ & "*" Then
            P$ = Chr$(13) & Chr$(10) & "(" & Trim(Str$(i)) & ")" & List2.List(i)
            Put #1, LOF(1) + 1, P$
            Found = True
        End If
    Next i
    If Found Then
        P$ = Chr$(13) & Chr$(10)
        Put #1, LOF(1) + 1, P$
    Else
        SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & SN$ & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]Your Search For [" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & "><U>" & fPhr$ & "</U>" & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]Was Not Found"
        Close #1
        List1.AddItem (SN$ & "*" & fPhr$)
        List5.RemoveItem (0)
        Label3.Caption = List5.ListCount
        Label10.Caption = List1.ListCount  '" Completed"
        LogLine "Find " & Chr$(34) & fPhr$ & Chr$(34) & " sent to " & SN$ & "."
        Status.Caption = "Ready."
        Exit Sub
    End If
    Close #1
    mff! = 1
    Var = 1
    
MailSearch:
    Status = "Sending Find Requested."
    AOL& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        OpenSend& = FindWindowEx(mdi&, 0&, "AOL Child", "Write Mail")
        EditTo& = FindWindowEx(OpenSend&, 0&, "_AOL_Edit", vbNullString)
        EditCC& = FindWindowEx(OpenSend&, EditTo&, "_AOL_Edit", vbNullString)
        EditSubject& = FindWindowEx(OpenSend&, EditCC&, "_AOL_Edit", vbNullString)
        Rich& = FindWindowEx(OpenSend&, 0&, "RICHCNTL", vbNullString)
        Combo& = FindWindowEx(OpenSend&, 0&, "_AOL_Combobox", vbNullString)
        fCombo& = FindWindowEx(OpenSend&, 0&, "_AOL_Fontcombo", vbNullString)
        Button1& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        Button2& = FindWindowEx(OpenSend&, Button1&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        For DoIt& = 1 To 13
            SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
        Next DoIt&
    Loop Until OpenSend& <> 0& And EditTo& <> 0& And EditCC& <> 0& And EditSubject& <> 0& And Rich& <> 0& And SendButton& <> 0& And Combo& <> 0& And fCombo& <> 0& & SendButton& <> Button1& And SendButton& <> Button2&
    
    Call SendMessageByString(EditTo&, WM_SETTEXT, 0, SN$)
      
    Open SERVER_FIND_FILENAME For Binary Access Read As #1
        
    If LOF(1) < 28900 Then
       Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, "•´)•[X-Treme Server '99]" & " - Find Results For [" & fPhr$ & "]")
    Else
       Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, "•´)•[X-Treme Server '99]" & " - Find Results #" & Trim(Str$(Var)) & "For [" & fPhr$ & "]")
    End If
    
   If Found Then
        Var = Var + 1
        TheFind$ = String(28900, 0)
        Get #1, mff!, TheFind$
        If InStr(1, TheFind$, Chr$(0)) Then TheFind$ = Left(TheFind$, InStr(1, TheFind$, Chr$(0)) - 1)
        If Right(TheFind$, 2) <> Chr$(13) & Chr$(10) Then
            Do
                DoEvents
                TheFind$ = Left(TheFind$, Len(TheFind$) - 1)
            Loop Until Right(TheFind$, 2) = Chr$(13) & Chr$(10)
        End If
        If Len(TheFind$) + mff! >= LOF(1) Then LastFind = True
    Else
        LastFind = True
    End If
    
   TheFind$ = FindRes$ & TheFind$
    
  Ch = Chr$(13) & Chr$(10)
  Ch2 = Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)
          
  Hj1$ = "<P ALIGN=CENTER><FONT COLOR=""#0000CC"" SIZE=3""><B>" & "       «–=•(·•· " & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & "X-Treme Server '99" & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ·•·)•=–»·" & Ch: DoEvents:
  Hj2$ = Hj1$ & "<P ALIGN=CENTER><FONT COLOR=""#0000CC"" SIZE=3""><B>" & "        «–=•(·•·  " & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & "  By TiTo *  " & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "·•·)•=–»" & Ch & "<FONT COLOR=""#0000CC"" SIZE=3""></B>Get A Copy <A HREF=""http://members.tripod.com/TiTo_Vb5/index.htm"">here!</A></B>" & Ch: DoEvents:
  
  Status = "Sending Mail Requested."
  Call SendMessageByString(Rich&, WM_SETTEXT, 0&, Hj2$ & Ch & TheFind$)
  Close #1
  Do Until EditTo& = 0 Or AOLErrorBox& <> 0&
    AOL& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    Wel = FindChildByTitle(mdi&, "Welcome,")
    Welc$ = String(255, 0)
    WhichWel = GetWindowText(Wel&, Welc$, 250)
    If WhichWel < 8 Then
         Status = "Not Signed On.Please Sign On"
         Exit Do
    End If
    OpenMail& = FindWindowEx(mdi&, 0&, "AOL Child", "Write Mail")
    EditTo& = FindWindowEx(OpenMail&, 0&, "_AOL_Edit", vbNullString)
    AOLErrorBox& = FindWindowEx(mdi&, 0&, "AOL Child", "Error")
    AOLErrorView& = FindWindowEx(AOLErrorBox&, 0&, "_AOL_View", vbNullString)
    If AOLErrorView& <> 0& Then
              nCounter& = ErrorNameCount
              MyString$ = ErrorName(1)
              For DoThis = 0 To List6.ListCount - 1
                If InStr(Trim(LCase(List6.List(DoThis))), Trim(LCase(MyString$))) Then List6.RemoveItem DoThis
              Next DoThis
              List6.AddItem MyString$
              For DoThis = 0 To List5.ListCount - 1
                    SN$ = Left(List5.List(DoThis), InStr(1, List5.List(DoThis), ";") - 1)
                    If InStr(Trim(UCase$(SN$)), Trim(UCase$(MyString$))) Then
                      List5.RemoveItem (DoThis)
                      Label3.Caption = List5.ListCount
                    End If
              Next DoThis
              SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & MyString$ & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "] Your Mailbox Is FULL!": DoEvents
              SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & MyString$ & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "] You're Now Banned From The Server!": DoEvents
              Label10.Caption = List1.ListCount
              Label9.Caption = List3.ListCount
              Label3.Caption = List5.ListCount
              Label1.Caption = List6.ListCount
              Status = "Ready"
         Call PostMessage(OpenMail&, WM_CLOSE, 0&, 0&)
         Do
           DoEvents
           Call PostMessage(AOLErrorBox&, WM_CLOSE, 0&, 0&)
           NoError& = FindWindow("#32770", "America Online")
           NoErrButton& = FindWindowEx(NoError&, 0&, "Button", "&No")
           Call PostMessage(NoErrButton&, WM_KEYDOWN, VK_SPACE, 0&)
           Call PostMessage(NoErrButton&, WM_KEYUP, VK_SPACE, 0&)
          Loop Until NoErrButton& <> 0&
       Exit Sub
  End If
  Status = "Sending ..Please Wait"
  AOLErrorBox& = FindWindowEx(mdi&, 0&, "AOL Child", "Error")
  If AOLErrorBox& <> 0& Then Exit Do
  Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
  Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
  StatWindow& = FindWindowEx(mdi&, 0&, "AOL Child", "Status")
  If StatWindow& <> 0& Then Call PostMessage(StatWindow&, WM_CLOSE, 0&, 0&)
Loop

If LastFind Then
       SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & SN$ & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]Your Search For [" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & "><U>" & fPhr$ & "</U>" & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]Was Found.Check Mail For Results."
  Else
        mff! = mff! + Len(TheFind$)
        GoTo MailSearch
  End If
  
   List1.AddItem (SN$ & "*" & fPhr$)
   
   LogLine "Find " & Chr$(34) & fPhr$ & Chr$(34) & " sent to " & SN$ & "."
   List5.RemoveItem (0)
   Label3.Caption = List5.ListCount
   Label10.Caption = List1.ListCount  '" Completed"
   Status = "Ready."
 End If
If List3.ListCount Then
   If List6.ListCount Then
      For DoThis = 0 To List3.ListCount - 1
            SN$ = Left(List3.List(DoThis), InStr(1, List3.List(DoThis), "-") - 2)
            For DoIt = 0 To List6.ListCount - 1
                 If InStr(Trim(UCase$(SN$)), Trim(UCase$(List6.List(DoIt)))) Then
                  List3.RemoveItem (DoThis)
                  Label9.Caption = List3.ListCount
                  Exit Sub
                 End If
            Next DoIt
       Next DoThis
End If
If UCase$(List3.List(0)) Like "*- SEND*" Then
  Counter = Counter + 1
  SN$ = Left(List3.List(0), InStr(1, List3.List(0), "-") - 2)
   DoEvents
   AOL& = FindWindow("AOL Frame25", vbNullString)
   mdi& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
   sNum$ = Mid$(List3.List(0), InStr(1, List3.List(0), "-") + 7)
   Label2 = Trim(sNum$)
   sNum$ = Mid$(List3.List(0), InStr(1, List3.List(0), "-") + 7)
   Label2 = Trim(sNum$)
   If IsNumeric(Label2) Then
            On Error Resume Next
            index = Label2
            If Err Then
                List3.RemoveItem (0)
                Label9.Caption = List3.ListCount '& " Waiting"
                Exit Sub
            End If
            If index >= List2.ListCount Then
                List3.RemoveItem (0)
                Label9.Caption = List3.ListCount ' & " Waiting"
                Exit Sub
            End If
            If index < 0 Then
                List3.RemoveItem (0)
                Label9.Caption = List3.ListCount '& " Waiting"
                Exit Sub
            End If
     If MenuForm.itemFlashMail.Checked Then
       fMail& = FindFlashMail
       If fMail& = 0 Then MsgBox "Please Open Your Imcoming/Saved Mails", 16: Exit Sub
       mTree& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
       Call ShowWindow(FindFlashMail, SW_MINIMIZE)
     Else
       MailBox& = FindMailBox
       If MailBox& = 0& Then MsgBox "Please Open Your MailBox ", 16: Exit Sub
       TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
       TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
       mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
       Call ShowWindow(FindMailBox, SW_MINIMIZE)
     End If
     
     Call SendMessage(FindForwardWindow, WM_CLOSE, 0, 0)
     Call SendMessage(List2.hwnd, LB_SETCURSEL, index, 0)
     Call SendMessage(mTree&, LB_SETCURSEL, index, 0)
     
     Status = "Checking For Mail #" & Trim(Str(Label2))
          
     If MenuForm.ItemWrite.Checked Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[ " & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & SN$ & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ]-Preparing Mail[#" & Trim(Str(Label2)) & " ]..Please Wait": Timeout 0.01
     
     If MenuForm.itemFlashMail.Checked = True Then
       Status = "Clicking Read Button."
        MailOpenEmailFlash (index)
     End If
     If MenuForm.itemNewMail.Checked = True Then
       Status = "Clicking Read Button."
        MailOpenEmailNew (index)
     End If
     If MenuForm.itemOld.Checked = True Then
       Status = "Clicking Read Button."
        MailOpenEmailOld (index)
     End If
     If MenuForm.itemSent.Checked = True Then
       Status = "Clicking Read Button."
        MailOpenEmailSent (index)
     End If
    StartTime = Timer
    Status = "Proccesing Mail # " & Trim(Str(Label2))
    Call RunMenuByString("S&top Incoming Text")
    
    
    
 Do
  DoEvents
       Call RunMenuByString("S&top Incoming Text")
       OpenForward& = FindForwardWindow
       If OpenForward& = 0 Then Exit Sub
       If FindSendWindow <> 0& Then Exit Do
       Call SendMessage(ForWardIcon, WM_LBUTTONDOWN, 0, 0&)
       Call SendMessage(ForWardIcon, WM_LBUTTONUP, 0, 0&)
 Loop
 Do
  DoEvents
    Call PostMessage(FindForwardWindow, WM_CLOSE, 0, 0&)
 Loop Until FindForwardWindow = 0
 Do
  DoEvents
        Status = "Proccesing Mail...Please Wait"
       
        AOL& = FindWindow("AOL Frame25", vbNullString)
        mdi& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
        OpenSend& = FindSendWindow
        EditTo& = FindWindowEx(OpenSend&, 0&, "_AOL_Edit", vbNullString)
        EditCC& = FindWindowEx(OpenSend&, EditTo&, "_AOL_Edit", vbNullString)
        EditSubject& = FindWindowEx(OpenSend&, EditCC&, "_AOL_Edit", vbNullString)
        Rich& = FindWindowEx(OpenSend&, 0&, "RICHCNTL", vbNullString)
        Combo& = FindWindowEx(OpenSend&, 0&, "_AOL_Combobox", vbNullString)
        fCombo& = FindWindowEx(OpenSend&, 0&, "_AOL_Fontcombo", vbNullString)
        Button1& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        Button2& = FindWindowEx(OpenSend&, Button1&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        For DoIt& = 1 To 13
            SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
        Next DoIt&
   Loop Until OpenSend& <> 0& And EditTo& <> 0& And EditCC& <> 0& And EditSubject& <> 0& And Rich& <> 0& And SendButton& <> 0& And Combo& <> 0& And fCombo& <> 0& & SendButton& <> Button1& And SendButton& <> Button2&
    
   Call SendMessageByString(EditTo&, WM_SETTEXT, 0, SN$)
    
   Status = "Setting The Text In The Mail."
   If MenuForm.itemRemove.Checked = True Then
      TempSubject$ = GetText(EditSubject&)
      TempSubject$ = Right(TempSubject$, Len(TempSubject$) - 5)
      Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, TempSubject$)
   End If
    
   Ch = Chr$(13) & Chr$(10)
   Ch2 = Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)
    
   Hj1$ = "<P ALIGN=CENTER><FONT COLOR=""#0000CC"" SIZE=3""><B>" & "       «–=•(·•· " & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & "X-Treme Server '99" & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ·•·)•=–»·" & Ch: DoEvents:
   Hj2$ = Hj1$ & "<P ALIGN=CENTER><FONT COLOR=""#0000CC"" SIZE=3""><B>" & "        «–=•(·•·  " & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & "  By TiTo *  " & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "·•·)•=–»" & Ch & "<FONT COLOR=""#0000CC"" SIZE=3""></B>Get A Copy <A HREF=""http://members.tripod.com/TiTo_Vb5/index.htm"">here!</A></B>" & Ch: DoEvents:
       
   If MenuForm.itemComments.Checked Then
      b$ = "</FONT></B><P ALIGN=LEFT>Mail Index #: " & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & "><B>" & Trim(Str(Label2)) & "</FONT></B>" & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & mComm$
   Else
      b$ = "</FONT></B><P ALIGN=LEFT>Mail Index #: " & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & "><B>" & Trim(Str(Label2)) & "</FONT></B>" & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "No Current Message"
   End If
       
   Call SendMessageByString(Rich&, WM_SETTEXT, 0, Hj2$ & b$) ' hj5$
    
   SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
    
   For DoIt& = 1 To 11
        SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
   Next DoIt&
   
Do Until EditTo& = 0& Or AOLErrorBox& <> 0&
    AOL& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    Wel = FindChildByTitle(mdi&, "Welcome,")
    Welc$ = String(255, 0)
    WhichWel = GetWindowText(Wel&, Welc$, 250)
    If WhichWel < 8 Then
         Status = "Not Signed On.Please Sign On"
         Exit Do
    End If
    EditTo& = FindWindowEx(OpenSend&, 0&, "_AOL_Edit", vbNullString)
    AOLErrorBox& = FindWindowEx(mdi&, 0&, "AOL Child", "Error")
    AOLErrorView& = FindWindowEx(AOLErrorBox&, 0&, "_AOL_View", vbNullString)
    If AOLErrorView& <> 0& Then
           nCounter& = ErrorNameCount
           MyString$ = ErrorName(1)
           For DoThis = 0 To List6.ListCount - 1
                If InStr(Trim(LCase(List6.List(DoThis))), Trim(LCase(MyString$))) Then List6.RemoveItem DoThis
           Next DoThis
           List6.AddItem MyString$
           SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & MyString$ & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "] Your Mailbox Is FULL!": DoEvents
           SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey," & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & MyString$ & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " You're Now Banned From The Server!": DoEvents
           Label10.Caption = List1.ListCount
           Label9.Caption = List3.ListCount
           Label3.Caption = List5.ListCount
           Label1.Caption = List6.ListCount
           Status = "Ready"
         For DoThis = 0 To List3.ListCount - 1
           SN$ = Left(List3.List(DoThis), InStr(1, List3.List(DoThis), "-") - 2)
           If InStr(Trim(UCase$(SN$)), Trim(UCase$(MyString$))) Then List3.RemoveItem DoThis
         Next DoThis
         Call PostMessage(OpenSend&, WM_CLOSE, 0&, 0&)
         Do
           DoEvents
           AOLErrorBox& = FindWindowEx(mdi&, 0&, "AOL Child", "Error")
           Call PostMessage(AOLErrorBox&, WM_CLOSE, 0&, 0&)
           DoEvents
           NoError& = FindWindow("#32770", "AOL Mail")
           NoErrButton& = FindWindowEx(NoError&, 0&, "Button", "&No")
           Call PostMessage(NoErrButton&, WM_KEYDOWN, VK_SPACE, 0&)
           Call PostMessage(NoErrButton&, WM_KEYUP, VK_SPACE, 0&)
          Loop Until NoErrorButton& <> 0&
        Exit Sub
     End If
     oWindow& = FindWindow("#32770", "America Online")
     oButton& = FindWindowEx(oWindow&, 0&, "Button", "OK")
     If oButton& <> 0& Then
        oStatic& = FindWindowEx(oWindow&, 0&, "Static", vbNullString)
        oStatic& = FindWindowEx(oWindow&, oStatic&, "Static", vbNullString)
        oString$ = GetText(oStatic)
        If InStr(Trim(UCase$(oString$)), Trim(UCase$("That message *"))) Then
        Call PostMessage(oButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(oButton&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(OpenSend&, WM_CLOSE, 0&, 0&)
        LogDead (List2 & Chr(9) & oString$)
        Do
           DoEvents
           NoError& = FindWindow("#32770", "AOL Mail")
           NoErrButton& = FindWindowEx(NoError&, 0&, "Button", "&No")
           Call PostMessage(NoErrButton&, WM_KEYDOWN, VK_SPACE, 0&)
           Call PostMessage(NoErrButton&, WM_KEYUP, VK_SPACE, 0&)
        Loop Until NoErrButton& <> 0&
        SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & SN$ & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "] Mail # " & Trim(Label2) & " Was Not Sent.Maybe Mail Is Dead": DoEvents
        List3.RemoveItem (0)
        Status = "Ready"
        Exit Sub
      End If
    End If
    
    AOLErrorBox& = FindWindowEx(mdi&, 0&, "AOL Child", "Error")
    If AOLErrorBox& <> 0& Then Exit Do
    Status = "Clicking Send..Please Wait."
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
    StatWindow& = FindWindowEx(mdi&, 0&, "AOL Child", "Status")
    If StatWindow& <> 0& Then Call PostMessage(StatWindow&, WM_CLOSE, 0&, 0&)
  Loop
  If MenuForm.ItemKill.Checked = True Then
      Call KillWait
  End If
  If MenuForm.itemNewMail.Checked = True Then
     Status = "Clicking Keep As New Button."
     Call ClickKeepAsNew
  End If
  List3.RemoveItem (0)
  List1.AddItem (SN$ & "*" & Label2)
  Status = "Ready."
  a$ = ReadINI("Server", "TOTAL", App.Path & "\Server.ini")
  a$ = a$ + 1
  Call WriteINI("Server", "Total", a$, App.Path & "\Server.ini")
  b$ = ReadINI("Server", "Cur", App.Path & "\Server.ini")
  b$ = b$ + 1
  Call WriteINI("Server", "Cur", b$, App.Path & "\Server.ini")
  If MenuForm.itemIm.Checked Then InstantMessage SN$, "Hey [" & SN$ & "]-#" & Trim(Label2) & " Was Sent."
  If MenuForm.ItemChat.Checked Then
    If mChat$ <> "" Then
       SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[ " & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & SN$ & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ]- #" & Trim(Label2) & " Was Sent." & mChat$
    Else
      SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & SN$ & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]- #" & Trim(Label2) & " Was Sent.[Total Served (" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & b$ & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & ")]"
    End If
  End If
  Label10.Caption = List1.ListCount
  Label9.Caption = List3.ListCount
  LogLine "Mail " & Label2 & " sent to " & SN$ & "."
  Else
  SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[ " & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & SN$ & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ]- #" & Trim(Label2) & " It's Not Avaliable."
  List3.RemoveItem (0)
  Label9.Caption = List3.ListCount
  End If
 End If
End If
Counter = Counter + 1
If Counter >= 50 Then
Counter = Counter + 1
Var = 1
mlf! = 1
SendList:
If List4.ListCount Then
  L = List4.ListCount
  Label16.Caption = List4.ListCount '& " Waiting For List"
  If L = 0 Then Exit Sub
  SN$ = ""
  For DoThis = 0 To L - 1
    DoEvents
     SN$ = SN$ & List4.List(DoThis) & ", "
  Next DoThis
 
 Status = "Scanning Screen Name(s)."
 
 SN$ = Left(SN$, Len(SN$) - 2)
    If List6.ListCount Then
       For DoIt = 0 To List4.ListCount - 1
          For DoThis = 0 To List6.ListCount - 1
                If InStr(Trim(LCase$(List4.List(DoIt))), Trim(LCase$(List6.List(DoThis)))) Then
                 List4.RemoveItem (DoIt)
                 Label16.Caption = List4.ListCount
                 Exit Sub
                End If
          Next DoThis
       Next DoIt
    End If
SendList2:
    Open SERVER_FILENAME For Binary Access Read As #1
    AOL& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        OpenSend& = FindWindowEx(mdi&, 0&, "AOL Child", "Write Mail")
        Status = "Waiting For First Window Handle."
        EditTo& = FindWindowEx(OpenSend&, 0&, "_AOL_Edit", vbNullString)
        EditCC& = FindWindowEx(OpenSend&, EditTo&, "_AOL_Edit", vbNullString)
        EditSubject& = FindWindowEx(OpenSend&, EditCC&, "_AOL_Edit", vbNullString)
        Rich& = FindWindowEx(OpenSend&, 0&, "RICHCNTL", vbNullString)
        Combo& = FindWindowEx(OpenSend&, 0&, "_AOL_Combobox", vbNullString)
        fCombo& = FindWindowEx(OpenSend&, 0&, "_AOL_Fontcombo", vbNullString)
        Button1& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        Button2& = FindWindowEx(OpenSend&, Button1&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        For DoIt& = 1 To 13
            SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
        Next DoIt&
    Loop Until OpenSend& <> 0& And EditTo& <> 0& And EditCC& <> 0& And EditSubject& <> 0& And Rich& <> 0& And SendButton& <> 0& And Combo& <> 0& And fCombo& <> 0& & SendButton& <> Button1& And SendButton& <> Button2&
    
    Call SendMessageByString(EditTo&, WM_SETTEXT, 0, SN$)
    Status = "Preparing List"
    If LOF(1) < 28900 Then
        Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, "•´)•[X-Treme Server '99]" & "- Mail List")
    Else
        Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, "•´)•[X-Treme Server '99]" & "- Mail List #" & Trim(Str$(Var)))
    End If
    thelist$ = String(28900, 0)
    Get #1, mlf!, thelist$
    If InStr(1, thelist$, Chr$(0)) Then thelist$ = Left(thelist$, InStr(1, thelist$, Chr$(0)) - 1)
    If Right(thelist$, 2) <> Chr$(13) & Chr$(10) Then
        Do
            DoEvents
            thelist$ = Left(thelist$, Len(thelist$) - 1)
        Loop Until Right(thelist$, 2) = Chr$(13) & Chr$(10)
    End If
    If Len(thelist$) + mlf! >= LOF(1) Then LastList = True
    Close #1
    mList$ = "<FONT  COLOR=""#000000"" SIZE=4><U><B><P ALIGN=LEFT>Mail List:</B></U></FONT><FONT  COLOR=""#FF0000"" SIZE=3>" & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)
    
    Ch = Chr$(13) & Chr$(10)
    Ch2 = Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)
    
    Hj1$ = "<P ALIGN=CENTER><FONT COLOR=""#0000CC"" SIZE=3""><B>" & "       «–=•(·•· " & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & "X-Treme Server '99" & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " ·•·)•=–»·" & Ch: DoEvents:
    Hj2$ = Hj1$ & "<P ALIGN=CENTER><FONT COLOR=""#0000CC"" SIZE=3""><B>" & "        «–=•(·•·  " & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & "  By TiTo *  " & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "·•·)•=–»" & Ch & "<FONT COLOR=""#0000CC"" SIZE=3""></B>Get A Copy <A HREF=""http://members.tripod.com/TiTo_Vb5/index.htm"">here!</A></B>" & Ch: DoEvents:
        
    Call SendMessageByString(Rich&, WM_SETTEXT, 0, Hj2$ & mList$ & thelist$)
    If MChatBot = True Then Call SetFocusAPI(frmChat.hwnd)
 
 Do Until EditTo& = 0 Or AOLErrorBox& <> 0
    
    AOL& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    Wel = FindChildByTitle(mdi&, "Welcome,")
    Welc$ = String(255, 0)
    WhichWel = GetWindowText(Wel&, Welc$, 250)
    If WhichWel < 8 Then
         Status = "Not Signed On.Please Sign On"
         Exit Do
    End If
    OpenMail& = FindWindowEx(mdi&, 0&, "AOL Child", "Write Mail")
    EditTo& = FindWindowEx(OpenMail&, 0&, "_AOL_Edit", vbNullString)
    AOLErrorBox& = FindWindowEx(mdi&, 0&, "AOL Child", "Error")
    AOLErrorView& = FindWindowEx(AOLErrorBox&, 0&, "_AOL_View", vbNullString)
    If AOLErrorView& <> 0& Then
           nCounter& = ErrorNameCount
           For People& = 1 To nCounter&
              MyString$ = ErrorName(People&)
              For DoThis = 0 To List6.ListCount - 1
                If InStr(Trim(LCase(List6.List(DoThis))), Trim(LCase(MyString$))) Then List6.RemoveItem DoThis
              Next DoThis
              List6.AddItem MyString$
              For DoThis = 0 To List4.ListCount - 1
                  If InStr(Trim(LCase$(List4.List(DoThis))), Trim(LCase$(MyString$))) Then List4.RemoveItem DoThis
              Next DoThis
              SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & MyString$ & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "] Your Mailbox Is FULL!": DoEvents
              SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey," & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & MyString$ & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " You're Now Banned From The Server!": DoEvents
              Label10.Caption = List1.ListCount
              Label9.Caption = List3.ListCount
              Label3.Caption = List5.ListCount
              Label1.Caption = List6.ListCount
              Status = "Ready"
        Next People&
        Call PostMessage(OpenMail&, WM_CLOSE, 0, 0)
        Do
          DoEvents
           Call PostMessage(AOLErrorBox&, WM_CLOSE, 0&, 0&)
           NoError& = FindWindow("#32770", "America Online")
           NoErrButton& = FindChildByTitle(NoError&, "&No")
           Call PostMessage(NoErrButton&, WM_KEYDOWN, VK_SPACE, 0&)
           Call PostMessage(NoErrButton&, WM_KEYUP, VK_SPACE, 0&)
          Loop Until NoErrButton& <> 0&
       Exit Sub
    End If
    
    Status = "Sending ..Please Wait"
    StatWindow& = FindWindowEx(mdi&, 0&, "AOL Child", "Status")
    If StatWindow& <> 0& Then Call PostMessage(StatWindow&, WM_CLOSE, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
Loop
Counter = 0
Status = "Ready"
If L >= 4 Then
      Timeout (1)
      DoEvents
      DoEvents
      
      SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & List4.List(0) & ", " & List4.List(1) & ", " & List4.List(2) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]-List #" & Trim(Str$(Var)) & " was sent."
      
      If L = 4 Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & List4.List(3) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]-List #" & Trim(Str$(Var)) & " was sent.": If LastList = True Then GoTo Last: DoEvents
      If L = 5 Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & List4.List(3) & ", " & List4.List(4) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]-List #" & Trim(Str$(Var)) & " was sent.": If LastList = True Then GoTo Last: DoEvents
      
      If L >= 6 Then
          SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & List4.List(3) & ", " & List4.List(4) & ", " & List4.List(5) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]-List #" & Trim(Str$(Var)) & " was sent."
          Timeout (1)
          If L = 6 And LastList = True Then GoTo Last
      End If
      
      If L = 7 Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & List4.List(6) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]-List #" & Trim(Str$(Var)) & " was sent.": If LastList = True Then GoTo Last: DoEvents
      If L = 8 Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & List4.List(6) & ", " & List4.List(7) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]-List #" & Trim(Str$(Var)) & " was sent.": If LastList = True Then GoTo Last: DoEvents
      If L >= 9 Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & List4.List(6) & ", " & List4.List(7) & ", " & List4.List(8) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]-List #" & Trim(Str$(Var)) & " was sent.": If L = 9 And LastList = True Then GoTo Last: DoEvents
      If L = 10 Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & List4.List(9) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]-List #" & Trim(Str$(Var)) & " was sent.": If LastList = True Then GoTo Last: DoEvents
      If L = 11 Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & List4.List(9) & ", " & List4.List(10) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]-List #" & Trim(Str$(Var)) & " was sent.": If LastList = True Then GoTo Last: DoEvents
      If L >= 12 Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & List4.List(9) & ", " & List4.List(10) & ", " & List4.List(11) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]-List #" & Trim(Str$(Var)) & " was sent.": If L = 12 And LastList = True Then GoTo Last: DoEvents
      If L = 13 Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & List4.List(12) & ", " & List4.List(10) & ", " & List4.List(11) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]-List #" & Trim(Str$(Var)) & " was sent.": If LastList = True Then GoTo Last: DoEvents
      If L = 14 Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & List4.List(12) & ", " & List4.List(13) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]-List #" & Trim(Str$(Var)) & " was sent.": If LastList = True Then GoTo Last: DoEvents
      If L >= 15 Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & List4.List(12) & ", " & List4.List(13) & ", " & List4.List(14) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]-List #" & Trim(Str$(Var)) & " was sent.": If L = 15 And LastList = True Then GoTo Last: DoEvents
      If L = 16 Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & List4.List(15) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]-List #" & Trim(Str$(Var)) & " was sent.": If LastList = True Then GoTo Last: DoEvents
      If L = 17 Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & List4.List(15) & ", " & List4.List(16) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]-List #" & Trim(Str$(Var)) & " was sent.": If LastList = True Then GoTo Last: DoEvents
      
      If L >= 18 Then
          SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & List4.List(15) & ", " & List4.List(16) & ", " & List4.List(17) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]-List #" & Trim(Str$(Var)) & " was sent.": DoEvents
          Timeout (1)
            If L = 18 And LastList = True Then GoTo Last
      End If
      
      If L = 19 Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & List4.List(18) & ", " & List4.List(19) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]-List #" & Trim(Str$(Var)) & " was sent.": DoEvents: If LastList = True Then GoTo Last: DoEvents
      If L = 20 Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & List4.List(18) & ", " & List4.List(19) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]-List #" & Trim(Str$(Var)) & " was sent.": DoEvents: If LastList = True Then GoTo Last: DoEvents
      If L >= 21 Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & List4.List(18) & ", " & List4.List(19) & ", " & List4.List(20) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]-List #" & Trim(Str$(Var)) & " was sent.": DoEvents: If L = 21 And LastList = True Then GoTo Last: DoEvents
      If L = 22 Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & List4.List(21) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]-List #" & Trim(Str$(Var)) & " was sent.": DoEvents: If LastList = True Then GoTo Last: DoEvents
      If L = 23 Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & List4.List(21) & ", " & List4.List(22) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]-List #" & Trim(Str$(Var)) & " was sent.": DoEvents: If LastList = True Then GoTo Last: DoEvents
      If L >= 24 Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & List4.List(21) & ", " & List4.List(22) & ", " & List4.List(23) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]-List #" & Trim(Str$(Var)) & " was sent.": DoEvents: If L = 24 And LastList = True Then GoTo Last: DoEvents
      If L = 25 Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & List4.List(24) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]-List #" & Trim(Str$(Var)) & " was sent.": DoEvents: If LastList = True Then GoTo Last: DoEvents
      If L = 26 Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & List4.List(24) & ", " & List4.List(25) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]-List #" & Trim(Str$(Var)) & " was sent.": DoEvents: If LastList = True Then GoTo Last: DoEvents
      If L >= 27 Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & List4.List(24) & ", " & List4.List(25) & ", " & List4.List(26) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]-List #" & Trim(Str$(Var)) & " was sent.": DoEvents: If L = 27 And LastList = True Then GoTo Last: DoEvents
      If L = 28 Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & List4.List(27) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]-List #" & Trim(Str$(Var)) & " was sent.": DoEvents: If LastList = True Then GoTo Last: DoEvents
      If L = 29 Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & List4.List(27) & ", " & List4.List(28) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]-List #" & Trim(Str$(Var)) & " was sent.": DoEvents: If LastList = True Then GoTo Last: DoEvents
      
      If L = 30 Then
          SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & List4.List(27) & ", " & List4.List(28) & ", " & List4.List(29) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "] - List #" & Trim(Str$(Var)) & " was sent.": DoEvents
          Timeout (1)
      End If
    Else
        SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "Hey,[" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & SN$ & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & "]- List #" & Trim(Str$(Var)) & " was sent.": DoEvents
        If MenuForm.itemStatus.Checked Then
           If Server.List3.ListCount Then SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•–There Is Now " & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & Trim(Str(List3.ListCount)) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " Commands Pending"
        End If
    End If
Last:
If LastList = True Then
       For i = 1 To L
            DoEvents
            List1.AddItem (List4.List(0) & "*" & "LIST")
            LogLine "List sent to " & List4.List(0) & "."
            On Error Resume Next
            List4.RemoveItem (0)
        Next i
        Label10.Caption = List1.ListCount '& " Completed"
        Label9.Caption = List3.ListCount '& " Waiting"
        Label16.Caption = List4.ListCount '& " Waiting For List"
        Exit Sub
End If
If LastList = False Then
        Var = Var + 1
        mlf! = mlf! + Len(thelist$)
        GoTo SendList2
End If
End If
End If
End Sub

Private Sub Timer2_Timer()
 Static Counter As Integer
 Counter = Counter + 1
 If Counter > 6 Then
 If List4.ListCount > 5 Then Exit Sub
 Timeout (1)
    SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•– X-Treme Server '99 Serving" & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">" & Trim(Str$(Server.List2.ListCount)) & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " Mails":  DoEvents
    SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•– Type " & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">/" & AOLUserSN & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " Send List For The Lists"
    SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•– Type " & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">/" & AOLUserSN & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " Send X-X For Mails Index":  DoEvents:
    SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "–•´)•– Type " & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & ">/" & AOLUserSN & "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & ">" & " Find X-X For Search Query":  DoEvents
   ' SendChat "<FONT COLOR=" & Chr(34) & "#0000CC" & Chr(34) & " FACE=""Arial Narrow""><B>" & "<B>Download The Latest Copy <A HREF=""http://members.tripod.com/TiTo_Vb5/index.htm"">here!</A></B>"
 Timeout (1)
 End If
End Sub
Private Sub Timer3_Timer()
If RoomBust = True Then Exit Sub
Dim OK As Long, Button As Long, NoButton As Long, AOModal As Long
AOL& = FindWindow("AOL Frame25", vbNullString)                'AOL Window Handle    ''
mdi& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)      'MDIClient Handle     ''
OK& = FindWindow("#32770", "America Online")
Button& = FindWindowEx(OK&, 0&, "Button", "OK")
If OK& <> 0& Then
     Button& = FindWindowEx(OK&, 0&, "Button", "OK")
      If InStr(Trim(UCase$(MsgMessage)), Trim(UCase$("That message is no longer available for forwarding."))) Then Exit Sub
      If InStr(Trim(UCase$(MsgMessage)), Trim(UCase$("You Are Now Ignoring Instant Messages"))) Then Exit Sub
      If InStr(Trim(UCase$(MsgMessage)), Trim(UCase$("You Are No Longer Ignoring Instant Messages"))) Then Exit Sub
      Call PostMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
      Call PostMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
End If
End Sub
Private Sub Timer5_Timer()
Dim AOL As Long, mdi As Long
AOL = FindWindow("AOL Frame25", vbNullString)
mdi = FindChildByClass(AOL&, "MDIClient")
Text2 = AOLUserSN
End Sub
Private Sub Timer6_Timer()
'Chat1.ScanChat
End Sub
Private Sub Timer7_Timer()
Dim AOIcon As Long, AOPalette As Long
AOPalette& = FindWindow("_AOL_PALETTE", vbNullString)
AOIcon& = FindWindowEx(AOPalette&, 0&, "_AOL_Icon", vbNullString)
If AOPalette& <> 0& Then
    AOIcon& = FindWindowEx(AOPalette&, 0&, "_AOL_Icon", vbNullString)
    Call SendMessage(AOIcon&, WM_LBUTTONDOWN, 0, 0&)
    Call SendMessage(AOIcon&, WM_LBUTTONUP, 0, 0&)
End If
If MenuForm.itemIdleBot.Checked = False Then Exit Sub
AOModal& = FindWindow("_AOL_Modal", vbNullString)
AOIcon& = FindWindowEx(AOModal&, 0&, "_AOL_Icon", vbNullString)
If AOModal& Then
   AOIcon& = FindWindowEx(AOModal&, 0&, "_AOL_Icon", vbNullString)
   ClickIcon (AOIcon&)
End If
End Sub
