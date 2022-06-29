VERSION 5.00
Object = "{DE8D4E3E-DD62-11D2-821F-444553540001}#1.0#0"; "CHATSCAN³.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "vybers help"
   ClientHeight    =   7305
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8385
   FillColor       =   &H80000004&
   ForeColor       =   &H80000004&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   8385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "help"
      Height          =   375
      Left            =   480
      TabIndex        =   26
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      Height          =   1215
      Left            =   2040
      TabIndex        =   24
      Top             =   1440
      Width           =   1335
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "hope this helps you learn the basics of c-chat and idles"
         Height          =   975
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.TextBox Text6 
      Height          =   1095
      Left            =   5040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Text            =   "helpidle.frx":0000
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   5640
      Top             =   6720
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   6240
      Top             =   6720
   End
   Begin chatscan³.Chat Chat2 
      Left            =   120
      Top             =   5520
      _ExtentX        =   4022
      _ExtentY        =   2275
   End
   Begin VB.TextBox Text5 
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Text            =   "helpidle.frx":00DF
      Top             =   4080
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   3720
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "helpidle.frx":014B
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   3720
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "helpidle.frx":0170
      Top             =   120
      Width           =   3135
   End
   Begin chatscan³.Chat Chat1 
      Left            =   3720
      Top             =   1680
      _ExtentX        =   4022
      _ExtentY        =   2275
   End
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   0
      TabIndex        =   3
      Top             =   1440
      Width           =   2055
      Begin VB.ListBox List1 
         Height          =   840
         ItemData        =   "helpidle.frx":01B4
         Left            =   120
         List            =   "helpidle.frx":01B6
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   3720
      Top             =   3120
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   2040
      TabIndex        =   2
      Top             =   0
      Width           =   1335
      Begin VB.TextBox Text2 
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Text            =   "helpidle.frx":01B8
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "helpidle.frx":01F0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label18 
      Caption         =   "some shit you can do with a timer"
      Height          =   375
      Left            =   5520
      TabIndex        =   22
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label Label17 
      Caption         =   "00:00"
      Height          =   255
      Left            =   6120
      TabIndex        =   21
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label16 
      Caption         =   "0/0/0"
      Height          =   255
      Left            =   6120
      TabIndex        =   20
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label Label15 
      Caption         =   "date:"
      Height          =   255
      Left            =   5640
      TabIndex        =   19
      Top             =   6240
      Width           =   615
   End
   Begin VB.Label Label14 
      Caption         =   "time:"
      Height          =   255
      Left            =   5640
      TabIndex        =   18
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "c-chat off"
      Height          =   255
      Left            =   6240
      TabIndex        =   17
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "c-chat on"
      Height          =   255
      Left            =   6240
      TabIndex        =   16
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   $"helpidle.frx":01F8
      Height          =   975
      Left            =   2640
      TabIndex        =   15
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   "some extra c-chat options i decided to add just the basics you can add to this"
      Height          =   615
      Left            =   2520
      TabIndex        =   14
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label9 
      Caption         =   "timer keeps you online by sending  text to chat every 60 secs thats why interval is set at 60000"
      Height          =   975
      Left            =   4320
      TabIndex        =   12
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "The c-chat options"
      Height          =   615
      Left            =   6120
      TabIndex        =   11
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "what to advertise on load"
      Height          =   495
      Left            =   6960
      TabIndex        =   10
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "vybers help tools"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Menu options 
      Caption         =   "&options"
      Begin VB.Menu idle 
         Caption         =   "idle"
         Begin VB.Menu on 
            Caption         =   "on"
         End
         Begin VB.Menu off1 
            Caption         =   "off"
         End
      End
      Begin VB.Menu chat 
         Caption         =   "c-chat"
         Begin VB.Menu on1 
            Caption         =   "on"
         End
         Begin VB.Menu off 
            Caption         =   "off"
         End
      End
      Begin VB.Menu clear 
         Caption         =   "clear msgs"
      End
      Begin VB.Menu view 
         Caption         =   "view msg"
      End
      Begin VB.Menu minimize 
         Caption         =   "minimize"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Chat1_ChatMsg(Screen_Name As String, What_Said As String)

If Screen_Name = getuser Then 'so only you can give the commands
   Dim lngSpace As Long, strCommand As String, strArgument1 As String
   Dim strArgument2 As String, lngComma As Long
   If InStr(What_Said, ".") = 1& Then ' tells text wil start with period
      lngSpace& = InStr(What_Said$, " ") 'gets what said
      If lngSpace& = 0& Then
         strCommand$ = What_Said$
      Else
         strCommand$ = Left(What_Said$, lngSpace& - 1&)
      End If
      Select Case strCommand$
      
      
      
      
  

 Case ".imson"
       Call IMsOn
        ChatSend "<b>" + (LCase(getuser) + " Ims Are On")
     
       
       Case ".imsoff"
       Call IMsOff
       ChatSend "<b>" + (LCase(getuser) + "Ims Are Off")
       
        Case ".x"
            If Len(What_Said$) > 4& Then
               strArgument1$ = Right(What_Said$, Len(What_Said$) - lngSpace&)
               chatignorebyname (strArgument1$)
                 ChatSend "<b>" + ("Xed ") + strArgument1$
               End If
               
              
               
        Case ".unx"
            If Len(What_Said$) > 4& Then
               strArgument1$ = Right(What_Said$, Len(What_Said$) - lngSpace&)
               chatignorebyname (strArgument1$)
                ChatSend "<b>" + ("Un Xed ") + strArgument1$
               End If
               
               Case ".im"
            lngComma& = InStr(What_Said$, ",")
            If lngComma& <> 0& And lngComma& > lngSpace& + 2& And lngComma& < Len(What_Said$) - 2& Then
               strArgument1$ = Trim(Mid(What_Said$, lngSpace& + 1&, lngComma& - lngSpace& - 1&))
               strArgument2$ = Trim(Right(What_Said$, Len(What_Said$) - lngComma& - 1&))
               If strArgument1$ <> "" And strArgument2$ <> "" Then
                  Call InstantMessage(strArgument1$, strArgument2$)
                  End If
                End If
              
                 Case ".mail"
            lngComma& = InStr(What_Said$, ",")
            If lngComma& <> 0& And lngComma& > lngSpace& + 2& And lngComma& < Len(What_Said$) - 2& Then
               strArgument1$ = Trim(Mid(What_Said$, lngSpace& + 1&, lngComma& - lngSpace& - 1&))
               strArgument2$ = Trim(Right(What_Said$, Len(What_Said$) - lngComma& - 1&))
               If strArgument1$ <> "" And strArgument2$ <> "" Then
                  sendmail strArgument1$, strArgument2$, strArgument2$
               End If
               End If
               
                   Case ".imx"
            If Len(What_Said$) > 4& Then
            strArgument1$ = Right(What_Said$, Len(What_Said$) - lngSpace&)
            imignore strArgument1$
            ChatSend "<b>" + ("Not Allowing IMs From" + strArgument1$)
            End If
            
             Case ".imunx"
             If Len(What_Said$) > 4& Then
             strArgument1$ = Right(What_Said$, Len(What_Said$) - lngSpace&)
             IMUnIgnore strArgument1$
             ChatSend "<b>" + ("Allowing Ims From" + strArgument1$)
             End If
             
              Case ".bust"
            If Len(What_Said$) > 4& Then
            strArgument1$ = Right(What_Said$, Len(What_Said$) - lngSpace&)
            ChatSend "<b>" + ("Busting room" + strarugment1)
            Do
            
            Call privateroom(strArgument1$)
            pause 2
            SendKeys "{enter}"
            Loop Until GetCaption(findroom) = strArgument1$
            ChatSend "<b>" + ("Busted into " + strarguments1$)
            End If
            
             Case ".idleon"
 Timer1.Enabled = True
 
 Case ".idleoff"
 Timer1.Enabled = False
     End Select
    End If
    
   End If
   
 
 
 
End Sub

Private Sub Chat2_ChatMsg(Screen_Name As String, What_Said As String)
If InStr(What_Said$, "/msg") Then 'gets what said and if said calls function
  List1.AddItem Screen_Name ' adds screen name that said it
Form2.List1.AddItem Screen_Name + ": " + What_Said
  pause 0.8 ' pause
ChatSend ("••" + Screen_Name + " message recorded.")
  End If
  
  
        
End Sub

Private Sub Command1_Click()
'text1 = mesage
'text2= font color example     <p><fontColor = "#F"F0000 " "face = "Aria"l Narrow"> put this in text2 make font color red and arial
'text3= advertisemtn example    <p><b>v</b>ybers - help tools <fontface="Wingdings 3"><b>g</b>
'text4= second line of advertisment example  <p><b>s</b>tatus.. -<b>l</b>oaded-
'text5= c-chat list some i have encluded .imson  .IMsOff  .im name, message .x name .unx name .bust room  .sendmail name , subject ,message
'text6= some shit about me   about-
'Handle -vyber
'email - vbforfree@ hotmail.com
'aim- its vyber
'comments - email on what
'you think bugs or whatever
'and if you want me to make
'another program example i
'enjoy helping people
'later -vyber

End Sub

Private Sub clear_Click()
List1.clear 'clears list 1
Form2.List1.clear 'clears msgs on form2
End Sub

Private Sub Form_Load()
'EMAIL ME WHAT YOU THINK WHAT SHOULD I ADD OR IF YOU NEED SOME HELP
'vbforfree@hotmail.com or im me at its vyber later
'      -vyber-
formontop Me 'keeps form on top
ChatSend Text3
pause 0.3
ChatSend Text4
End Sub


Private Sub Label1_Click()
WindowState = 1
Chat2.ScanOn 'turns ocx on for scanning chat for messages
 'minimize on start
Timer1.Enabled = True 'makes timer enabled
End Sub

Private Sub Label12_Click()
Chat1.ScanOn
End Sub

Private Sub Label13_Click()
Chat1.ScanOff
End Sub

Private Sub Label19_Click()
main.Show
End Sub

Private Sub Label2_Click()
Chat2.ScanOff
Timer1.Enabled = False 'stops idle
End Sub

Private Sub Label3_Click()
List1.clear 'clears list 1
Form2.List1.clear
End Sub

Private Sub Label4_Click()
WindowState = 1
End Sub

Private Sub Label6_Click()
Form1.Hide 'hides this form
Form2.Show 'shows message form
End Sub

Private Sub minimize_Click()
WindowState = 1
End Sub

Private Sub off_Click()
Label4.Caption = "c-chat off"
Chat1.ScanOff
End Sub

Private Sub off1_Click()
Label4.Caption = "idle off..."
Chat2.ScanOff
Timer1.Enabled = False 'stops idle
End Sub

Private Sub on_Click()
Label4.Caption = "idle...."
WindowState = 1
Chat2.ScanOn 'turns ocx on for scanning chat for messages
 'minimize on start
Timer1.Enabled = True 'makes timer enabled
End Sub

Private Sub on1_Click()
Label4.Caption = "c-chat on"
Chat1.ScanOn
End Sub

Private Sub Text2_Change()
'ex font color this wil send messages red with arial narrow font
End Sub

Private Sub Timer1_Timer()
Label5.Caption = Val(Label5) + 1
ChatSend Text2 + Text1 + " away" + " [" & Label3.Caption & "] mins "
'ok in the first box text2 i use that to change your font color or font type its easyer
'that way just put the html in there text2 is your message and label 5 is
'your away time i had label5 caption start at 1 bec of the little mins and min
'thing between 1 and 2 and dont fill like fixing it val(label5) + 1 just increases the value of
'of label 5 by one every min to determine how long you have been along
End Sub

Private Sub Timer2_Timer()
Label16.Caption = Date

End Sub

Private Sub Timer3_Timer()
Label17.Caption = Time
End Sub

Private Sub view_Click()
Form1.Hide 'hides this form
Form2.Show 'shows message form
End Sub
