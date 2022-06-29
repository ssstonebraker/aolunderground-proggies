VERSION 5.00
Begin VB.Form Questions 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   360
      TabIndex        =   1
      Text            =   "I'm going insain"
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "ok"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Message"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Questions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
' this whole welcome widow is a lot of bullshit i wrote when i was high. I'm sure u can see what i was trying to go for, most people  are retarded using servers so this would help those first time servers to get servin faster and sending out more porn, mp3's moveis warez w/e
Me.Hide
pissmeoff = 0
'Gets user name
If Welcome.Combo1.Text = "New UseR?" Then
    answer$ = InputBox("i need a name", "4:20")
End If
If answer$ = "" Then
    answer$ = InputBox("So... You fuckin think you don't got a name ehh... p   u   t    i     n     y    o   u   r    f    u   c   k   i   n   nikname!, i hope u only see this message once more in ur lifetime, you shouldn't be seeing it right now, but u thought you were a hot shot clickin around My COMUTER ;0)", "4:20")
    If answer$ = "" Then End
    pissmeoff = pissmeoff + 1
End If
Call WriteToINI("user", "user", answer$, App.Path & "\" & answer$)
User = GetFromINI("user", "user", App.Path & "\" & answer$)
'gets list size
Do:
    answer$ = InputBox("Well hello der " & User & ", My Name = insain, i will be ur fuckin tour guide to my options.  If u havn't figured it out yet, there are secret tunnels everywhere on dis server.  lol i hope this is not six, the possiblities of insain knowing ur six is horrible.  cause he don't like six. anywayz,  How many lines of warez/porn/mp3 w/e would u want in each list u send to people.  don't read this twice just put in an answer i'm bored!         (1 - 500 ???) Pist me off rate = " & pissmeoff & " times already!", "4:20")
        If answer$ = "" Then
            pissmeoff = pissmeoff + 1
            If pissmeoff = 10 Then
                End
            End If
        End If
Loop Until IsNumeric(answer$)
Call WriteToINI("listcount", "listcount", answer$, App.Path & "\" & User)
'gets block size
Do:
    answer$ = InputBox("How many blocks to ur friends house?,  InsaiN Wants to know, i never had any friends so i put a lot on the block size ;-) (1 - 500) Pist me off rate = " & pissmeoff & " times already!", "4:20")
        If answer$ = "" Then
            pissmeoff = pissmeoff + 1
            If pissmeoff = 10 Then
                End
            End If
        End If
Loop Until IsNumeric(answer$)
Call WriteToINI("blocksize", "blocksize", answer$, App.Path & "\" & User)
'gets restart time
Do:
    answer$ = InputBox(User & "this is what weed does to u :-(! Don't ever ask me why i don't have the delete sent mail option cause its pointless.  Eww ur gonna get traced... what is u... bitch?, i suggest u restart aol every 500 mailz but 1000 is coo too, what u want it to be? Pist me off rate = " & pissmeoff & " times already!", "4:20")
    If answer = "" Then
        pissmeoff = pissmeoff + 1
        If pissmeoff = 10 Then
            End
        End If
    End If
Loop Until IsNumeric(answer$)
Call WriteToINI("restartaol", "restartaol", answer$, App.Path & "\" & User)
'gets max find per person limit
Do:
    answer$ = InputBox("::Yawn::,  O.k whats da Max Find Limit per Person, this option is for dem retards that think finds kill servers, inasain = da best so this option is not to be disabled, finds are always on cause dats da way i like it. Pist me off rate = " & pissmeoff & " times already!", "4:20")
    If answer = "" Then
        pissmeoff = pissmeoff + 1
        If pissmeoff = 10 Then
            End
        End If
    End If
Loop Until IsNumeric(answer$)
Call WriteToINI("Findperperson", "Findperperson", answer$, App.Path & "\" & User)
Me.Show
MsgBox "Write in the message you want peeps to see when u send dem mailz lets get goin i wanna start servin Pist me off rate = " & pissmeoff & " times already!", vbOKOnly, "4:20"
KeepFormOnTop Me
End Sub

Private Sub Label2_Click()
Me.Hide
FormNotOnTop Me
ServerForm.Show
End Sub
