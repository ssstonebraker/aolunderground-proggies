VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00004000&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Cartman's Eating-Range"
   ClientHeight    =   6936
   ClientLeft      =   3108
   ClientTop       =   2460
   ClientWidth     =   8064
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0E42
   ScaleHeight     =   6936
   ScaleWidth      =   8064
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   8880
      Top             =   3600
   End
   Begin VB.Shape Shape1 
      Height          =   732
      Left            =   120
      Top             =   120
      Width           =   1932
   End
   Begin VB.Label win2 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "GAME OVER!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   9855
   End
   Begin VB.Label win 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "GAME OVER!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   0
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   9852
   End
   Begin VB.Line cartrig 
      Visible         =   0   'False
      X1              =   5280
      X2              =   5280
      Y1              =   6840
      Y2              =   5520
   End
   Begin VB.Line cartlef 
      Visible         =   0   'False
      X1              =   4680
      X2              =   4680
      Y1              =   6840
      Y2              =   5520
   End
   Begin VB.Image pic2 
      Height          =   1488
      Left            =   6240
      Picture         =   "Form1.frx":FE624
      Top             =   5520
      Visible         =   0   'False
      Width           =   1848
   End
   Begin VB.Image pic1 
      Height          =   1488
      Left            =   1320
      Picture         =   "Form1.frx":FF862
      Top             =   5520
      Visible         =   0   'False
      Width           =   1848
   End
   Begin VB.Label poofs 
      BackStyle       =   0  'Transparent
      Caption         =   "25"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.Label score 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label2 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CHEESY POOFS:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "SCORE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Line CheesyPoof 
      BorderColor     =   &H000080FF&
      BorderWidth     =   10
      X1              =   9960
      X2              =   9975
      Y1              =   7200
      Y2              =   7215
   End
   Begin VB.Image cheesypoofs 
      Height          =   624
      Left            =   -120
      Picture         =   "Form1.frx":100A77
      Top             =   6480
      Width           =   792
   End
   Begin VB.Image cartman 
      Height          =   1488
      Left            =   3720
      Picture         =   "Form1.frx":103359
      Top             =   5520
      Width           =   1848
   End
   Begin VB.Image kennyd 
      Height          =   912
      Left            =   -2808
      Picture         =   "Form1.frx":10456E
      Top             =   1320
      Width           =   960
   End
   Begin VB.Menu mnugame 
      Caption         =   "&Game"
      Begin VB.Menu mnugamestart 
         Caption         =   "&Start/Restart"
      End
      Begin VB.Menu mnugameend 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&?"
      Begin VB.Menu mnuhelpinfo 
         Caption         =   "&Info"
      End
      Begin VB.Menu mnuhelphelp 
         Caption         =   "I&nstructions"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================
'       SOUTH PARK
'===========================================
'
'By Richard Nicol & Oliver Twardowski
'
'rnicol@khi-ro.co.uk & oli-t@topmail.de
'
'-------------------------------------------
'Instructions:
'
'Fire the cheesy poofs into Cartman's mouth!
'
'                        Mouseclick = Fire a cheesy
'                        poof in the mouse
'                        pointer direction
'
'============================================
'
'
'
'   Please help improve this game!!  ¦¬)
'   -----------------------------
'   Mail us!!!
'
'
'(all GFX and code were done by us)
'
'PSP4 & VB5 / VB 6


Dim kenny
Dim g_speed
Dim firex
Dim firey
Dim mo
Dim poofstart
Dim poofspilt
Dim cartmanshake
Dim pointsperhit
Dim extrashake
Dim extrapoints

Private Sub Form_Load()

poofstart = 30              'Cheesy poofs at start
poofspilt = poofstart       'Spill counter
cartmanshake = 20           'How much cartman shakes at start
extrashake = 10             'How much he shakes extra each go
pointsperhit = 10           'How many points you get for the first cheesy poof
extrapoints = 10            'How much the points per poof go up each go

poofs.Caption = poofstart   'Show cheesy poof counter

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If poofs < 1 Then           'If you're out of cheesy poofs
    finishit                'End game routine
    Exit Sub
End If

If Y < Form1.Height / 1.3 Then                                      'Click over a certain height
    If CheesyPoof.X1 > 0 And CheesyPoof.Y1 < Form1.Height Then      'If cheesy poof exits screen -
                                                                    'Do nothing
    Else                                                            'If not -
        Randomize
        If score > 500 Then kenny = "on"        'Check score and start Kenny if over 500

        cartman.Picture = pic1.Picture          'restore cartmans expression
        poofs = poofs - 1                       'take a cheesy poof out of the packet

        firey = Y                               'Set vertical direction to fire cheesy poof
        firex = X                               'Set horrizontal direction to fire cheesy poof
        CheesyPoof.Y2 = (Form1.Height)          'Put cheesy poof in the bottom corner of screen
        CheesyPoof.Y1 = (Form1.Height)               ' "
        CheesyPoof.X2 = 0                            ' "
        CheesyPoof.X1 = 0                            ' "
        mo = 20                                 'Set the momentum of the cheesy poof to 20
        g_speed = 1                             'Set the gravity speed (the rate the poof falls) to 1

        cartmanshake = cartmanshake + extrashake        'Increase cartmans movement
        pointsperhit = pointsperhit + extrapoints       'Increase points for the cheesy poof

    End If
End If
End Sub

Private Sub mnugameend_Click()
    
    End
    
End Sub

Private Sub mnugamestart_Click()

poofstart = 30              'Cheesy poofs at start
poofspilt = poofstart       'Spill counter
cartmanshake = 20           'How much cartman shakes at start
extrashake = 10             'How much he shakes extra each go
pointsperhit = 10           'How many points you get for the first cheesy poof
extrapoints = 10            'How much the points per poof go up each go

poofs.Caption = poofstart   'Show cheesy poof counter

End Sub

Private Sub mnuhelphelp_Click()
    
    Form3.Show 1
    
End Sub

Private Sub mnuhelpinfo_Click()
    
    Form2.Show 1
    
End Sub

Private Sub Timer1_Timer()

If kenny = "on" Then                            'If Kenny is activated
    kennyd.Left = kennyd.Left + 70              'Slide him accross the screen (H)
    kennyd.Top = kennyd.Top - 10                'Slide him accross the screen (V)
End If

cartlef.X1 = (cartman.Left + (cartman.Width / 2)) - 130     'Set up cartman mouth boundaries
cartlef.X2 = (cartman.Left + (cartman.Width / 2)) - 130
cartrig.X1 = (cartman.Left + (cartman.Width / 2)) + 230
cartrig.X2 = (cartman.Left + (cartman.Width / 2)) + 230

If CheesyPoof.X1 < cartrig.X1 Then                          'Check if the cheesy poof is in Cartman's mouth
    If CheesyPoof.X1 > cartlef.X1 Then
        If CheesyPoof.Y1 > 6720 Then
            If CheesyPoof.Y1 < 6920 Then
                cartman.Picture = pic2.Picture              'If cartman catches it, change his pic
                score = score + pointsperhit                'Add to score
                CheesyPoof.X1 = Form1.Width + 200                'Move cheesy poof off the screen
                CheesyPoof.X2 = Form1.Width + 200                ' "
                CheesyPoof.Y1 = Form1.Height + 200               ' "
                CheesyPoof.Y2 = Form1.Height + 200               ' "
                poofspilt = poofspilt - 1                   'Change spill counter
            End If
        End If
    End If
End If

If CheesyPoof.X1 > 0 And CheesyPoof.Y1 < Form1.Height Then        'If cheesy poof is on screen
    CheesyPoof.Y1 = CheesyPoof.Y1 + g_speed                       'make it fall according to the gravity velocity
    CheesyPoof.Y2 = CheesyPoof.Y2 + g_speed                       ' "
End If

g_speed = g_speed + 1.5                                 'Increase the gravity velocity

shootx = firex - (firex / 2)                            'Boring maths stuff (direction to move it right)
shooty = ((Form1.Height) - firey) / 2                   'Boring maths stuff (direction to move it up)

CheesyPoof.X1 = CheesyPoof.X1 + (shootx * mo / 100)               'More boring maths
CheesyPoof.X2 = CheesyPoof.X2 + (shootx * mo / 100)               '   How far to move right
CheesyPoof.Y1 = CheesyPoof.Y1 - (shooty * mo / 100)               'More boring maths
CheesyPoof.Y2 = CheesyPoof.Y2 - (shooty * mo / 100)               '   How far to move up

mo = mo / 1.05                                          'Decrease momentum

Randomize

If cartman.Left > 720 And cartman.Left < 7440 Then      'Check where cartman is on screen
    If Int(g_speed / 2) = g_speed / 2 Then cartman.Left = cartman.Left + (Rnd * cartmanshake) - (cartmanshake / 2)  'Shake it baby
Else
    If cartman.Left < 721 Then                          'If too far left
        If Int(g_speed / 2) = g_speed / 2 Then cartman.Left = cartman.Left + (Rnd * (cartmanshake / 2)) 'Shake it baby ( only right )
    End If
    If cartman.Left > 7439 Then                         'If too far right
        If Int(g_speed / 2) = g_speed / 2 Then cartman.Left = cartman.Left + (Rnd * (cartmanshake / 2)) - (cartmanshake / 2) 'Shake it baby ( only left )
    End If
End If

End Sub

Private Sub finishit()

win.Visible = True                  'Show 'GAME OVER' text

If poofspilt > 1 Then               'If you didn't win 100%
    win2.Caption = "You spilt " & poofspilt & " 'Cheesy Poofs', biyatch!  I am going home."     'Sort Cartman quote
Else
    win2.Caption = "Heh! I'm not fat, I'm big boned!"      'Sort Cartman quote
End If

win2.Visible = True                 'Show Cartman quote
Timer1.Enabled = False              'Stop game

End Sub
