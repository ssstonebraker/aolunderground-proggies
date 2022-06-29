VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Minimize 
         Caption         =   "&Minimize Regular"
      End
      Begin VB.Menu DivideBar 
         Caption         =   "-"
      End
      Begin VB.Menu About 
         Caption         =   "&About"
      End
      Begin VB.Menu Divide 
         Caption         =   "-"
      End
      Begin VB.Menu Exir 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu Options 
      Caption         =   "Options"
      Begin VB.Menu IMS 
         Caption         =   "IMS"
         Begin VB.Menu On 
            Caption         =   "On"
         End
         Begin VB.Menu Off 
            Caption         =   "Off "
         End
      End
      Begin VB.Menu Divide2 
         Caption         =   "-"
      End
      Begin VB.Menu Idle 
         Caption         =   "&Idle Bot"
      End
      Begin VB.Menu Divide3 
         Caption         =   "-"
      End
      Begin VB.Menu Id 
         Caption         =   "I&dentifier"
      End
      Begin VB.Menu Divideme 
         Caption         =   "-"
      End
      Begin VB.Menu Secret 
         Caption         =   "&Secret Access"
      End
      Begin VB.Menu ffd 
         Caption         =   "-"
      End
      Begin VB.Menu Greetz 
         Caption         =   "&Greetz"
      End
      Begin VB.Menu Divide4 
         Caption         =   "-"
      End
      Begin VB.Menu Sites 
         Caption         =   "Sites"
         Begin VB.Menu KnK 
            Caption         =   "&KnK"
         End
         Begin VB.Menu MySite 
            Caption         =   "&My Site"
         End
         Begin VB.Menu KTHON 
            Caption         =   "&KTHON"
         End
         Begin VB.Menu Plugin 
            Caption         =   "&Best Plugin Warehouse"
         End
      End
      Begin VB.Menu Divide5 
         Caption         =   "-"
      End
      Begin VB.Menu Advertize 
         Caption         =   "&Advertize"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub About_Click()
MsgBox "SuP yo, I made this program for fun, and I am in no way am responcible for what you do with this program, neither is anyone else, so if you get caught its your own head. Now if you know me, please don't ask to make a prog with me for a while I'm working on a top secret prog with someone and it's gonna be BIG!, a first of its kind. L8'z.", vbOKOnly, " Water Rapids By FeaR"
End Sub

Private Sub Advertize_Click()
ChatSend "" & (" ")
ChatSend "<font face=""Arial Narrow""></B></I></U></S>" & BlueGreen("‹v^•][ Water Rapids")
TimeOut 0.3
ChatSend "<font face=""Arial Narrow""></B></I></U></S>" & BlueGreen("‹v^•][ By FeaR")
TimeOut 0.3
ChatSend "<font face=""Arial Narrow""></B></I></U></S>" & BlueGreen("‹v^•][ User: " & UserSN + "")
TimeOut 0.3
ChatSend "<font face=""Arial Narrow""></B></I></U></S>" & BlueGreen("‹v^•][ Time: " & TrimTime2 + "")
TimeOut 0.3
ChatSend "<font face=""Arial Narrow""></B></I></U></S>" & BlueGreen("‹v^•][ Date: " & TrimDate + "")
TimeOut 0.3
ChatSend "<font face=""Arial Narrow""></B></I></U></S>" & BlueGreen("‹v^•][ Please keep all hands in the tube, Thank You!")
TimeOut 0.3
ChatSend "" & (" ")
End Sub

Private Sub Exir_Click()
End
End Sub

Private Sub Greetz_Click()
MsgBox "Ok I can't name everything cause I have so many friends and a lot of people help me, but I would especially like to say wussup to ItaL, onix, wizy, wizo, dolan, slush, ginn, comp, blo0d, sect, chud, and sushi...oh and to wHa!...to everyone else I'm sorry... trust me!...L8'z", vbOKOnly, " Greetz"
End Sub

Private Sub Id_Click()
Form4.Show
End Sub

Private Sub Idle_Click()
Form3.Show
End Sub

Private Sub KnK_Click()
Call KeyWord("http://www.nwozone.com/knk4o/index2.htm")
End Sub

Private Sub KTHON_Click()
Call KeyWord("http://members.xoom.com/kthon/goodies.html")
End Sub

Private Sub Minimize_Click()
Form1.WindowState = 1
End Sub

Private Sub MySite_Click()
Call KeyWord("http://fear99.cjb.net")
End Sub

Private Sub Off_Click()
Call IM_Keyword("$IM_OFF", " FeaR OwnZ ")
End Sub

Private Sub On_Click()
Call IM_Keyword("$IM_ON", " FeaR OwnZ ")
End Sub

Private Sub Plugin_Click()
Call KeyWord("http://www.dirtysouth.net./filter/home.html")
End Sub

Private Sub RB_Click()
Form4.Show
End Sub

Private Sub Secret_Click()
Form5.Show
End Sub
