VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form menu 
   Caption         =   "Form1"
   ClientHeight    =   1245
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   2940
   LinkTopic       =   "Form1"
   ScaleHeight     =   1245
   ScaleWidth      =   2940
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check3 
      Caption         =   "save2last list"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "chat notify????"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "load iface??????"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CmDialog1 
      Left            =   2160
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label clearlists 
      Caption         =   "used for clearing both lists"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu mnuadvertise 
         Caption         =   "Advertise"
         Begin VB.Menu mnuaolchatroomadv 
            Caption         =   "Aol ChatRoom"
         End
      End
      Begin VB.Menu asdf 
         Caption         =   "-"
      End
      Begin VB.Menu mnucontact 
         Caption         =   "Contact"
      End
      Begin VB.Menu sdfgsadfg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuminimize 
         Caption         =   "Minimize"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnusettings 
      Caption         =   "Settings"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu mnuloadintro 
         Caption         =   "Load Intro"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuchatnotifyon 
         Caption         =   "Chat Notify On"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnusavetolastopenedlistafteraction 
         Caption         =   "Save To Last Opened List After Action"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnulistoptions 
      Caption         =   "List Options"
      Begin VB.Menu mnusendaims 
         Caption         =   "Send # of Aims to Aol Chat"
      End
      Begin VB.Menu asdfasdf 
         Caption         =   "-"
      End
      Begin VB.Menu mnusavesnandpwlist 
         Caption         =   "Save Sn and Pw List"
      End
      Begin VB.Menu mnuloadsnandpwlist 
         Caption         =   "Load Sn and Pw List"
      End
      Begin VB.Menu dfgeggf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuaddsnsandpws 
         Caption         =   "Add Sn(s) and Pw(s)"
      End
      Begin VB.Menu fhfghjgf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuclearbothlists 
         Caption         =   "Clear Sn and Pw List"
      End
      Begin VB.Menu hhgh4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuremoveselected 
         Caption         =   "Remove Selected"
      End
   End
   Begin VB.Menu mnumore 
      Caption         =   "More"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu mnusourcecode 
         Caption         =   "Source Code"
      End
   End
End
Attribute VB_Name = "menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'loads formstate for the form: menu
LoadFormState menu

'if check1 isnt checked...
If menu.Check1.Value = 0 Then

'then in the popup menu, it unchecks it
menu.mnuloadintro.Checked = False
'ends the if statement
End If

'checks to see if chatnotify is on...
If menu.Check2.Value = 0 Then
menu.mnuchatnotifyon.Checked = False
End If
'checks to see if save2list is on....
If menu.Check3.Value = 0 Then
menu.mnusavetolastopenedlistafteraction.Checked = False
End If
End Sub

Private Sub mnuabout_Click()

End Sub

Private Sub mnuaddsnsandpws_Click()
'this coding was by my friend syfa.
'commented by source

'on error skip coding
On Error GoTo bah:

'minimizes it so that only the input box is visible
frmmain.WindowState = 1

'Do statement..do events
Do: DoEvents

'xxx is the variable used ...definded as inputbox
xxx = InputBox("enter sn:pw in this format please.")

'adds what was choosen after taking out the :
frmmain.List1.AddItem (Mid(xxx, 1, InStr(1, xxx, ":") - 1))

'adds what was choosen after taking out the :
frmmain.List2.AddItem (Mid(xxx, InStr(xxx, ":") + 1))
If menu.Check3.Value = 1 Then
Call Save2Lists(frmmain.List1, frmmain.List2, menu.CmDialog1.FileName)
End If
'makes input box appear until vbcancel is choosen.
Loop Until xxx = vbCancel

'back to normal with windowstate :)
frmmain.WindowState = 0
'this where to goes to when errored(basically skips
'over the coding)
bah:
frmmain.WindowState = 0
End Sub

Private Sub mnuaolchatroomadv_Click()
Call adv
End Sub

Private Sub mnuchatnotifyon_Click()
LoadFormState menu

If mnuchatnotifyon.Checked = True Then

mnuchatnotifyon.Checked = False

Check2.Value = 0

Else

mnuchatnotifyon.Checked = True

Check2.Value = 1

End If

SaveFormState menu
End Sub

Private Sub mnuclearbothlists_Click()
'makes the caption = the input box
clearlists.Caption = InputBox("Type: Y or N:", "Are You Sure?", "N")

'if after they choose an answer if its a Y or y then
If clearlists.Caption = "Y" Or clearlists.Caption = "y" Then

'it clears both lists
frmmain.List1.Clear
frmmain.List2.Clear

'if not then moves on with coding
End If

'if the caption after choosing a answer is a N, or n
If clearlists.Caption = "N" Or clearlists.Caption = "n" Then

'then nothing happens because you dont want to clear
End If

'resets frmmains listcount caption's
frmmain.list1count.Caption = frmmain.List1.ListCount
frmmain.list2count.Caption = frmmain.List2.ListCount
End Sub

Private Sub mnucontact_Click()
Call MsgBox("aim:hackertaLk  ,   aol: itzdasource@aol.com   ,   mail: source@terrorfx.com")
End Sub

Private Sub mnuexit_Click()
Unload frmintro
Unload frmmain
Unload menu
End Sub

Private Sub mnuloadintro_Click()
'if its already checked then
If mnuloadintro.Checked = True Then

'it unchecks it because you want it to do opposite.
mnuloadintro.Checked = False

'sets check1's value to 0 (unchecked)
menu.Check1.Value = 0
'if not that....
SaveFormState menu
Else

'if its not checked..then it needs to be checked:
mnuloadintro.Checked = True

'sets check1's value to 1 (checked)
menu.Check1.Value = 1

'ends if statement
SaveFormState menu
End If

'saves form settings(state) for frmmain & menu
SaveFormState frmmain
SaveFormState menu
End Sub

Private Sub mnuloadsnandpwlist_Click()
'previous work with common dialog was commented(see save/load cmdialog)
    CmDialog1.DialogTitle = "[Evil Aim Sn Maker 2] Load Sn And Pw List"
    
    CmDialog1.InitDir = App.Path
    
    CmDialog1.FLAGS = &H4
    
    CmDialog1.Filter = "list files (*.lst)|*.lst|all files (*.*)|*.*"
    
    CmDialog1.ShowOpen
    
    'if the filename you choose exists then...
    If FileExists(CmDialog1.FileName) = True Then
        
        'it loads them using a sub from module
        Call Load2Lists(frmmain.List1, frmmain.List2, CmDialog1.FileName)
    End If

'sets frmmain's two list count captions to there listcount
frmmain.list1count.Caption = frmmain.List1.ListCount
frmmain.list2count.Caption = frmmain.List2.ListCount

'horizontal scroll bar to both list1 and list2
Call AddHScroll(frmmain.List1)
Call AddHScroll(frmmain.List2)
End Sub

Private Sub mnuminimize_Click()
frmmain.WindowState = 1
End Sub

Private Sub mnuremoveselected_Click()
'removes list1 & 2's selected item
ListRemoveSelected (List1)
ListRemoveSelected (List2)
End Sub

Private Sub mnusavesnandpwlist_Click()
'previous work with common dialog was commented(see save/load cmdialog)
    CmDialog1.DialogTitle = "[Evil Aim Sn Maker 2 By Source] Save Sn And Pw"
    
    CmDialog1.InitDir = App.Path 'default dir: what exe is in.
    
    CmDialog1.FLAGS = &H4
    
    CmDialog1.Filter = "list files (*.lst)|*.lst|all files (*.*)|*.*" 'file extention options
    
    CmDialog1.ShowSave 'show save ;x
    
    'If FileExists(cmDialog1.FileName) = True Then
        'calls sub from module

'save2lists(from module), list1, list2, and the file name u choose from common dialog.
Call Save2Lists(frmmain.List1, frmmain.List2, CmDialog1.FileName)

'sets frmmain's list1count.caption to the listcount
frmmain.list1count.Caption = frmmain.List1.ListCount

'sets frmmain's list2count.caption to the listcount
frmmain.list2count.Caption = frmmain.List2.ListCount

'adds horizontal scroll if needed to list1
Call AddHScroll(frmmain.List1)

'adds horizontal scrol if needed to list2
Call AddHScroll(frmmain.List2)
End Sub

Private Sub mnusavetolastopenedlistafteraction_Click()
If mnusavetolastopenedlistafteraction.Checked = True Then
mnusavetolastopenedlistafteraction.Checked = False
Check3.Value = 0
Else
mnusavetolastopenedlistafteraction.Checked = True
Check3.Value = 1
End If
SaveFormState menu
End Sub

Private Sub mnusendaims_Click()
Call chatsend2("" + frmmain.list1count.Caption + " aims made with evil aim sn maker 2")
End Sub

