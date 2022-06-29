Attribute VB_Name = "Flyman_mIRC"
'All codes written by flyman
'             III RRRRR      cccc
'              i  R   RR  ccc
'              i  R    R ccc
'mmm      mmm  i  R   R  ccc
'm  m     m m  i  R  R   ccc
'm  m    m  m  i  R R      ccc
'm    m     m III R RRRR     cccc
'
' About: This was intended for personal use, for those who
' like to customize and play with mIRC. Unlike any other
' mIRC bas I have seen uses lots of useless coding that
' you don't need to do, so I showed a easier way with out
' sendkeys and using Win32.API
'
' Author: Lance Seidman(side-man) aka flyman.
' Location: I am from Westlake Village, CA.
' Age: 16 (Nov 4th, 1985).
' URL: http://www.deadbyte.com/flyman
'
' Release Info: This is Flyman_mIRC Release 1
' Release Date: 05/20/02 (May 20th, 2002)
' Release Time: 5:45pm (Pacific Time)

Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Integer, ByVal bRevert As Integer) As Integer
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETTEXT = &H189
Public Const BM_SETCHECK = &HF1
Public Const BM_GETCHECK = &HF0
Public Const MF_BYPOSITION = &H400&

Public Const CB_GETCOUNT = &H146
Public Const CB_GETLBTEXT = &H148
Public Const CB_SETCURSEL = &H14E

Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5


Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const SW_SHOW = 5

Public Const VK_SPACE = &H20

Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOVE = &HF012
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112

Sub mIRC_Channel_Send(txt As String)
'Ex: Call mIRC_Channel_Send("hello")
Dim mIRC As Long
Dim MDIClient As Long
Dim mircchannel As Long
Dim editx As Long

mIRC = FindWindow("mirc", vbNullString)
MDIClient = FindWindowEx(mIRC, 0&, "mdiclient", vbNullString)
mircchannel = FindWindowEx(MDIClient, 0&, "mirc_channel", vbNullString)
editx = FindWindowEx(mircchannel, 0&, "edit", vbNullString)
Call SendMessageByString(editx, WM_SETTEXT, 0&, txt)

Do
    DoEvents
    mIRC = FindWindow("mirc", vbNullString)
    MDIClient = FindWindowEx(mIRC, 0&, "mdiclient", vbNullString)
    mircchannel = FindWindowEx(MDIClient, 0&, "mirc_channel", vbNullString)
    editx = FindWindowEx(mircchannel, 0&, "edit", vbNullString)
    Call SendMessageLong(editx, WM_CHAR, 13, 0&)
Loop Until editx <> 0
End Sub
Sub mIRC_Main_Hide()
'Ex: Call mIRC_Main_Hide
Dim mIRC As Long

mIRC = FindWindow("mirc", vbNullString)
Call ShowWindow(mIRC, SW_HIDE)
End Sub
Sub mIRC_Main_Show()
'Ex: Call mIRC_Main_Show
Dim mIRC As Long

mIRC = FindWindow("mirc", vbNullString)
Call ShowWindow(mIRC, SW_SHOW)
End Sub
Sub mIRC_Main_Change_Caption(txt As String)
'Ex: Call mIRC_Main_Change_Caption ("Testing")
Dim mIRC As Long

mIRC = FindWindow("mirc", vbNullString)
Call SendMessageByString(mIRC, WM_SETTEXT, 0&, txt)
End Sub
Sub mIRC_Channel_Count()
'Ex: Call mIRC_Channel_Count
Dim lCount As Long
Dim mIRC As Long
Dim MDIClient As Long
Dim mircchannel As Long
Dim listbox As Long

mIRC = FindWindow("mirc", vbNullString)
MDIClient = FindWindowEx(mIRC, 0&, "mdiclient", vbNullString)
mircchannel = FindWindowEx(MDIClient, 0&, "mirc_channel", vbNullString)
listbox = FindWindowEx(mircchannel, 0&, "listbox", vbNullString)
lCount = SendMessageLong(listbox, LB_GETCOUNT, 0&, 0&)
MsgBox lCount
End Sub
Sub mIRC_Nicks_ToList()
'Ex: Call mIRC_Nicks_ToList(list1)
Dim mIRC As Long
Dim MDIClient As Long
Dim mircchannel As Long
Dim listbox As Long

mIRC = FindWindow("mirc", vbNullString)
MDIClient = FindWindowEx(mIRC, 0&, "mdiclient", vbNullString)
mircchannel = FindWindowEx(MDIClient, 0&, "mirc_channel", vbNullString)
listbox = FindWindowEx(mircchannel, 0&, "listbox", vbNullString)
Call AddListToListbox(listbox, Form1.List1) ' change to your form name
End Sub
Public Sub AddListToListbox(TheList As Long, NewList As listbox)
' This sub will only work with standard listboxes.
Dim lCount As Long, Item As String, i As Integer, TheNull As Integer
' get the item count in the list
lCount = SendMessageLong(TheList, LB_GETCOUNT, 0&, 0&)
For i = 0 To lCount - 1
Item = String(255, Chr(0))
Call SendMessageByString(TheList, LB_GETTEXT, i, Item)
TheNull = InStr(Item, Chr(0))
' remove any null characters that might be on the end of the string
If TheNull <> 0 Then
NewList.AddItem Mid$(Item, 1, TheNull - 1)
Else
NewList.AddItem Item
End If
Next
End Sub
Sub mIRC_Main_Mini()
'Ex: Call mIRC_Main_Mini
Dim mIRC As Long

mIRC = FindWindow("mirc", vbNullString)
Call ShowWindow(mIRC, SW_MINIMIZE)
End Sub
Sub mIRC_Main_Max()
'Ex: Call mIRC_Main_Max
Dim mIRC As Long

mIRC = FindWindow("mirc", vbNullString)
Call ShowWindow(mIRC, SW_MAXIMIZE)
End Sub
Sub mIRC_Main_Normal()
'Ex: Call mIRC_Main_Normal
Dim mIRC As Long

mIRC = FindWindow("mirc", vbNullString)
Call ShowWindow(mIRC, SW_NORMAL)
End Sub
Sub mIRC_Nick_DNS(nickname As String)
'Ex: Call mIRC_Nick_DNS("flyman")
Call mIRC_Channel_Send("/dns " + nickname)
End Sub
Sub mIRC_Nick_WhoIs(nickname As String)
'Ex: Call mIRC_Nick_WhoIs("flyman")
Call mIRC_Channel_Send("/whois " + nickname)
End Sub
Sub mIRC_Channel_Scroll(tr As String)
'Ex: Call mIRC_Channel_Scroll ("hello")
'Note: Sends to open channel, also can lag Windows IRCDs.
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
Call mIRC_Channel_Send(tr)
End Sub
Sub mIRC_Channel_Ping(channel As String)
'Ex: Call mIRC_Channel_Ping("#flyman_2000")
Call mIRC_Channel_Send("/ping " + channel)
End Sub
Sub mIRC_Nick_Ping(nick As String)
'Ex: Call mIRC_Nick_Ping("flyman")
Call mIRC_Channel_Send("/ping " + nick)
End Sub
Sub mIRC_Channel_Spam(url As String, msg As String)
'Ex: Call mIRC_Channel_Spam("http://www.deadbyte.com/flyman", " - Elite site!")
Call mIRC_Channel_Send(url + msg)
End Sub
Sub mIRC_Channel_Spam_Scroll(url As String, msg As String)
'Ex: Call mIRC_Channel_Spam_Scroll("http://www.deadbyte.com/flyman", " - Elite site!")
Call mIRC_Channel_Send(url + msg)
Call mIRC_Channel_Send(url + msg)
Call mIRC_Channel_Send(url + msg)
Call mIRC_Channel_Send(url + msg)
Call mIRC_Channel_Send(url + msg)
Call mIRC_Channel_Send(url + msg)
Call mIRC_Channel_Send(url + msg)
Call mIRC_Channel_Send(url + msg)
Call mIRC_Channel_Send(url + msg)
Call mIRC_Channel_Send(url + msg)
Call mIRC_Channel_Send(url + msg)
Call mIRC_Channel_Send(url + msg)
Call mIRC_Channel_Send(url + msg)
Call mIRC_Channel_Send(url + msg)
Call mIRC_Channel_Send(url + msg)
Call mIRC_Channel_Send(url + msg)
End Sub
