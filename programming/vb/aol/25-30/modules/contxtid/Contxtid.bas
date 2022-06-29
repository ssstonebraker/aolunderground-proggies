Option Explicit
'=====================================================================
'=====================================================================
'
'This source code contains the following routines:
'  o SetAppHelp() 'Called in the main Form_Load event to register your
'                 'program with WINHELP.EXE
'  o QuitHelp()    'Deregisters your program with WINHELP.EXE. Should
'                  'be called in your main Form_Unload event
'  o HelpWindowSize(x,y,dx,dy) ' Position help window in a screen
'                              ' independent manner
'  o SearchHelp()  'Brings up the windows help KEYWORD SEARCH dialog box
'  o ShowHelpTopic(Topicnum) 'Brings up context sensitive help based on
'                  'any of the following CONTEXT IDs
'***********************************************************************
'
'=====================================================================
'List of Context IDs for <AOKILLA>
'=====================================================================
Global Const Hlp_ThuGGish_ = 140
Global Const Hlp_TOSer_ = 310
Global Const Hlp_IMs_ = 540
Global Const Hlp_MassMailer_ = 560
Global Const Hlp_Annoy_Menu_ = 660
Global Const Hlp_Mail_Bomber_ = 40
Global Const Hlp_Options_ = 220
Global Const Hlp_FiX_Mail_ = 160
Global Const Hlp_Phisher_ = 90
Global Const Hlp_MMBot_ = 50
Global Const Hlp_Mail_ReZeLL_ = 100
Global Const Hlp_Guide_Impersonator_ = 130
Global Const Hlp_TOS_Password_ = 590
Global Const Hlp_Quick_Install_ = 270
Global Const Hlp_Fake_Forward = 80
Global Const Hlp_Macros_ = 680
Global Const Hlp_MaiLBoMBeR_ = 170
Global Const Hlp_Personal_Info = 230
Global Const Hlp_Punt_ = 260
Global Const Hlp_Roll_Dice = 280
Global Const Hlp_many_ = 180
Global Const Hlp_RollHell_ = 290
Global Const Hlp_Drive_Hell = 60
Global Const Hlp_What_do = 300
Global Const Hlp_Aboutxxx = 710
Global Const Hlp_Scroll = 720
Global Const Hlp_Quick_FTP = 740
Global Const Hlp_DxL_Shit = 750
Global Const Hlp_View_Change = 760
Global Const Hlp_Bust_In = 770
'=====================================================================
'
'
'  Help engine section.
Dim m_hWndMainWindow As Integer ' hWnd to tell WINHELP the helpfile owner

' Commands to pass WinHelp()
Global Const HELP_CONTEXT = &H1 '  Display topic in ulTopic
Global Const HELP_QUIT = &H2    '  Terminate help
Global Const HELP_INDEX = &H3   '  Display index
Global Const HELP_HELPONHELP = &H4      '  Display help on using help
Global Const HELP_SETINDEX = &H5        '  Set the current Index for multi index help
Global Const HELP_KEY = &H101           '  Display topic for keyword in offabData
Global Const HELP_MULTIKEY = &H201
Global Const HELP_CONTENTS = &H3     ' Display Help for a particular topic
Global Const HELP_SETCONTENTS = &H5  ' Display Help contents topic
Global Const HELP_CONTEXTPOPUP = &H8 ' Display Help topic in popup window
Global Const HELP_FORCEFILE = &H9    ' Ensure correct Help file is displayed
Global Const HELP_COMMAND = &H102    ' Execute Help macro
Global Const HELP_PARTIALKEY = &H105 ' Display topic found in keyword list
Global Const HELP_SETWINPOS = &H203  ' Display and position Help window

Type HELPWININFO
  wStructSize As Integer
  x As Integer
  y As Integer
  dx As Integer
  dy As Integer
  wMax As Integer
  rgChMember As String * 2
End Type
Declare Function WinHelp Lib "User" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, dwData As Any) As Integer
Declare Function WinHelpByStr Lib "User" Alias "Winhelp" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData$) As Integer
Declare Function WinHelpByNum Lib "User" Alias "Winhelp" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData&) As Integer
Dim MainWindowInfo As HELPWININFO

Sub HelpWindowSize (x As Integer, y As Integer, wx As Integer, wy As Integer)
'=====================================================================
'  TO SET THE SIZE AND POSITION OF THE MAIN HELP WINDOW...
'=====================================================================
'     o   Call HelpWindowSize(x, y, dx, dy), where:
'             x = 1-1024 (position from left edge of screen)
'             y = 1-1024 (position from top of screen)
'             dx= 1-1024 (width)
'             dy= 1-1024 (height)
'
    Dim Result%
    MainWindowInfo.x = x
    MainWindowInfo.y = y
    MainWindowInfo.dx = wx
    MainWindowInfo.dy = wy
    Result% = WinHelp(m_hWndMainWindow, App.HelpFile, HELP_SETWINPOS, MainWindowInfo)
End Sub

Sub quithelp ()
    Dim Result%
    Result% = WinHelp(m_hWndMainWindow, App.HelpFile, HELP_QUIT, Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0))
End Sub

Sub SearchHelp ()
'=====================================================================
'  TO ADD KEYWORD SEARCH CAPABILITY...
'=====================================================================
'     o   In your Help|Search menu selection, simply enter:
'         Call SearchHelp() 'To invoke helpfile keyword search dialog
'
    Dim Result%

    Result% = WinHelp(m_hWndMainWindow, App.HelpFile, HELP_PARTIALKEY, ByVal "")

End Sub

Sub SetAppHelp (ByVal hWndMainWindow)
'=====================================================================
'To use these subroutines to access WINHELP, you need to add
'at least this one subroutine call to your code
'     o  In the Form_Load event of your main Form enter:
'        Call SetAppHelp(Me.hWnd) 'To setup helpfile variables
'         (If you are not interested in keyword searching or context
'         sensitive help, this is the only call you need to make!)
'=====================================================================
    m_hWndMainWindow = hWndMainWindow
    If Right$(Trim$(App.Path), 1) = "\" Then
        App.HelpFile = App.Path + "AOKILLA.HLP"
    Else
        App.HelpFile = App.Path + "\AOKILLA.HLP"
    End If
    MainWindowInfo.wStructSize = 14
    MainWindowInfo.x = 256
    MainWindowInfo.y = 256
    MainWindowInfo.dx = 512
    MainWindowInfo.dy = 512
    MainWindowInfo.rgChMember = Chr$(0) + Chr$(0)
End Sub

Sub ShowHelpContents ()
'=====================================================================
'  DISPLY HELP STARTUP TOPIC IN RESPONSE TO A COMMAND BUTTON or MENU ...
'=====================================================================
'
    Dim Result%

    Result% = WinHelpByNum(m_hWndMainWindow, App.HelpFile, HELP_CONTENTS, CLng(0))

End Sub

Sub ShowHelpTopic (ByVal ContextID As Long)
'=====================================================================
'  FOR CONTEXT SENSITIVE HELP IN RESPONSE TO A COMMAND BUTTON ...
'=====================================================================
'     o   For 'Help button' controls, you can call:
'         Call ShowHelpTopic(<any Hlpxxx entry above>)
'=====================================================================
'  TO ADD FORM LEVEL CONTEXT SENSITIVE HELP...
'=====================================================================
'     o  For FORM level context sensetive help, you should set each
'        Me.HelpContext=<any Hlp_xxx entry above>
'
    Dim Result%

    Result% = WinHelpByNum(m_hWndMainWindow, App.HelpFile, HELP_CONTEXT, CLng(ContextID))

End Sub

Sub ShowPopupHelp (ByVal ContextID As Long)
'=====================================================================
'  FOR POPUP HELP IN RESPONSE TO A COMMAND BUTTON ...
'=====================================================================
    Dim Result%

    Result% = WinHelpByNum(m_hWndMainWindow, App.HelpFile, HELP_CONTEXTPOPUP, CLng(ContextID))

End Sub

