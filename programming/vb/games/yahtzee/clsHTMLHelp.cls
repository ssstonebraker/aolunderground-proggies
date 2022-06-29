VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HTMLHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ********************************************************
' HTML Help class module version 3.0d
' (c)September 1999, Delmar Computing Services
'
' For use with Microsoft Visual Basic(r)
' versions 4 (32-bit), 5 and 6
'
' Developed by David Liske, Tipton, Michigan, USA
' Microsoft HTML Help MVP 1999
' http://www.vbexplorer.com/htmlhelp.asp
'
' Please send any performance or functionality
' modifications of this file to delmar@tc3net.com
' ________________________________________________________
'
' Proof-of-concept testing and some HTML Help
' API research provided by Robert Chandler,
' The HelpWare Group, and Varian Corporation,
' Melbourne, Vic, Australia
' Microsoft HTML Help MVP 1999
' http://www.helpware.net
'
' Some registry functionality re-developed from
' original code written by Dave Scarmozzino
' http://www.TheScarms.com
'
' Beta testing (July 1999):
' Lani Hardage, MDL Information Systems, Inc.
' Steve Hsu, TREEV, Inc.
' John Hunt, Lotus Development Corporation
' Shirley Kelly, Corbel, A SunGard Company
' Valerie A. Lipow, Compuware Corporation
' Leyden Martinez, Copextel, S.A., Cuba
' Alejandro Sicilia, Copextel, S.A., Cuba
'
' Further testing by Dana Cline,
' Lucent Technologies, Boulder, Colorado
' Microsoft HTML Help MVP 1998, 1999
'
' Modifications of HTML Help version-checking and
' IE version-checking code, and Windows NT/2000
' problem-solving,
' by Leonardo Presciuttini, Italy (August 1999)
'
' GetLongPath_Legacy logic checks and code modifications,
' and all-around support of a programming geek's
' chocolate and strange snacking needs, by:
' Marnella Liske, RN, BSN
' University of Michigan Medical Center
'
' ********************************************************

Option Explicit
Option Compare Text

' Public declarations
Public frm As Object
Public hwnd As Long
Public lpPrevWndFunc As Long

Public Enum PopupType
  HH_CHM_POPUP = &H1
  HH_RESOURCE_POPUP = &H2
  HH_TEXT_POPUP = &H4
End Enum

Public Enum HHVersion
  HH_1_0 = &H10
  HH_1_1 = &H11
  HH_1_1A = &H12
  HH_1_1B = &H13
  HH_1_2 = &H14
  HH_1_21 = &H15
  HH_1_21A = &H16
  HH_1_22 = &H17
  HH_1_3 = &H18
End Enum

Public Enum IEVersion
  IE_3_0 = &H100
  IE_3_0_OSR2 = &H101
  IE_3_01 = &H102
  IE_3_02 = &H103
  IE_4_0_PP2 = &H104
  IE_4_0 = &H105
  IE_4_01 = &H106
  IE_4_01_SP1 = &H107
  IE_4_01_SP2 = &H108
  IE_5_0_Beta1 = &H109
  IE_5_0_Beta2 = &H10A
  IE_5_0 = &H10B
  IE_5_0A = &H10C
  IE_5_0B = &H10D
  IE_5_0C = &H10E
End Enum

' HTML Help Version Constants
Private Const extHH_1_0 = "4.72.7290"
Private Const extHH_1_1 = "4.72.7323"
Private Const extHH_1_1A = "4.72.7325"
Private Const extHH_1_1B = "4.72.8164.0"
Private Const extHH_1_2 = "4.73.8252"
Private Const extHH_1_21 = "4.73.8412"
Private Const extHH_1_21A = "4.73.8474"
Private Const extHH_1_22 = "4.73.8561"
Private Const extHH_1_3 = "4.74.8566"

' Internet Explorer Version Constants
Private Const extIE_3_0 = "4.70.1155"
Private Const extIE_3_0_OSR2 = "4.70.1158"
Private Const extIE_3_01 = "4.70.1215"
Private Const extIE_3_02 = "4.70.1300"
Private Const extIE_4_0_PP2 = "4.71.1008.3"
Private Const extIE_4_0 = "4.71.1712.5"
Private Const extIE_4_01 = "4.72.2106.7"
Private Const extIE_4_01_SP1 = "4.72.3110.03"
Private Const extIE_4_01_SP2 = "4.72.3612.1707"
Private Const extIE_5_0_Beta1 = "5.00.0518.5"
Private Const extIE_5_0_Beta2 = "5.00.0910.1308"
Private Const extIE_5_0 = "5.00.2014.213"
Private Const extIE_5_0A = "5.00.2314.1000"
Private Const extIE_5_0B = "5.00.2614.3500"
Private Const extIE_5_0C = "5.0.2717.2000"

Private Const extUNKNOWN = "unknown"

' HTML Help Constants
Private Const HH_DISPLAY_TOPIC = &H0            '  WinHelp equivalent
Private Const HH_DISPLAY_TOC = &H1              '  WinHelp equivalent
Private Const HH_DISPLAY_INDEX = &H2            '  WinHelp equivalent
Private Const HH_DISPLAY_SEARCH = &H3           '  WinHelp equivalent
Private Const HH_SET_WIN_TYPE = &H4
Private Const HH_GET_WIN_TYPE = &H5
Private Const HH_GET_WIN_HANDLE = &H6
Private Const HH_SYNC = &H9
Private Const HH_ADD_NAV_UI = &HA               ' not currently implemented
Private Const HH_ADD_BUTTON = &HB               ' not currently implemented
Private Const HH_GETBROWSER_APP = &HC           ' not currently implemented
Private Const HH_KEYWORD_LOOKUP = &HD           '  WinHelp equivalent
Private Const HH_DISPLAY_TEXT_POPUP = &HE       ' display string resource id
                                                ' or text in a popup window
                                                ' value in dwData
Private Const HH_HELP_CONTEXT = &HF             '  display mapped numeric
Private Const HH_CLOSE_ALL = &H12               '  WinHelp equivalent
Private Const HH_ALINK_LOOKUP = &H13            '  ALink version of
                                                '  HH_KEYWORD_LOOKUP
Private Const HH_SET_GUID = &H1A                ' For Microsoft Installer -- dwData is a pointer to the GUID string

' HTML Help window constants. These are also used
' in the window definitions in HHP files
Private Const HHWIN_PROP_ONTOP = &H2              ' Top-most window (not currently implemented)
Private Const HHWIN_PROP_NOTITLEBAR = &H4         ' no title bar
Private Const HHWIN_PROP_NODEF_STYLES = &H8       ' no default window styles (only HH_WINTYPE.dwStyles)
Private Const HHWIN_PROP_NODEF_EXSTYLES = &H10    ' no default extended window styles (only HH_WINTYPE.dwExStyles)
Private Const HHWIN_PROP_TRI_PANE = &H20          ' use a tri-pane window
Private Const HHWIN_PROP_NOTB_TEXT = &H40         ' no text on toolbar buttons
Private Const HHWIN_PROP_POST_QUIT = &H80         ' post WM_QUIT message when window closes
Private Const HHWIN_PROP_AUTO_SYNC = &H100        ' automatically ssync contents and index
Private Const HHWIN_PROP_TRACKING = &H200         ' send tracking notification messages
Private Const HHWIN_PROP_TAB_SEARCH = &H400       ' include search tab in navigation pane
Private Const HHWIN_PROP_TAB_HISTORY = &H800      ' include history tab in navigation pane
Private Const HHWIN_PROP_TAB_BOOKMARKS = &H1000   ' include bookmark tab in navigation pane
Private Const HHWIN_PROP_CHANGE_TITLE = &H2000    ' Put current HTML title in title bar
Private Const HHWIN_PROP_NAV_ONLY_WIN = &H4000    ' Only display the navigation window
Private Const HHWIN_PROP_NO_TOOLBAR = &H8000      ' Don't display a toolbar
Private Const HHWIN_PROP_MENU = &H10000           ' Menu
Private Const HHWIN_PROP_TAB_ADVSEARCH = &H20000  ' Advanced FTS UI.
Private Const HHWIN_PROP_USER_POS = &H40000       ' After initial creation, user controls window size/position

Private Const HHWIN_PARAM_PROPERTIES = &H2        ' valid fsWinProperties
Private Const HHWIN_PARAM_STYLES = &H4            ' valid dwStyles
Private Const HHWIN_PARAM_EXSTYLES = &H8          ' valid dwExStyles
Private Const HHWIN_PARAM_RECT = &H10             ' valid rcWindowPos
Private Const HHWIN_PARAM_NAV_WIDTH = &H20        ' valid iNavWidth
Private Const HHWIN_PARAM_SHOWSTATE = &H40        ' valid nShowState
Private Const HHWIN_PARAM_INFOTYPES = &H80        ' valid apInfoTypes
Private Const HHWIN_PARAM_TB_FLAGS = &H100        ' valid fsToolBarFlags
Private Const HHWIN_PARAM_EXPANSION = &H200       ' valid fNotExpanded
Private Const HHWIN_PARAM_TABPOS = &H400          ' valid tabpos
Private Const HHWIN_PARAM_TABORDER = &H800        ' valid taborder
Private Const HHWIN_PARAM_HISTORY_COUNT = &H1000  ' valid cHistory
Private Const HHWIN_PARAM_CUR_TAB = &H2000        ' valid curNavType

Private Const HHWIN_BUTTON_EXPAND = &H2           ' Expand/contract button
Private Const HHWIN_BUTTON_BACK = &H4             ' Back button
Private Const HHWIN_BUTTON_FORWARD = &H8          ' Forward button
Private Const HHWIN_BUTTON_STOP = &H10            ' Stop button
Private Const HHWIN_BUTTON_REFRESH = &H20         ' Refresh button
Private Const HHWIN_BUTTON_HOME = &H40            ' Home button
Private Const HHWIN_BUTTON_BROWSE_FWD = &H80      ' not implemented
Private Const HHWIN_BUTTON_BROWSE_BCK = &H100     ' not implemented
Private Const HHWIN_BUTTON_NOTES = &H200          ' not implemented
Private Const HHWIN_BUTTON_CONTENTS = &H400       ' not implemented
Private Const HHWIN_BUTTON_SYNC = &H800           ' Locate button
Private Const HHWIN_BUTTON_OPTIONS = &H1000       ' Options button
Private Const HHWIN_BUTTON_PRINT = &H2000         ' Print button
Private Const HHWIN_BUTTON_INDEX = &H4000         ' not implemented
Private Const HHWIN_BUTTON_SEARCH = &H8000        ' not implemented
Private Const HHWIN_BUTTON_HISTORY = &H10000      ' not implemented
Private Const HHWIN_BUTTON_BOOKMARKS = &H20000    ' not implemented
Private Const HHWIN_BUTTON_JUMP1 = &H40000        ' Jump1 button
Private Const HHWIN_BUTTON_JUMP2 = &H80000        ' Jump2 button
Private Const HHWIN_BUTTON_ZOOM = &H100000        ' Font sizing button
Private Const HHWIN_BUTTON_TOC_NEXT = &H200000    ' Browse next TOC topic button
Private Const HHWIN_BUTTON_TOC_PREV = &H400000    ' Browse previous TOC topic button

' Default button set
Private Const HHWIN_DEF_BUTTONS = _
            (HHWIN_BUTTON_EXPAND Or _
             HHWIN_BUTTON_BACK Or _
             HHWIN_BUTTON_OPTIONS Or _
             HHWIN_BUTTON_PRINT)

' Button IDs
Private Const IDTB_EXPAND = 200
Private Const IDTB_CONTRACT = 201
Private Const IDTB_STOP = 202
Private Const IDTB_REFRESH = 203
Private Const IDTB_BACK = 204
Private Const IDTB_HOME = 205
Private Const IDTB_SYNC = 206
Private Const IDTB_PRINT = 207
Private Const IDTB_OPTIONS = 208
Private Const IDTB_FORWARD = 209
Private Const IDTB_NOTES = 210             ' not implemented
Private Const IDTB_BROWSE_FWD = 211
Private Const IDTB_BROWSE_BACK = 212
Private Const IDTB_CONTENTS = 213          ' not implemented
Private Const IDTB_INDEX = 214             ' not implemented
Private Const IDTB_SEARCH = 215            ' not implemented
Private Const IDTB_HISTORY = 216           ' not implemented
Private Const IDTB_BOOKMARKS = 217         ' not implemented
Private Const IDTB_JUMP1 = 218
Private Const IDTB_JUMP2 = 219
Private Const IDTB_CUSTOMIZE = 221
Private Const IDTB_ZOOM = 222
Private Const IDTB_TOC_NEXT = 223
Private Const IDTB_TOC_PREV = 224

Private Enum HHACT_
  HHACT_TAB_CONTENTS
  HHACT_TAB_INDEX
  HHACT_TAB_SEARCH
  HHACT_TAB_HISTORY
  HHACT_TAB_FAVORITES
    
  HHACT_EXPAND
  HHACT_CONTRACT
  HHACT_BACK
  HHACT_FORWARD
  HHACT_STOP
  HHACT_REFRESH
  HHACT_HOME
  HHACT_SYNC
  HHACT_OPTIONS
  HHACT_PRINT
  HHACT_HIGHLIGHT
  HHACT_CUSTOMIZE
  HHACT_JUMP1
  HHACT_JUMP2
  HHACT_ZOOM
  HHACT_TOC_NEXT
  HHACT_TOC_PREV
  HHACT_NOTES

  HHACT_LAST_ENUM
End Enum

Private Enum HHWIN_NAVTYPE_
  HHWIN_NAVTYPE_TOC
  HHWIN_NAVTYPE_INDEX
  HHWIN_NAVTYPE_SEARCH
  HHWIN_NAVTYPE_HISTORY       ' not implemented
  HHWIN_NAVTYPE_FAVORITES     ' not implemented
End Enum

Enum HHWIN_NAVTAB_
  HHWIN_NAVTAB_TOP
  HHWIN_NAVTAB_LEFT
  HHWIN_NAVTAB_BOTTOM
End Enum

Private Const HH_MAX_TABS = 19               ' maximum number of tabs

Private Enum HH_TAB_
  HH_TAB_CONTENTS
  HH_TAB_INDEX
  HH_TAB_SEARCH
  HH_TAB_HISTORY
  HH_TAB_FAVORITES
End Enum

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type tagHH_WINTYPE
  cbStruct As Long            ' IN: size of this structure including all Information Types
  fUniCodeStrings As Long     ' IN/OUT: TRUE if all strings are in UNICODE
  pszType  As String          ' IN/OUT: Name of a type of window
  fsValidMembers As Long      ' IN: Bit flag of valid members (HHWIN_PARAM_)
  fsWinProperties As Long     ' IN/OUT: Properties/attributes of the window (HHWIN_)
  pszCaption As String        ' IN/OUT: Window title
  dwStyles  As Long           ' IN/OUT: Window styles
  dwExStyles As Long          ' IN/OUT: Extended Window styles
  rcWindowPos As RECT         ' IN: Starting position, OUT: current position
  nShowState As Long          ' IN: show state (e.g., SW_SHOW)
  hwndHelp As Long            ' OUT: window handle
  hwndCaller As Long          ' OUT: who called this window
  paInfoTypes As Long         ' IN: Pointer to an array of Information Types

  ' The following members are only valid if HHWIN_PROP_TRI_PANE is set

  hwndToolBar As Long         ' OUT: toolbar window in tri-pane window
  hwndNavigation As Long      ' OUT: navigation window in tri-pane window
  hwndHTML As Long            ' OUT: window displaying HTML in tri-pane window
  iNavWidth As Long           ' IN/OUT: width of navigation window
  rcHTML As RECT              ' OUT: HTML window coordinates

  pszToc As String            ' IN: Location of the table of contents file
  pszIndex As String          ' IN: Location of the index file
  pszFile As String           ' IN: Default location of the html file
  pszHome As String           ' IN/OUT: html file to display when Home button is clicked
  fsToolBarFlags As Long      ' IN: flags controling the appearance of the toolbar
  fNotExpanded As Long        ' IN: TRUE/FALSE to contract or expand, OUT: current state
  curNavType As Long          ' IN/OUT: UI to display in the navigational pane
  tabpos As HHWIN_NAVTAB_     ' IN/OUT: HHWIN_NAVTAB_TOP, HHWIN_NAVTAB_LEFT, or HHWIN_NAVTAB_BOTTOM
  idNotify As Long            ' IN: ID to use for WM_NOTIFY messages
  tabOrder(HH_MAX_TABS) As Byte ' IN/OUT: tab order: Contents, Index, Search, History, Favorites, Reserved 1-5, Custom tabs
  cHistory As Long            ' IN/OUT: number of history items to keep (default is 30)
  pszJump1 As String          ' Text for HHWIN_BUTTON_JUMP1
  pszJump2 As String          ' Text for HHWIN_BUTTON_JUMP2
  pszUrlJump1 As String       ' URL for HHWIN_BUTTON_JUMP1
  pszUrlJump2 As String       ' URL for HHWIN_BUTTON_JUMP2
  rcMinSize As RECT           ' Minimum size for window (ignored in version 1)
  cbInfoTypes As Long         ' size of paInfoTypes;
End Type

' UDT for mouse cursor position
Private Type POINTAPI
  x As Long
  y As Long
End Type

' UDT for text popups
Private Type tagHH_POPUP
  cbStruct As Integer                         ' sizeof this structure
  hinst As Long                               ' instance handle for string resource
  idString As Long                            ' string resource id, or text id if pszFile
                                              ' is specified in HtmlHelp call
  pszText As String                           ' used if idString is zero
  pt As POINTAPI                              ' top center of popup window
  clrForeground As ColorConstants             ' either use VB constant or &HBBGGRR
  clrBackground As ColorConstants             ' either use VB constant or &HBBGGRR
  rcMargins As RECT                           ' amount of space between edges of window and
                                              ' text, -1 for each member to ignore
  pszFont As String                           ' facename, point size, char set, BOLD ITALIC
                                              ' UNDERLINE
End Type

' UDT for keyword and ALink searches
Private Type tagHH_AKLINK
  cbStruct          As Long
  fReserved         As Boolean
  pszKeywords       As String
  pszUrl            As String
  pszMsgText        As String
  pszMsgTitle       As String
  pszWindow         As String
  fIndexOnFail      As Boolean
End Type

' UDT for accessing the Search tab
Private Type tagHH_FTS_QUERY
  cbStruct          As Long
  fUniCodeStrings   As Long
  pszSearchQuery    As String
  iProximity        As Long
  fStemmedSearch    As Long
  fTitleOnly        As Long
  fExecute          As Long
  pszWindow         As String
End Type

' Constants for converting the cursor to What's This Help
Private Const WM_SYSCOMMAND = &H112
Private Const SC_CONTEXTHELP = &HF180&
                                               
' Message Box Constants
Private Const MB_ABORTRETRYIGNORE = &H2&
Private Const MB_APPLMODAL = &H0&
Private Const MB_COMPOSITE = &H2
Private Const MB_DEFAULT_DESKTOP_ONLY = &H20000
Private Const MB_DEFBUTTON1 = &H0&
Private Const MB_DEFBUTTON2 = &H100&
Private Const MB_DEFBUTTON3 = &H200&
Private Const MB_DEFMASK = &HF00&
Private Const MB_ICONASTERISK = &H40&
Private Const MB_ICONEXCLAMATION = &H30&
Private Const MB_ICONHAND = &H10&
Private Const MB_ICONINFORMATION = MB_ICONASTERISK
Private Const MB_ICONMASK = &HF0&
Private Const MB_ICONQUESTION = &H20&
Private Const MB_ICONSTOP = MB_ICONHAND
Private Const MB_MISCMASK = &HC000&
Private Const MB_MODEMASK = &H3000&
Private Const MB_NOFOCUS = &H8000&
Private Const MB_OK = &H0&
Private Const MB_OKCANCEL = &H1&
Private Const MB_PRECOMPOSED = &H1
Private Const MB_RETRYCANCEL = &H5&
Private Const MB_SETFOREGROUND = &H10000
Private Const MB_SYSTEMMODAL = &H1000&
Private Const MB_TASKMODAL = &H2000&
Private Const MB_TYPEMASK = &HF&
Private Const MB_USEGLYPHCHARS = &H4
Private Const MB_YESNO = &H4&
Private Const MB_YESNOCANCEL = &H3&

' Registry API call Constants
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const ERROR_SUCCESS = 0&
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const SYNCHRONIZE = &H100000
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Private Const MAX_PATH = 260
Private Const INVALID_HANDLE_VALUE = -1

' RegCreateKeyEx options
Private Const REG_OPTION_NON_VOLATILE = 0

' Registry data types
Private Const REG_NONE = 0
Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4

' FindFirstFile return values
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_MORE_DATA = 234
Private Const ERROR_NO_MORE_ITEMS = 259&

' Constants for Registry top-level keys
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_USERS = &H80000003
Private Const HKEY_DYN_DATA = &H80000006
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_CLASSES_ROOT = &H80000000

Private Const MAX_SIZE = 2048
Private Const MAX_INISIZE = 8192

Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const GWL_WNDPROC = (-4)

' Constants for GetLongPath_Legacy
Private Const SINGLE_QUOTE = """"

' Constants for determining OS
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Const UNKNOWN_OS = 0
Private Const WINDOWS_NT_3_51 = 1
Private Const WINDOWS_95 = 2
Private Const WINDOWS_NT_4 = 3
Private Const WINDOWS_98 = 4
Private Const WINDOWS_2000 = 5

' UDT for determining OS
Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type

' UDT for message box API calls
Private Type MSGBOXPARAMS
  cbSize As Long
  hWndOwner As Long
  hInstance As Long
  lpszText As String
  lpszCaption As String
  dwStyle As Long
  lpszIcon As String
  dwContextHelpId As Long
  lpfnMsgBoxCallback As Long
  dwLanguageId As Long
End Type

' Registry UDT's
Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type

Private Type HH_REG_VALUES
  pszFileName     As String
  pszFilePath     As String
End Type

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
  dwFileAttributes As Long
  ftCreationTime As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime As FILETIME
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * MAX_PATH
  cAlternate As String * 14
End Type

Private Type VS_FIXEDFILEINFO
  dwSignature As Long
  dwStrucVersionl As Integer     '  e.g. = &h0000 = 0
  dwStrucVersionh As Integer     '  e.g. = &h0042 = .42
  dwFileVersionMSl As Integer    '  e.g. = &h0003 = 3
  dwFileVersionMSh As Integer    '  e.g. = &h0075 = .75
  dwFileVersionLSl As Integer    '  e.g. = &h0000 = 0
  dwFileVersionLSh As Integer    '  e.g. = &h0031 = .31
  dwProductVersionMSl As Integer '  e.g. = &h0003 = 3
  dwProductVersionMSh As Integer '  e.g. = &h0010 = .1
  dwProductVersionLSl As Integer '  e.g. = &h0000 = 0
  dwProductVersionLSh As Integer '  e.g. = &h0031 = .31
  dwFileFlagsMask As Long        '  = &h3F for version "0.42"
  dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
  dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
  dwFileType As Long             '  e.g. VFT_DRIVER
  dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
  dwFileDateMS As Long           '  e.g. 0
  dwFileDateLS As Long           '  e.g. 0
End Type

' HTML Help API declarations
Private Declare Function HTMLHelp Lib "hhctrl.ocx" _
    Alias "HtmlHelpA" (ByVal hwnd As Long, _
    ByVal lpHelpFile As String, _
    ByVal wCommand As Long, _
    ByVal dwData As Long) As Long
    
Private Declare Function HTMLHelpCallSearch Lib "hhctrl.ocx" _
    Alias "HtmlHelpA" (ByVal hwnd As Long, _
    ByVal lpHelpFile As String, _
    ByVal wCommand As Long, _
    ByRef dwData As tagHH_FTS_QUERY) As Long
    
Private Declare Function HTMLHelpKeyWord Lib "hhctrl.ocx" _
    Alias "HtmlHelpA" (ByVal hwnd As Long, _
    ByVal lpHelpFile As String, _
    ByVal wCommand As Long, _
    dwData As tagHH_AKLINK) As Long
    
Private Declare Function htmlHelpTextPopup Lib "hhctrl.ocx" _
    Alias "HtmlHelpA" (ByVal hwnd As Long, _
    ByVal lpHelpFile As String, _
    ByVal wCommand As Long, _
    ByRef dwData As tagHH_POPUP) As Long
    
Private Declare Function htmlHelpTopic Lib "hhctrl.ocx" _
    Alias "HtmlHelpA" (ByVal hwnd As Long, _
    ByVal lpHelpFile As String, _
    ByVal wCommand As Long, _
    ByVal dwData As String) As Long
    
' Subclassing API declarations
Private Declare Function DefWindowProc Lib "user32" _
    Alias "DefWindowProcA" (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

' Registry API declarations
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" _
    Alias "RegCreateKeyExA" (ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    ByVal Reserved As Long, _
    ByVal lpClass As String, _
    ByVal dwOptions As Long, _
    ByVal samDesired As Long, _
    lpSecurityAttributes As SECURITY_ATTRIBUTES, _
    phkResult As Long, _
    lpdwDisposition As Long) As Long
    
Private Declare Function ExpandEnvironmentStrings _
    Lib "kernel32" Alias "ExpandEnvironmentStringsA" _
    (ByVal lpSrc As String, _
    ByVal lpDst As String, _
    ByVal nSize As Long) As Long
        
Private Declare Function RegCloseKey Lib "advapi32.dll" _
    (ByVal hKey As Long) As Long
        
Private Declare Function RegDeleteValue Lib "advapi32.dll" _
    Alias "RegDeleteValueA" _
    (ByVal hKey As Long, _
    ByVal lpValueName As String) As Long

Private Declare Function RegEnumKey Lib "advapi32.dll" _
    Alias "RegEnumKeyA" _
    (ByVal hKey As Long, _
    ByVal dwIndex As Long, _
    ByVal lpName As String, _
    ByVal cbName As Long) As Long

Private Declare Function RegEnumValue Lib "advapi32.dll" _
    Alias "RegEnumValueA" _
    (ByVal hKey As Long, _
    ByVal dwIndex As Long, _
    ByVal lpValueName As String, _
    lpcbValueName As Long, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    lpData As Any, _
    lpcbData As Long) As Long
    
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" _
    Alias "RegOpenKeyExA" (ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    ByVal ulOptions As Long, _
    ByVal samDesired As Long, _
    phkResult As Long) As Long

Private Declare Function RegQueryValue Lib "advapi32.dll" _
    Alias "RegQueryValueA" _
    (ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    ByVal lpValue As String, _
    lpcbValue As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" _
    Alias "RegQueryValueExA" (ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, lpData As Any, _
    lpcbData As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32.dll" _
    Alias "RegSetValueExA" _
    (ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal Reserved As Long, _
    ByVal dwType As Long, _
    lpData As Any, _
    ByVal cbData As Long) As Long

' Calls to find actual file
Private Declare Function FindClose Lib "kernel32" _
    (ByVal hFindFile As Long) As Long
    
Private Declare Function FindFirstFile Lib "kernel32" _
    Alias "FindFirstFileA" (ByVal lpFileName As String, _
    lpFindFileData As WIN32_FIND_DATA) As Long

' Declarations to retrieve version information
Private Declare Function GetFileVersionInfo& _
    Lib "version.dll" Alias "GetFileVersionInfoA" _
    (ByVal lptstrFilename As String, _
    ByVal dwHandle As Long, _
    ByVal dwLen As Long, _
    lpData As Byte)
        
Private Declare Function GetFileVersionInfoSize& _
    Lib "version.dll" Alias "GetFileVersionInfoSizeA" _
    (ByVal lptstrFilename As String, _
    lpdwHandle As Long)

Private Declare Function VerQueryValue& Lib "version.dll" _
    Alias "VerQueryValueA" _
    (pBlock As Byte, _
    ByVal lpSubBlock As String, _
    lpBuffer As Long, _
    puLen As Long)

' Declaration to copy memory contents from one area to another
Private Declare Sub CopyMem Lib "kernel32" _
    Alias "RtlMoveMemory" _
    (Destination As Any, _
    Source As Any, _
    ByVal Length As Long)

' Call to get the current mouse position
Private Declare Function GetCursorPos& Lib "user32" _
    (lpPoint As POINTAPI)

' Call to translate short file path to long file path
' (Win98 and Win2k and above only - see comments for the
' GetLongFilePath_Legacy procedure)
Private Declare Function GetLongPathName Lib "kernel32" _
    (ByRef pszShortPath As String, _
    ByRef lpszLongPath As String, _
    ByVal cchBuffer As Long) As Long

' Call to determine OS version
Private Declare Function GetVersionExA Lib "kernel32" _
    (lpVersionInformation As OSVERSIONINFO) As Integer

' Message box API declaration
Private Declare Function MessageBoxIndirect Lib "user32" _
    Alias "MessageBoxIndirectA" _
    (lpMsgBoxParams As MSGBOXPARAMS) As Long
      
'local variable(s) to hold property value(s)
Private mvarCHMFile As String
Private mvarHHALink As String
Private mvarHHDefaultURL As String
Private mvarHHInstalled As Boolean
Private mvarHHKeyword As String
Private mvarHHMsgText As String
Private mvarHHMsgTitle As String
Private mvarHHShowOnTop As Boolean
Private mvarHHTopicID As Long
Private mvarHHTopicURL As String
Private mvarHHWindow As String
Private mvarHHRegFileName As String
Private mvarHHRegFilePath As String
Private mvarHHRegFileExists As Boolean
Private mvarHHPopupFile As String
Private mvarHHPopupType As PopupType
Private mvarHHPopupText As String
Private mvarHHPopupID As Long
Private mvarHHPopupTextColor As Long
Private mvarHHPopupBackColor As Long
Private mvarHHPopupCustomTextColor As Long
Private mvarHHPopupCustomBackColor As Long
Private mvarHHPopupCustomColors As Boolean
Private mvarHHPopupTextFont As String
Private mvarHHPopupTextSize As String
Private mvarHHPopupTextBold As Boolean
Private mvarHHPopupTextItalic As Boolean
Private mvarHHPopupTextUnderline As Boolean
Private mvarHHCtrlPath As String
Private mvarHHVersion As String
Private mvarIEVersion As String
Private mvarHHFriendlyName As String
Private mvarIEFriendlyName As String

' Module-level variables
Private strHTMLHelpPath As String
Private strWindow As String
Private strTopic As String
Private lngTopicID As Long

Public Sub HHDisplayTopicURL(Optional ByRef CallingForm As Long)

' Displays a specific topic via the HHTopicURL property

  On Error GoTo ErrHandler
    
  Dim hwnd As Long
    
  If ValidHHFile(mvarCHMFile) = False Then
    Exit Sub
  End If
    
  If EnsureFileExists(mvarCHMFile) = False Then
    Exit Sub
  End If
    
  If Trim(mvarHHWindow) = "" Then
    If mvarHHShowOnTop Then
      htmlHelpTopic CallingForm, mvarCHMFile, _
          HH_DISPLAY_TOPIC, mvarHHTopicURL
    Else
      htmlHelpTopic hwnd, mvarCHMFile, _
          HH_DISPLAY_TOPIC, mvarHHTopicURL
    End If
  Else
    If mvarHHShowOnTop Then
      htmlHelpTopic CallingForm, mvarCHMFile & ">" & mvarHHWindow, _
          HH_DISPLAY_TOPIC, mvarHHTopicURL
    Else
      htmlHelpTopic hwnd, mvarCHMFile & ">" & mvarHHWindow, _
          HH_DISPLAY_TOPIC, mvarHHTopicURL
    End If
  End If
    
  Exit Sub

ErrHandler:
    
  Select Case Err.Number
  Case 91
    MessageBoxExclamation "The HHDisplayContents method was called " & _
        "without a form being specified, while HHShowOnTop " & _
        "was set to True."
    Exit Sub
  Case Else
      Resume Next
  End Select
      
End Sub

Public Sub HHDisplayTopicID(Optional CallingForm As Long)

' Displays a specific topic via the HHTopicID property
    
  On Error GoTo ErrHandler
    
  Dim hwnd As Long
    
  If ValidHHFile(mvarCHMFile) = False Then
    Exit Sub
  End If
    
  If EnsureFileExists(mvarCHMFile) = False Then
    Exit Sub
  End If
  
  If Trim(mvarHHWindow) = "" Then
    If mvarHHShowOnTop Then
      HTMLHelp CallingForm, mvarCHMFile, _
          HH_HELP_CONTEXT, mvarHHTopicID
    Else
      HTMLHelp hwnd, mvarCHMFile, _
          HH_HELP_CONTEXT, mvarHHTopicID
    End If
  Else
    If mvarHHShowOnTop Then
      HTMLHelp CallingForm, mvarCHMFile & ">" & mvarHHWindow, _
          HH_HELP_CONTEXT, mvarHHTopicID
    Else
      HTMLHelp hwnd, mvarCHMFile & ">" & mvarHHWindow, _
          HH_HELP_CONTEXT, mvarHHTopicID
    End If
  End If
    
  Exit Sub

ErrHandler:
    
  Select Case Err.Number
  Case 91
    MessageBoxExclamation "The HHDisplayContents method was called " & _
        "without a form being specified, while HHShowOnTop " & _
        "was set to True."
    Exit Sub
  Case Else
    Resume Next
  End Select
      
End Sub

Public Sub HHDisplaySearch(Optional ByRef CallingForm As Long)

' Forces the Help window to display the Search tab

  On Error GoTo ErrHandler
    
  Dim hwnd As Long
    
  If ValidHHFile(mvarCHMFile) = False Then
    Exit Sub
  End If
    
  If EnsureFileExists(mvarCHMFile) = False Then
    Exit Sub
  End If
    
  Dim HH_FTS_QUERY As tagHH_FTS_QUERY

  With HH_FTS_QUERY
    .cbStruct = Len(HH_FTS_QUERY)
    .fStemmedSearch = 0&
    .fTitleOnly = 0&
    .fUniCodeStrings = 0&
    .iProximity = 0&
    .pszSearchQuery = ""
    .pszWindow = ""
    .fExecute = 1&
  End With
    
  If Trim(mvarHHWindow) = "" Then
    If mvarHHShowOnTop Then
      HTMLHelpCallSearch CallingForm, _
          mvarCHMFile, _
          HH_DISPLAY_SEARCH, HH_FTS_QUERY
    Else
      HTMLHelpCallSearch hwnd, _
          mvarCHMFile, _
          HH_DISPLAY_SEARCH, HH_FTS_QUERY
    End If
  Else
    If mvarHHShowOnTop Then
      HTMLHelpCallSearch CallingForm, _
          mvarCHMFile & ">" & mvarHHWindow, _
          HH_DISPLAY_SEARCH, HH_FTS_QUERY
    Else
      HTMLHelpCallSearch hwnd, _
          mvarCHMFile & ">" & mvarHHWindow, _
          HH_DISPLAY_SEARCH, HH_FTS_QUERY
    End If
  End If
  Exit Sub

ErrHandler:
    
  Select Case Err.Number
  Case 91
    MessageBoxExclamation "The HHDisplayContents method was called " & _
        "without a form being specified, while HHShowOnTop " & _
        "was set to True."
    Exit Sub
  Case Else
    Resume Next
  End Select
      
End Sub

Public Sub HHDisplayKeyword(Optional ByRef CallingForm As Long)

' Displays a topic specified by the HHKeyword property.
' This will search for a KLink keyword in the topics
' themselves.  Also searches the entries of an Index
' file (*.hhk) used in a tripane window.

  On Error GoTo ErrHandler
    
  Dim hwnd As Long
    
  If ValidHHFile(mvarCHMFile) = False Then
    Exit Sub
  End If
    
  If EnsureFileExists(mvarCHMFile) = False Then
    Exit Sub
  End If
    
  If mvarHHShowOnTop Then
    htmlHelpTopic CallingForm, mvarCHMFile, _
        HH_DISPLAY_TOPIC, vbNullString
  Else
    htmlHelpTopic hwnd, mvarCHMFile, _
        HH_DISPLAY_TOPIC, vbNullString
  End If
    
  Dim ALinkStruct As tagHH_AKLINK
  
  ALinkStruct.cbStruct = Len(ALinkStruct)
  ALinkStruct.fReserved = False
  ALinkStruct.pszKeywords = mvarHHKeyword
    
  ' Translate empty strings to Null strings
  If mvarHHDefaultURL = "" Then
    mvarHHDefaultURL = vbNullString
  End If
    
  If mvarHHMsgText = "" Then
    mvarHHMsgText = vbNullString
  End If
    
  If mvarHHMsgTitle = "" Then
    mvarHHMsgTitle = vbNullString
  End If
    
  ' Set up the default topic to use if the
  ' specified keyword is not found.  This is
  ' set via the HHDefaultURL property.
  ALinkStruct.pszUrl = mvarHHDefaultURL
    
  ' Set up the message box to display if the
  ' specified keyword is not found.  These are
  ' set via the HHMsgText and HHMshgTitle properties.
  ALinkStruct.pszMsgText = mvarHHMsgText
  ALinkStruct.pszMsgTitle = mvarHHMsgTitle
    
  ' Use the HHWindow property if it's set.
  If Trim(mvarHHWindow) <> "" Then
    ALinkStruct.pszWindow = mvarHHWindow
  End If
    
  ' Set to False to enable the default URL
  ' and message box functions.
  ALinkStruct.fIndexOnFail = False
    
  If mvarHHShowOnTop Then
    HTMLHelpKeyWord CallingForm, mvarCHMFile, _
        HH_KEYWORD_LOOKUP, ALinkStruct
  Else
    HTMLHelpKeyWord hwnd, mvarCHMFile, _
        HH_KEYWORD_LOOKUP, ALinkStruct
  End If
    
  Exit Sub

ErrHandler:
    
  Select Case Err.Number
  Case 91
    MessageBoxExclamation "The HHDisplayContents method was called " & _
        "without a form being specified, while HHShowOnTop " & _
        "was set to True."
    Exit Sub
  Case Else
    Resume Next
  End Select
      
End Sub

Public Sub HHDisplayIndex(Optional ByRef CallingForm As Long)

' Force the Help window to display the Index file
' (*.hhk) in the left pane

  On Error GoTo ErrHandler
    
  Dim hwnd As Long
    
  If ValidHHFile(mvarCHMFile) = False Then
    Exit Sub
  End If
    
  If EnsureFileExists(mvarCHMFile) = False Then
    Exit Sub
  End If
  
  If Trim(mvarHHWindow) = "" Then
    If mvarHHShowOnTop Then
      HTMLHelp CallingForm, mvarCHMFile, _
          HH_DISPLAY_INDEX, 0
    Else
      HTMLHelp hwnd, mvarCHMFile, _
          HH_DISPLAY_INDEX, 0
    End If
  Else
    If mvarHHShowOnTop Then
      HTMLHelp CallingForm, mvarCHMFile & ">" & mvarHHWindow, _
          HH_DISPLAY_INDEX, 0
    Else
      HTMLHelp hwnd, mvarCHMFile & ">" & mvarHHWindow, _
          HH_DISPLAY_INDEX, 0
    End If
  End If
    
  Exit Sub

ErrHandler:
    
  Select Case Err.Number
  Case 91
    MessageBoxExclamation "The HHDisplayContents method was called " & _
        "without a form being specified, while HHShowOnTop " & _
        "was set to True."
    Exit Sub
  Case Else
    Resume Next
  End Select
      
End Sub

Public Sub HHDisplayContents(Optional ByRef CallingForm As Long)
    
' Force the Help window to display the Contents file
' (*.hhc) in the left pane
    
  On Error GoTo ErrHandler
    
  Dim hwnd As Long
    
  If ValidHHFile(mvarCHMFile) = False Then
    Exit Sub
  End If
    
  If EnsureFileExists(mvarCHMFile) = False Then
    Exit Sub
  End If
    
  If Trim(mvarHHWindow) = "" Then
    If mvarHHShowOnTop Then
      HTMLHelp CallingForm, mvarCHMFile, _
          HH_DISPLAY_TOC, 0
    Else
      HTMLHelp hwnd, mvarCHMFile, _
          HH_DISPLAY_TOC, 0
    End If
  Else
    If mvarHHShowOnTop Then
      HTMLHelp CallingForm, mvarCHMFile & ">" & mvarHHWindow, _
          HH_DISPLAY_TOC, 0
    Else
      HTMLHelp hwnd, mvarCHMFile & ">" & mvarHHWindow, _
          HH_DISPLAY_TOC, 0
    End If
  End If
    
  Exit Sub

ErrHandler:
    
  Select Case Err.Number
  Case 91
    MessageBoxExclamation "The HHDisplayContents method was called " & _
        "without a form being specified, while HHShowOnTop " & _
        "was set to True."
    Exit Sub
  Case Else
    Resume Next
  End Select
      
End Sub

Public Sub HHDisplayALink(Optional ByRef CallingForm As Long)

' Displays a topic specified by the HHALink property.

  On Error GoTo ErrHandler
    
  Dim hwnd As Long
    
  If ValidHHFile(mvarCHMFile) = False Then
    Exit Sub
  End If
    
  If EnsureFileExists(mvarCHMFile) = False Then
    Exit Sub
  End If

  Dim ALinkStruct As tagHH_AKLINK
  
  ALinkStruct.cbStruct = Len(ALinkStruct)
  ALinkStruct.fReserved = False
  ALinkStruct.pszKeywords = mvarHHALink
    
  ' Translate empty strings to Null strings
  If mvarHHDefaultURL = "" Then
    mvarHHDefaultURL = vbNullString
  End If
    
  If mvarHHMsgText = "" Then
    mvarHHMsgText = vbNullString
  End If
    
  If mvarHHMsgTitle = "" Then
    mvarHHMsgTitle = vbNullString
  End If
    
  ' Set up the default topic to use if the
  ' specified keyword is not found.  This is
  ' set via the HHDefaultURL property.
  ALinkStruct.pszUrl = mvarHHDefaultURL
    
  ' Set up the message box to display if the
  ' specified keyword is not found.  These are
  ' set via the HHMsgText and HHMshgTitle properties.
  ALinkStruct.pszMsgText = mvarHHMsgText
  ALinkStruct.pszMsgTitle = mvarHHMsgTitle
    
  ' Use the HHWindow property if it's set.
  If Trim(mvarHHWindow) <> "" Then
    ALinkStruct.pszWindow = mvarHHWindow
  End If
    
  ' Set to False to enable the default URL
  ' and message box functions.
  ALinkStruct.fIndexOnFail = False
  
  If mvarHHShowOnTop Then
    HTMLHelpKeyWord CallingForm, mvarCHMFile, _
        HH_ALINK_LOOKUP, ALinkStruct
  Else
    HTMLHelpKeyWord hwnd, mvarCHMFile, _
        HH_ALINK_LOOKUP, ALinkStruct
  End If
    
  Exit Sub

ErrHandler:
    
  Select Case Err.Number
  Case 91
    MessageBoxExclamation "The HHDisplayContents method was called " & _
        "without a form being specified, while HHShowOnTop " & _
        "was set to True."
    Exit Sub
  Case Else
    Resume Next
  End Select
      
End Sub

Public Sub HHRegister(FileToRegister As String)

' Registers the specified HTML Help file in
' HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\HTML Help

  Dim lngHandle As Long
  Dim lngDisposition As Long
  Dim lngLenData As Long
  Dim lngResult As Long
  Dim strValue As String
  Dim secSecAttributes As SECURITY_ATTRIBUTES
  Dim strFilePath As String
  Dim intPosition As Integer
  Dim intLength As Integer
  
  If FileToRegister = "" Then Exit Sub
  
  If ValidHHFile(FileToRegister) = False Then Exit Sub
  
  GetLongPath_Legacy FileToRegister, FileToRegister, 256
  
  HHCheckRegistry FileToRegister
    
  If (mvarHHRegFileName <> "") Then
      
    ' The file is registered to begin with,
    ' so we need to say so and exit.
    MessageBoxInformation "The file " & FileToRegister & " is already registered.  " & _
        "HHRegister will not be run as no action need be taken."
    Exit Sub
      
  End If
  
  ' Copy it to get the path later
  strFilePath = FileToRegister
  
  If ValidHHFile(FileToRegister) = False Then
    Exit Sub
  Else
  
    If EnsureFileExists(FileToRegister) = False Then
      Exit Sub
    End If
    If InStr(FileToRegister, "\") = 0 Then
      MessageBoxExclamation "Cannot register " & FileToRegister & " without having a supplied path."""
      Exit Sub
    Else
      ' strip the file name itself off the path
      intPosition = 1
    
      Do While intPosition <> 0
        intLength = Len(FileToRegister)
        intPosition = InStr(1, FileToRegister, "\")
        FileToRegister = Right(FileToRegister, (intLength - intPosition))
      Loop
    
      ' Get the registered path
      strFilePath = Left(strFilePath, Len(strFilePath) - Len(FileToRegister))
    
      ' Register it
      lngResult = 99

      lngResult = RegCreateKeyEx(HKEY_LOCAL_MACHINE, _
          "Software\Microsoft\Windows\HTML Help", _
          0, _
          "", _
          REG_OPTION_NON_VOLATILE, _
          KEY_CREATE_SUB_KEY Or KEY_SET_VALUE, _
          secSecAttributes, _
          lngHandle, _
          lngDisposition)
      
      If lngResult <> ERROR_SUCCESS Then
        GoTo WriteRegValueError
      End If
    
      strValue = strFilePath
      lngLenData = Len(strValue) + 1
      lngResult = RegSetValueEx(lngHandle, _
          FileToRegister, _
          0, _
          REG_SZ, _
          ByVal strValue, _
          lngLenData)
    
      If lngResult = ERROR_SUCCESS Then
        lngResult = RegCloseKey(lngHandle)
        Exit Sub
      End If
      
    End If
  End If
  
  HHCheckRegistry CStr(FileToRegister)
  
  Exit Sub

WriteRegValueError:
  MessageBoxExclamation "HTML Help Class: HHRegister Error"

End Sub

Public Sub HHUnRegister(FileToUnRegister As String)

' Deletes the entry for the specified HTML Help file from
' HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\HTML Help

  Dim lngResult As Long
  Dim strFilePath As String
  Dim intPosition As Integer
  Dim intLength As Integer
  Dim lngHandle As Long
  
  If FileToUnRegister = "" Then Exit Sub
  
  If ValidHHFile(FileToUnRegister) = False Then Exit Sub
    
  GetLongPath_Legacy FileToUnRegister, FileToUnRegister, 256
  
  HHCheckRegistry FileToUnRegister
    
  If (mvarHHRegFileName = "") Then
      
    ' The file isn't registered to begin with,
    ' so we need to say so and exit.
    MessageBoxInformation "The file " & FileToUnRegister & " is not registered.  " & _
        "HHUnRegister will not be run as no action need be taken."
      
    Exit Sub
      
  End If
    
  ' strip the file name itself off the path
  intPosition = 1
    
  Do While intPosition <> 0
    intLength = Len(FileToUnRegister)
    intPosition = InStr(1, FileToUnRegister, "\")
    FileToUnRegister = Right(FileToUnRegister, (intLength - intPosition))
  Loop
    
  ' Delete the entry
  lngResult = RegOpenKeyEx(HKEY_LOCAL_MACHINE, _
      "Software\Microsoft\Windows\HTML Help", _
      0, _
      KEY_SET_VALUE, _
      lngHandle)
      
  If lngResult = ERROR_SUCCESS Then
    lngResult = RegDeleteValue(lngHandle, FileToUnRegister)
  End If
  
  HHCheckRegistry CStr(FileToUnRegister)
  
End Sub

Public Function HHCheckRegistry(ByRef FileToCheck As String) As Variant

' Verifies the specified HTML Help file has been registered in
' HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\HTML Help

  Dim sValues As String
  Dim intLength As Integer
  Dim boolResult As Boolean
  Dim intPosition As Integer
  Dim intFileNamePos As Integer
  Dim varRegValues As HH_REG_VALUES
  Dim strTempFileName As String
  
  ' Reassign the file name so the original
  ' doesn't become mangled
  strTempFileName = FileToCheck
  
  If strTempFileName = "" Then
    Exit Function
  End If
  
  HHCheckRegistry = False
  mvarHHRegFileExists = False
  
  If ValidHHFile(strTempFileName) = False Then
    Exit Function
  Else
    ' strip the file name itself off the path
    intPosition = 1
    
    Do While intPosition <> 0
      intLength = Len(strTempFileName)
      intPosition = InStr(1, strTempFileName, "\")
      strTempFileName = Right(strTempFileName, (intLength - intPosition))
    Loop
    
    ' Verify it
    varRegValues = EnumRegValue("HKLM", "Software\Microsoft\Windows\HTML Help", sValues, strTempFileName)
    
    With varRegValues
      If .pszFileName = strTempFileName Then
        
        ' Load the verified file name into the HHRegFileName
        ' property, and the registered path of the file into
        ' the HHRegFilePath property.
        Dim tempHHRegFilePath As String
        mvarHHRegFileName = .pszFileName
        tempHHRegFilePath = Right(.pszFilePath, Len(.pszFilePath) - InStr(.pszFilePath, "="))
        tempHHRegFilePath = Left(tempHHRegFilePath, (Len(tempHHRegFilePath) - 2))
        
        GetLongPath_Legacy tempHHRegFilePath, mvarHHRegFilePath, 256
        
        Dim lpFindFileData As WIN32_FIND_DATA
        Dim lngHandle As Long
        
        ' Verify the HTML Help file exists according to the
        ' registry data.
        boolResult = EnsureFileExists(mvarHHRegFilePath & "\" & mvarHHRegFileName)
        
        If boolResult Then
        
          ' If the file's in the registered location,
          ' set HHRegFileExists to True.
          mvarHHRegFileExists = True
          
        End If
      Else
        
        ' If the file isn't registered, clear the
        ' HHRegFileName and HHRegFilePath properties.
        mvarHHRegFileName = ""
        mvarHHRegFilePath = ""
        
      End If
    End With
    
  End If
  
End Function

Public Sub HHClose()
  
' Closes all open HTML Help windows.  Be careful when
' using this as it closes everything, not just the
' HTML Help windows for this app.
  
  On Error Resume Next
    
  Dim hwnd As Long

  HTMLHelp hwnd, mvarCHMFile, HH_CLOSE_ALL, 0
    
End Sub

Public Function HHSetHelpFile(ByVal intSelHelpFile As Integer) As String
  
' Set the string variable to
' include the application path

  Select Case intSelHelpFile
  Case 1
    HHSetHelpFile = App.Path & "\Library.chm"
  Case 2
    ' Popup text file for the above CHM
    HHSetHelpFile = "Library.txt"
  Case Else
    ' list other HTML Help files here
  End Select
  
End Function

Public Sub HHDisplayPopup(ByRef CallingForm As Long)

' Displays a text popup from any of three sources
' and with a number of options
  
  If mvarHHPopupType = 0 Then
    Exit Sub
  End If

  Dim pPoint As POINTAPI
  Dim hPopup As tagHH_POPUP
  Dim rRect As RECT
  Dim hwnd As Long
  Dim strFontString As String
  
  If mvarHHPopupType = HH_CHM_POPUP Then
    
    ' Check for a valid CHM if that option is selected
    
    If ValidHHFile(mvarCHMFile) = False Then
      Exit Sub
    End If
    
    If EnsureFileExists(mvarCHMFile) = False Then
      Exit Sub
    End If
    
  End If
  
  ' Get the current mouse pointer position
  GetCursorPos& pPoint
  
  ' Set the margins of the popup
  With rRect
    .Bottom = -1
    .Left = -1
    .Right = -1
    .Top = -1
  End With
  
  With hPopup
    .cbStruct = Len(hPopup)
  
    ' Clear any previously used color scheme
    .clrForeground = 0
    .clrBackground = 0
    
    If mvarHHPopupType = HH_RESOURCE_POPUP Then
    
      ' Fudge the Resource function by actually
      ' using HH_TEXT_POPUP
      .pszText = LoadResString(mvarHHPopupID)
      .hinst = 0
      
    ElseIf mvarHHPopupType = HH_TEXT_POPUP Then
    
      .idString = 0
      
      ' End the sub if no string is supplied
      If mvarHHPopupText = "" Then
        Exit Sub
      End If
      .pszText = mvarHHPopupText
      
    Else
      
      .idString = mvarHHPopupID
      .pszText = vbNullString
      
    End If
    
    ' Match the popup coordinates to the current
    ' mouse pointer coordinates
    .pt = pPoint
    
    If mvarHHPopupCustomColors = True Then
      .clrForeground = mvarHHPopupCustomTextColor
      .clrBackground = mvarHHPopupCustomBackColor
    Else
      .clrForeground = mvarHHPopupTextColor
      .clrBackground = mvarHHPopupBackColor
    End If
    
    .rcMargins = rRect
    
    Dim strBold As String
    Dim strItalic As String
    Dim strUnderline As String
    
    If mvarHHPopupTextFont = "" Then
      mvarHHPopupTextFont = "Arial"
    End If
    
    If mvarHHPopupTextSize = "" Then
      mvarHHPopupTextSize = "10"
    End If
    
    If mvarHHPopupTextBold = True Then
      strBold = "Bold "
    Else
      strBold = ""
    End If
    
    If mvarHHPopupTextItalic = True Then
      strItalic = "Italic "
    Else
      strItalic = ""
    End If
    
    If mvarHHPopupTextUnderline = True Then
      strUnderline = "Underline"
    Else
      strUnderline = ""
    End If
    
    strFontString = mvarHHPopupTextFont & ", " & _
        mvarHHPopupTextSize & ", ascii, " & _
        strBold & strItalic & strUnderline
        
    .pszFont = strFontString
    
  End With
  
  Select Case mvarHHPopupType
  Case HH_CHM_POPUP
  
    If ValidPopupFile(mvarHHPopupFile) = False Then
      Exit Sub
    Else
      htmlHelpTextPopup CallingForm, _
        mvarCHMFile & "::/" & mvarHHPopupFile, _
        HH_DISPLAY_TEXT_POPUP, hPopup
    End If
    
  Case HH_RESOURCE_POPUP
    htmlHelpTextPopup CallingForm, vbNullString, _
      HH_DISPLAY_TEXT_POPUP, hPopup
    
  Case HH_TEXT_POPUP
    htmlHelpTextPopup CallingForm, vbNullString, _
      HH_DISPLAY_TEXT_POPUP, hPopup
      
  Case Else
    Exit Sub
  End Select
  
  ' Clear the color and font scheme before the next use
  mvarHHPopupCustomTextColor = 0
  mvarHHPopupCustomBackColor = 0
  mvarHHPopupTextColor = 0
  mvarHHPopupBackColor = 0
  mvarHHPopupTextFont = "Arial"
  mvarHHPopupTextSize = "10"
  mvarHHPopupTextBold = False
  mvarHHPopupTextItalic = False
  mvarHHPopupTextUnderline = False
  
End Sub

Public Sub HHInvokeWhatsThisHelp(CallingForm As Long)
  
' Sets the mouse pointer to the What's This pointer.
' When the left button is clicked again, What's This
' is invoked for the control below the cursor.

  DefWindowProc CallingForm, WM_SYSCOMMAND, _
      SC_CONTEXTHELP, 0

End Sub

Public Function HHVerifyMinConfig(MinHHVersion As HHVersion, _
    MinIEVersion As IEVersion) As Boolean
  
' Verifies the minimum HTML Help and IE versions
' as specified by the developer.
    
  Dim boolHHVerified As Boolean
  Dim boolIEVerified As Boolean
  
  GetHHVersion
  
  GetIEVersion
  
  HHVerifyMinConfig = False
  
  Select Case MinHHVersion
  Case HH_1_0
    If mvarHHVersion >= extHH_1_0 Then
      boolHHVerified = True
    End If
  Case HH_1_1
    If mvarHHVersion >= extHH_1_1 Then
      boolHHVerified = True
    End If
  Case HH_1_1A
    If mvarHHVersion >= extHH_1_1A Then
      boolHHVerified = True
    End If
  Case HH_1_1B
    If mvarHHVersion >= extHH_1_1B Then
      boolHHVerified = True
    End If
  Case HH_1_2
    If mvarHHVersion >= extHH_1_2 Then
      boolHHVerified = True
    End If
  Case HH_1_21
    If mvarHHVersion >= extHH_1_21 Then
      boolHHVerified = True
    End If
  Case HH_1_21A
    If mvarHHVersion >= extHH_1_21A Then
      boolHHVerified = True
    End If
  Case HH_1_22
    If mvarHHVersion >= extHH_1_22 Then
      boolHHVerified = True
    End If
  Case HH_1_22
    If mvarHHVersion >= extHH_1_3 Then
      boolHHVerified = True
    End If
  Case Else
    mvarHHVersion = extUNKNOWN
  End Select

  Select Case MinIEVersion
  Case IE_3_0
    If mvarIEVersion >= extIE_3_0 Then
      boolIEVerified = True
    End If
  Case IE_3_0_OSR2
    If mvarIEVersion >= extIE_3_0_OSR2 Then
      boolIEVerified = True
    End If
  Case IE_3_01
    If mvarIEVersion >= extIE_3_01 Then
      boolIEVerified = True
    End If
  Case IE_3_02
    If mvarIEVersion >= extIE_3_02 Then
      boolIEVerified = True
    End If
  Case IE_4_0_PP2
    If mvarIEVersion >= extIE_4_0_PP2 Then
      boolIEVerified = True
    End If
  Case IE_4_0
    If mvarIEVersion >= extIE_4_0 Then
      boolIEVerified = True
    End If
  Case IE_4_01
    If mvarIEVersion >= extIE_4_01 Then
      boolIEVerified = True
    End If
  Case IE_4_01_SP1
    If mvarIEVersion >= extIE_4_01_SP1 Then
      boolIEVerified = True
    End If
  Case IE_4_01_SP2
    If mvarIEVersion >= extIE_4_01_SP2 Then
      boolIEVerified = True
    End If
  Case IE_5_0_Beta1
    If mvarIEVersion >= extIE_5_0_Beta1 Then
      boolIEVerified = True
    End If
  Case IE_5_0_Beta2
    If mvarIEVersion >= extIE_5_0_Beta2 Then
      boolIEVerified = True
    End If
  Case IE_5_0
    If mvarIEVersion >= extIE_5_0 Then
      boolIEVerified = True
    End If
  Case IE_5_0A
    If mvarIEVersion >= extIE_5_0A Then
      boolIEVerified = True
    End If
  Case IE_5_0B
    If mvarIEVersion >= extIE_5_0B Then
      boolIEVerified = True
    End If
  Case IE_5_0C
    If mvarIEVersion >= extIE_5_0C Then
      boolIEVerified = True
    End If
  Case Else
    mvarIEVersion = extUNKNOWN
  End Select
  
  GetHHFriendlyName
  GetIEFriendlyName
  
  HHVerifyMinConfig = (boolHHVerified And boolIEVerified)
  
End Function

Private Function GetVersionInfo(FileName) As String

' Retrieves the file version information for
' the specified file
  
  Dim varVersionSize As Long
  Dim varVersionHwnd As Long
  Dim bytVerBuf() As Byte
  Dim lngResult As Long
  Dim ffi As VS_FIXEDFILEINFO
  Dim ffiAddr As Long
  Dim ffiLen As Long
  Dim di As Long
  
  varVersionSize = GetFileVersionInfoSize(FileName, varVersionHwnd)
  If varVersionSize > 64000 Then varVersionSize = 64000
  
  ReDim bytVerBuf(varVersionSize + 1)
  
  lngResult = GetFileVersionInfo(FileName, varVersionHwnd, varVersionSize, bytVerBuf(0))
  di = VerQueryValue(bytVerBuf(0), "\", ffiAddr, ffiLen)
  
  CopyMem ffi, ByVal ffiAddr, Len(ffi)
  
  GetVersionInfo = Format$(ffi.dwFileVersionMSh) & "." & _
      Format$(ffi.dwFileVersionMSl, "00") & "."
      
  If ffi.dwFileVersionLSh > 0 Then
    GetVersionInfo = GetVersionInfo & Format$(ffi.dwFileVersionLSh, "00") & "." & _
        Format$(ffi.dwFileVersionLSl, "00")
  Else
    GetVersionInfo = GetVersionInfo & Format$(ffi.dwFileVersionLSl, "0000")
  End If
  
End Function

Private Function ValidHHFile(FileToVerify) As Boolean

' Verifies the suffix for the specified CHM
' Note this procedure does not verify actual CHM
  
  ValidHHFile = True
  
  If Right(FileToVerify, 3) <> "chm" Then
    MessageBoxExclamation FileToVerify & " is not a valid HTML Help file."
    Exit Function
  End If

End Function

Private Function ValidPopupFile(FileToVerify) As Boolean

' Verifies the suffix for the specified popup text file
' Note this procedure does not verify actual popup text file
  
  ValidPopupFile = True
  
  If Right(FileToVerify, 3) <> "txt" Then
    MessageBoxExclamation "The file specified as the text popup source, '" & _
        FileToVerify & "', is not a valid popup file."
    Exit Function
  End If

End Function

Public Function EnsureFileExists(ByRef FileToFind As String) As Boolean

' Ensures the specified file exists in its specified location

  Dim lpFindFileData As WIN32_FIND_DATA
  Dim lngHandle As Long
  
  EnsureFileExists = True
  
  If (InStr(FileToFind, "\") = 0) Then
    ' Check to see if the file is registered
    HHCheckRegistry FileToFind
    
    If mvarHHRegFileName <> "" Then
      ' If it's registered, use the registry info
      FileToFind = mvarHHRegFilePath & mvarHHRegFileName
    Else
      ' Otherwise, assume it's in App.Path
      FileToFind = App.Path & "\" & FileToFind
    End If
  End If
  
  lngHandle = FindFirstFile(FileToFind, lpFindFileData)
  
  If (lngHandle) = INVALID_HANDLE_VALUE Then
        
    MessageBoxExclamation "The file " & FileToFind & " does not exist." & Chr(10) & _
        "Please make sure the correct path and file name have been specified."
        
    EnsureFileExists = False
    
  Else
  
    FindClose lngHandle
    
  End If
  
End Function

Private Function EnumRegValue(ByVal strTopKey As String, _
    ByVal strSubKey As String, _
    strValues As String, _
    FileToCheck As String) As HH_REG_VALUES

' Enumerates registry values

  Dim strTempFileName As String
  Dim lngTopKey As Long
  Dim lngHandle As Long
  Dim lngResult As Long
  Dim lngValueLen As Long
  Dim lngIndex As Long
  Dim lngValue As Long
  Dim lngValueType As Long
  Dim lngData As Long
  Dim lngDataLen As Long
  Dim boolDone As Boolean
  Dim strValueName As String
  Dim strValue As String
  Dim strValueEx As String
  Dim HHRegValues As HH_REG_VALUES

  On Error GoTo EnumRegValueError

  ' Reassign the file name so the original
  ' doesn't become mangled
  strTempFileName = FileToCheck

  ' Clear any previous result
  EnumRegValue.pszFileName = ""
  EnumRegValue.pszFilePath = ""

  lngResult = 99
  lngTopKey = RegKeyID(strTopKey)
  
  If lngTopKey = 0 Then GoTo EnumRegValueError

  lngResult = RegOpenKeyEx(lngTopKey, strSubKey, 0, _
      KEY_QUERY_VALUE, lngHandle)
  If lngResult <> ERROR_SUCCESS Then GoTo EnumRegValueError
    
  Do While Not boolDone
        
    lngDataLen = MAX_SIZE
    lngValueLen = lngDataLen
    strValueName = Space$(lngDataLen)
        
    lngResult = RegEnumValue(lngHandle, lngIndex, _
        strValueName, lngValueLen, 0, lngValueType, _
        ByVal lngData, lngDataLen)
        
    If lngResult = ERROR_SUCCESS Then
    
      Select Case lngValueType
      Case REG_SZ, REG_EXPAND_SZ
        strValue = Space$(lngDataLen)
        strValueName = Left$(strValueName, lngValueLen)
        
        If strValueName = strTempFileName Then
          HHRegValues.pszFileName = strTempFileName
          lngResult = RegQueryValueEx(lngHandle, _
              strValueName, 0, lngValueType, _
              ByVal strValue, lngDataLen)
              
          If lngValueType = REG_EXPAND_SZ Then
            strValueEx = strValue
            strValue = String(MAX_SIZE, " ")
            lngValueLen = ExpandEnvironmentStrings(strValueEx, strValue, MAX_SIZE)
          End If
          
          If lngResult = ERROR_SUCCESS Then
            strValues = strValues & strValueName & _
                "=" & strValue & vbCr
            HHRegValues.pszFilePath = strValues
                            
          Else
            GoTo EnumRegValueError
          End If
          
          GoTo ExitRoutine
          
        End If
                    
      Case Else
      End Select
      
      lngIndex = lngIndex + 1
      
    Else
      boolDone = True
      
    End If
  Loop

ExitRoutine:
  strValues = strValues & vbCr
  If Len(strValues) = 1 Then strValues = strValues & vbCr
    
  lngResult = RegCloseKey(lngHandle)
  EnumRegValue = HHRegValues
    
  ' Clear any previous arguments
  strTopKey = ""
  strSubKey = ""
  strValues = ""
  strTempFileName = ""
    
  Exit Function

EnumRegValueError:
  EnumRegValue = HHRegValues
  
End Function

Public Function GetKeyInfo(ByVal key_name As String, _
    ByVal indent As Integer) As Boolean

' Used with HHInstalled method to verify whether or not
' HTML Help is installed on the system.

  Dim subkeys As Collection
  Dim subkey_values As Collection
  Dim subkey_num As Integer
  Dim subkey_name As String
  Dim subkey_value As String
  Dim Length As Long
  Dim hKey As Long
  Dim txt As String
    
  GetKeyInfo = True
  
  Set subkeys = New Collection
  Set subkey_values = New Collection
    
  If RegOpenKeyEx(HKEY_LOCAL_MACHINE, _
      key_name, _
      0&, KEY_ALL_ACCESS, hKey) <> ERROR_SUCCESS _
      Then
    GetKeyInfo = False
    Exit Function
  End If
    
  subkey_num = 0
  Do
    Length = 256
    subkey_name = Space$(Length)
    
    If RegEnumKey(hKey, subkey_num, _
        subkey_name, Length) _
        <> ERROR_SUCCESS Then Exit Do
        
    subkey_num = subkey_num + 1
        
    subkey_name = Left$(subkey_name, InStr(subkey_name, Chr$(0)) - 1)
    subkeys.Add subkey_name
    
    Length = 256
    subkey_value = Space$(Length)
    If RegQueryValue(hKey, subkey_name, _
        subkey_value, Length) _
        <> ERROR_SUCCESS _
        Then
      subkey_values.Add "Error"
    Else
      subkey_value = Left$(subkey_value, Length - 1)
      subkey_values.Add subkey_value
      End If
  Loop
    
  If RegCloseKey(hKey) <> ERROR_SUCCESS Then
    GetKeyInfo = False
  End If
    
  For subkey_num = 1 To subkeys.Count
    txt = txt & subkeys(subkey_num) & _
        ": " & subkey_values(subkey_num) & _
        vbCrLf
  Next subkey_num
    
End Function

Private Function RegKeyID(ByVal strTopKeyOrFile As String) As Long

' Translates the registry key constants

  Dim strDir As String

  RegKeyID = 0
  Select Case UCase$(strTopKeyOrFile)
  Case "HKCU"
    RegKeyID = HKEY_CURRENT_USER
  Case "HKLM"
    RegKeyID = HKEY_LOCAL_MACHINE
  Case "HKU"
    RegKeyID = HKEY_USERS
  Case "HKDD"
    RegKeyID = HKEY_DYN_DATA
  Case "HKCC"
    RegKeyID = HKEY_CURRENT_CONFIG
  Case "HKCR"
    RegKeyID = HKEY_CLASSES_ROOT
  Case Else
    On Error Resume Next
    strDir = Dir$(strTopKeyOrFile)
    If Err.Number = 0 And strDir <> "" Then RegKeyID = 1
  End Select
  Exit Function
  
End Function

Private Function GetLongPath_Legacy(ByVal strShortName As String, _
    strLongName As String, lngBufferLength As Long) As Long

' Translates a short path name into its actual long path
' name.  On Windows 98 and 2000 systems and later, uses
' the GetLongPathName API call in kernel32.  This is used
' on these systems as it renders a slightly faster
' processing time. GetLongPath_Legacy is set up in the
' following manner to make it compatible with the
' GetLongPathName API call:
'
' strShortName:
'     Pointer to a null-terminated path to be converted.
' strLongName:
'     Pointer to the buffer to receive the long path.
' lngBufferLength:
'     Specifies the size of the buffer, in characters.

' If the function succeeds, the return value is the
' length of the string copied to the strLongName
' parameter, in characters. This length does not include
' the terminating null character.
'
' If strLongName is too small on Windows 98 and 2000
' systems and later, the function returns the size of the
' buffer required to hold the long path, in characters.
' This is due to the use of the GetLongPathName API call
' on those operating systems.
'
' If the function fails, the return value is zero.

  On Error GoTo ErrHandler
  
  Dim strTempLongName As String
  Dim strTemp As String
  Dim strPathTemp As String
  Dim intLength As Integer
  Dim intPosition As Integer
  Dim intStart As Integer
  Dim lngHandle As Long
  Dim lpFindFileData As WIN32_FIND_DATA
  
  If GetWindowsVersion >= WINDOWS_98 Then
    ' For Windows 98 and later, and Windows 2000
    ' and later, use the following API call:
    GetLongPath_Legacy = GetLongPathName(strShortName, _
        strLongName, lngBufferLength)
    Exit Function
  End If
  
  GetLongPath_Legacy = 0
  
  ' If stored in the registry, in some cases it's
  ' enclosed in double-quotes, so we need to delete them
  If Left$(strShortName, 1) = SINGLE_QUOTE Then _
      strShortName = Right$(strShortName, Len(strShortName) - 1)
  If Right$(strShortName, 1) = SINGLE_QUOTE Then _
      strShortName = Left$(strShortName, Len(strShortName) - 1)
      
  ' Add \ to short name to prevent Instr from failing
  If Right$(strShortName, 1) <> "\" Then _
      strShortName = Left$(strShortName, Len(strShortName) & "\")
  
  ' Save the drive letter for later
  strLongName = Left(strShortName, 2)
  
  ' Strip the drive letter off the temporary string
  strPathTemp = Right(strShortName, Len(strShortName) - 3)
  
  ' Find the first backslash
  intPosition = InStr(strPathTemp, "\")
  
  Do While intPosition <> 0
    
    ' Get the individual component of the path name
    strTemp = Left(strPathTemp, intPosition - 1)
    
    ' Translate the short path component into its
    ' actual long path component as found by FindFirstFile
    lngHandle = FindFirstFile(strLongName & "\" & strTemp, lpFindFileData)
    
    If (lngHandle) = INVALID_HANDLE_VALUE Then
    
      ' The folder or file does not exist
      Exit Function
    
    End If
    
    ' Get rid of any null characters retrieved if
    ' from the registry
    strTempLongName = StripNulls(lpFindFileData.cFileName)
    
    ' Build the long path name, starting with the
    ' previously-saved drive letter
    strLongName = strLongName & "\" & strTempLongName
    
    ' Delete the short path component we just used
    strPathTemp = Right(strPathTemp, Len(strPathTemp) - (Len(strTemp) + 1))
    
    ' Find the next backslash
    intPosition = InStr(strPathTemp, "\")
    
  Loop
  
  ' Add the remainder, which is the name of the file
  strLongName = strLongName & "\" & strPathTemp
  
  GetLongPath_Legacy = Len(strLongName)

  Exit Function
  
ErrHandler:
  Select Case Err.Number
  Case 52
    strLongName = ""
    Exit Function
  Case Else
    Resume Next
  End Select

End Function

Private Function StripNulls(OriginalStr As String) As String

' Strips any trailing nulls from path names retrieved
' from the registry. This function is found in the
' following Microsoft(r) knowledge base articles:
' Q183009 "HOWTO: Enumerate Windows Using the WIN32 API"
' Q185476 "HOWTO: Search Directories to Find or List Files"
' Q190218 "HOWTO: Retrieve Settings From a Printer Driver"

  If (InStr(OriginalStr, Chr(0)) > 0) Then
    OriginalStr = Left(OriginalStr, _
        InStr(OriginalStr, Chr(0)) - 1)
  End If
  
  StripNulls = OriginalStr
  
End Function

Private Sub GetHHVersion()

' Retrieves the version of HTML Help on the system
  
  Dim varHHRegValues As HH_REG_VALUES
  Dim lngResult As Boolean
  Dim sHHValues As String
  Dim strTempHHctrlPath As String
  Dim strHHctrlPath As String
  
  ' Get the path of the registered copy of hhctrl.ocx
  varHHRegValues = EnumRegValue("HKCR", "CLSID\{4662DAB0-D393-11D0-9A56-00C04FB68B66}\InprocServer32", sHHValues, "")
  
  ' If hhctrl.ocx isn't registered, go past the next block.
  If Len(varHHRegValues.pszFilePath) = 0 Then
    MessageBoxCritical "Hhctrl.ocx is not registered.  " & _
        "Please install HTML Help."
    
  Else
    strHHctrlPath = Right(varHHRegValues.pszFilePath, (Len(varHHRegValues.pszFilePath) - 1))
  
    Dim lpFindFileData As WIN32_FIND_DATA
    Dim lngHandle As Long
    
    strTempHHctrlPath = StripNulls(strHHctrlPath)
    
    ' Translate short path name if registered that way
    lngResult = GetLongPath_Legacy(strTempHHctrlPath, strHHctrlPath, 256)
        
    ' Verify the HTML Help control exists according
    ' to the registry data.
    lngHandle = FindFirstFile(strHHctrlPath, lpFindFileData)
        
    If (lngHandle) = INVALID_HANDLE_VALUE Then
      MessageBoxCritical "Hhctrl.ocx is not in its registered location.  " & _
          "Please reinstall HTML Help."
      Exit Sub
    
    End If
  
    mvarHHVersion = GetVersionInfo(strHHctrlPath)
  
  End If

End Sub

Private Sub GetIEVersion()

' Retrieves the version of Internet Explorer on the system

  Dim varIERegValues As HH_REG_VALUES
  Dim sIEValues As String
  Dim strShdocvwPath As String
  Dim lngResult As Boolean
  Dim strIexploreTempPath As String
  Dim lngHandle As Long
  Dim lpFindFileData As WIN32_FIND_DATA

  ' Get the path of the registered copy of shdocvw.dll
  varIERegValues = EnumRegValue("HKCR", "CLSID\{0A89A860-D7B1-11CE-8350-444553540000}\InProcServer32", sIEValues, "")
  
  ' If shdocvw.dll isn't registered, go past the next block.
  If Len(varIERegValues.pszFilePath) = 0 Then
    MessageBoxCritical "Shdocvw.dll is not registered.  " & _
        "Please install Internet Explorer."
  Else
  
    strShdocvwPath = Right(varIERegValues.pszFilePath, (Len(varIERegValues.pszFilePath) - 1))
    
    ' Translate short path name if registered that way
    lngResult = GetLongPath_Legacy(StripNulls(strShdocvwPath), strShdocvwPath, 256)
    
    ' Verify shdocvw.dll exists according
    ' to the registry data.
    lngHandle = FindFirstFile(strShdocvwPath, lpFindFileData)
    
    If (lngHandle) = INVALID_HANDLE_VALUE Then
      MessageBoxCritical "Internet Explorer is not in its registered location.  " & _
          "Please reinstall Internet Explorer."
      Exit Sub
    
    End If
  
    mvarIEVersion = GetVersionInfo(strShdocvwPath)
    
  End If
  
End Sub

Private Function GetHHFriendlyName() As String

' Retrieves the friendly name of HTML Help as installed

  If mvarHHVersion = "" Then
    GetHHVersion
  End If
  ' Take into account how this module
  ' returns version numbers
  Select Case mvarHHVersion
  Case "4.72.7290.00"
    mvarHHFriendlyName = "1.0"
  Case "4.72.7323.00"
    mvarHHFriendlyName = "1.1"
  Case "4.72.7325.00"
    mvarHHFriendlyName = "1.1a"
  Case "4.72.8164.00"
    mvarHHFriendlyName = "1.1b"
  Case "4.73.8252.00"
    mvarHHFriendlyName = "1.2"
  Case "4.73.8412.00"
    mvarHHFriendlyName = "1.21"
  Case "4.73.8474.00"
    mvarHHFriendlyName = "1.21a"
  Case "4.73.8561.00"
    mvarHHFriendlyName = "1.22"
  Case "4.74.8566.00"
    mvarHHFriendlyName = "1.3"
  Case Else
    mvarHHFriendlyName = "unknown"
  End Select
  
End Function

Private Function GetIEFriendlyName() As String

' Retrieves the friendly name of Internet Explorer
' as installed
  
  If mvarIEVersion = "" Then
    GetIEVersion
  End If
  
  ' Take into account how this module
  ' returns version numbers
  Select Case mvarIEVersion
  Case "4.70.1155.0000"
    mvarIEFriendlyName = "3.0"
  Case "4.70.1158.0000"
    mvarIEFriendlyName = "3.0 OSR2"
  Case "4.70.1215.0000"
    mvarIEFriendlyName = "3.02"
  Case "4.70.1300.0000"
    mvarIEFriendlyName = "3.02"
  Case "4.71.1008.3000"
    mvarIEFriendlyName = "4.0 PP2"
  Case "4.71.1712.5000"
    mvarIEFriendlyName = "4.0"
  Case "4.72.2106.7000"
    mvarIEFriendlyName = "4.01"
  Case "4.72.3110.0300"
    mvarIEFriendlyName = "4.0 SP1"
  Case "4.72.3612.1707"
    mvarIEFriendlyName = "4.0 SP2"
  Case "5.00.0518.5000"
    mvarIEFriendlyName = "5.0 beta 1"
  Case "5.00.0910.1308"
    mvarIEFriendlyName = "5.0 beta 2"
  Case "5.00.2014.2130"
    mvarIEFriendlyName = "5.0"
  Case "5.00.2314.1000"
    mvarIEFriendlyName = "5.0a (Office 2000)"
  Case "5.00.2614.3500"
    mvarIEFriendlyName = "5.0b (Windows 98 SE)"
  Case "5.00.2717.2000"
    mvarIEFriendlyName = "5.0c (Icon Security Issue Update)"
  Case Else
    mvarIEFriendlyName = "unknown"
  End Select
  
End Function

Public Function GetWindowsVersion() As Long
            
  Dim osinfo As OSVERSIONINFO
  Dim retvalue As Integer

  osinfo.dwOSVersionInfoSize = 148
  osinfo.szCSDVersion = Space$(128)
  retvalue = GetVersionExA(osinfo)

  With osinfo
    Select Case .dwPlatformId
    Case VER_PLATFORM_WIN32_WINDOWS
      If .dwMinorVersion = 0 Then
        GetWindowsVersion = WINDOWS_95
      ElseIf .dwMinorVersion = 10 Then
        GetWindowsVersion = WINDOWS_98
      End If
    Case VER_PLATFORM_WIN32_NT
      If .dwMajorVersion = 3 Then
        GetWindowsVersion = WINDOWS_NT_3_51
      ElseIf .dwMajorVersion = 4 Then
        GetWindowsVersion = WINDOWS_NT_4
      ElseIf .dwMajorVersion = 5 Then
        GetWindowsVersion = WINDOWS_2000
      End If
    Case Else
      GetWindowsVersion = UNKNOWN_OS
    End Select
  End With

End Function

Private Function MessageBoxCritical(ByVal strMessage As String) As Long

' Displays a message box with a "critical" icon

  Dim mbpParams As MSGBOXPARAMS

  mbpParams.cbSize = Len(mbpParams)
  mbpParams.dwStyle = MB_OK Or MB_ICONSTOP
  mbpParams.lpszCaption = "HTML Help Class"
  mbpParams.lpszText = strMessage
    
  MessageBoxCritical = MessageBoxIndirect(mbpParams)
  
End Function

Private Function MessageBoxExclamation(ByVal strMessage As String) As Long

' Displays a message box with an "exclamation" icon

  Dim mbpParams As MSGBOXPARAMS

  mbpParams.cbSize = Len(mbpParams)
  mbpParams.dwStyle = MB_OK Or MB_ICONSTOP
  mbpParams.lpszCaption = "HTML Help Class"
  mbpParams.lpszText = strMessage
    
  MessageBoxExclamation = MessageBoxIndirect(mbpParams)
  
End Function

Private Function MessageBoxInformation(ByVal strMessage As String) As Long

' Displays a message box with an "information" icon

  Dim mbpParams As MSGBOXPARAMS

  mbpParams.cbSize = Len(mbpParams)
  mbpParams.dwStyle = MB_OK Or MB_ICONINFORMATION
  mbpParams.lpszCaption = "HTML Help Class"
  mbpParams.lpszText = strMessage
    
  MessageBoxInformation = MessageBoxIndirect(mbpParams)
  
End Function

Public Property Let HHWindow(ByVal vData As String)

' Specifies the HTML Help window for use
' with various methods.

  On Error Resume Next
    
  mvarHHWindow = vData
    
End Property

Public Property Get HHWindow() As String

  On Error Resume Next
    
  HHWindow = mvarHHWindow
    
End Property

Public Property Let HHTopicURL(ByVal vData As String)

' Specifies the topic path and file name for the
' HHDisplayTopicURL method

  mvarHHTopicURL = vData
    
End Property

Public Property Get HHTopicURL() As String
 
  On Error Resume Next
   
  HHTopicURL = mvarHHTopicURL
    
End Property

Public Property Let HHTopicID(ByVal vData As Long)

' Specifies the context integer for the
' HHDisplayTopicID method

  On Error Resume Next
    
  mvarHHTopicID = vData
    
End Property

Public Property Get HHTopicID() As Long

  On Error Resume Next
        
  HHTopicID = mvarHHTopicID
    
End Property

Public Property Let HHMsgTitle(ByVal vData As String)

' Specifies the title to display on a message box if
' a keyword cannot be found

  On Error Resume Next
    
  mvarHHMsgTitle = vData
    
End Property

Public Property Get HHMsgTitle() As String

  On Error Resume Next
    
  HHMsgTitle = mvarHHMsgTitle
    
End Property

Public Property Let HHMsgText(ByVal vData As String)

' Specifies the text to display in a message box if
' a keyword cannot be found

  On Error Resume Next
    
  mvarHHMsgText = vData
    
End Property

Public Property Get HHMsgText() As String

  On Error Resume Next
    
  HHMsgText = mvarHHMsgText
    
End Property

Public Property Let HHKeyword(ByVal vData As String)

' Specifies a keyword to search for using HHDisplayKeyword

  On Error Resume Next
    
  mvarHHKeyword = vData
    
End Property

Public Property Get HHKeyword() As String

  On Error Resume Next
    
  HHKeyword = mvarHHKeyword
    
End Property

Public Property Let HHDefaultURL(ByVal vData As String)

' Specifies the URL to use if a keyword cannot be found

  On Error Resume Next
    
  mvarHHDefaultURL = vData
    
End Property

Public Property Get HHDefaultURL() As String

  On Error Resume Next
    
  HHDefaultURL = mvarHHDefaultURL
    
End Property

Public Property Let HHALink(ByVal vData As String)

' Specifies an ALink keyword to search for using
' HHDisplayALink

  On Error Resume Next
    
  mvarHHALink = vData
    
End Property

Public Property Get HHALink() As String

  On Error Resume Next
    
  HHALink = mvarHHALink
    
End Property

Public Property Let CHMFile(ByVal vData As String)

' Path and file name of the HTML Help file to display

  On Error Resume Next
    
  mvarCHMFile = vData
    
End Property

Public Property Get CHMFile() As String

  On Error Resume Next
    
  CHMFile = mvarCHMFile
    
End Property

Public Property Let HHShowOnTop(ByVal vData As Boolean)

' If set to True, the HTML Help window will be set as
' a sibling of the calling window

  On Error Resume Next
    
  mvarHHShowOnTop = vData
    
End Property

Public Property Get HHShowOnTop() As Boolean

  On Error Resume Next
    
  HHShowOnTop = mvarHHShowOnTop
    
End Property

Public Property Get HHRegFileName() As String

' Used in conjunction with HHCheckRegistry,
' returns the confirmed CHM file name

  On Error Resume Next
    
  HHRegFileName = mvarHHRegFileName
    
End Property

Public Property Get HHRegFilePath() As String

' Used in conjunction with HHCheckRegistry,
' returns the path of the confirmed CHM file name

  On Error Resume Next
    
  HHRegFilePath = mvarHHRegFilePath
    
End Property

Public Property Get HHRegFileExists() As Boolean

' Used in conjunction with HHCheckRegistry, returns
' True if the file exists at the registered path,
' False if it doesn't

  On Error Resume Next
    
  HHRegFileExists = mvarHHRegFileExists
    
End Property

Public Property Let HHPopupFile(ByVal vData As String)

' Specifies the text file containing the popup text as
' listed in the [TEXT POPUPS] section of a CHM

  On Error Resume Next
    
  mvarHHPopupFile = vData
    
End Property

Public Property Let HHPopupID(ByVal vData As Long)

' Specifies the resource containing the to retrieve
' popup text from as listed in a Visual Basic project
' or the context integer of a popup topic as specified
' in a valid popup file in a CHM

  On Error Resume Next
    
  mvarHHPopupID = vData
    
End Property

Public Property Let HHPopupText(ByVal vData As String)

' Specifies the text to display in a simple text popup

  On Error Resume Next
    
  mvarHHPopupText = vData
    
End Property

Public Property Let HHPopupType(ByVal vData As PopupType)

' The type of popup to be generated by HHDisplayPopup

  On Error Resume Next
    
  mvarHHPopupType = vData
    
End Property

Public Property Let HHPopupTextColor(ByVal vData As ColorConstants)

' Sets the text color for a popup generated with HHDisplayPopup using
' the standard VB color constants

  On Error Resume Next
    
  mvarHHPopupTextColor = vData
    
End Property

Public Property Let HHPopupBackColor(ByVal vData As ColorConstants)

' Sets the back color of a popup generated with HHDisplayPopup using
' the standard VB color constants

  On Error Resume Next
    
  mvarHHPopupBackColor = vData
    
End Property

Public Property Let HHPopupCustomTextColor(ByVal vData As Long)

' Sets the text color for a popup generated with
' HHDisplayPopup using custom colors in the form &HBBGGRR.

  On Error Resume Next
    
  mvarHHPopupCustomTextColor = vData
    
End Property

Public Property Let HHPopupCustomBackColor(ByVal vData As Long)

' Sets the back color of a popup generated with
' HHDisplayPopup using custom colors in the form &HBBGGRR

  On Error Resume Next
    
  mvarHHPopupCustomBackColor = vData
    
End Property

Public Property Let HHPopupCustomColors(ByVal vData As Boolean)

' True of the HHPopupCustomTextColor and HHPopupCustomBackColor
' are going to be used, False if not

  On Error Resume Next
    
  mvarHHPopupCustomColors = vData
    
End Property

Public Property Get HHPopupCustomColors() As Boolean

  On Error Resume Next
        
  HHPopupCustomColors = mvarHHPopupCustomColors
    
End Property

Public Property Let HHPopupTextFont(ByVal vData As String)

' Name of the font to be used in a text popup

  On Error Resume Next
    
  mvarHHPopupTextFont = vData
    
End Property

Public Property Let HHPopupTextSize(ByVal vData As String)

' Point size of the font used in a text popup

  On Error Resume Next
    
  mvarHHPopupTextSize = vData
    
End Property

Public Property Let HHPopupTextBold(ByVal vData As Boolean)

' Set to True to make the popup text bold, False otherwise

  On Error Resume Next
    
  mvarHHPopupTextBold = vData
    
End Property

Public Property Let HHPopupTextItalic(ByVal vData As Boolean)

' Set to True to make the popup text italicized, False otherwise

  On Error Resume Next
    
  mvarHHPopupTextItalic = vData
    
End Property

Public Property Let HHPopupTextUnderline(ByVal vData As Boolean)

' Set to True to make the popup text underlined, False otherwise

  On Error Resume Next
    
  mvarHHPopupTextUnderline = vData
    
End Property

Public Property Get HHInstalled() As Boolean
   
' Verifies whether or not HTML Help is installed on
' the system.  This is done by checking the existence of
' HKEY_LOCAL_MACHINE\Software\CLASSES\TypeLib\{ADB880A2-D8FF-11CF-9377-00AA003B7A11}
   
  mvarHHInstalled = GetKeyInfo("SOFTWARE\Classes\TypeLib\{ADB880A2-D8FF-11CF-9377-00AA003B7A11}", 0)
   
  HHInstalled = mvarHHInstalled
    
End Property

Public Property Get HHVersion() As String

' Returns the current version of HTML Help on the
' system as a String expression

  On Error Resume Next
    
  GetHHVersion
        
  HHVersion = mvarHHVersion
    
End Property

Public Property Get IEVersion() As String

' Returns the current version of Internet Explorer on
' the system as a String expression

  On Error Resume Next
    
  GetIEVersion
        
  IEVersion = mvarIEVersion
    
End Property

Public Property Get HHFriendlyName() As String

' Returns the friendly name of the HTML Help version
' as a String expression (i.e, "1.21a")

  On Error Resume Next
    
  GetHHFriendlyName
        
  HHFriendlyName = mvarHHFriendlyName
    
End Property

Public Property Get IEFriendlyName() As String

' Returns the friendly name of the Internet Explorer
' version as a String expression (i.e, "5.0")

  On Error Resume Next
    
  GetIEFriendlyName
        
  IEFriendlyName = mvarIEFriendlyName
    
End Property
