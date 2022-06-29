Attribute VB_Name = "modMenu"
Option Explicit

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetObjectAPIBynum Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByVal lpObject As Long) As Long
Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function InsertMenuByNum Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Long) As Long
Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Public Const COLOR_MENU = 4

Public Const DT_LEFT = &H0
Public Const DT_SINGLELINE = &H20
Public Const DT_TOP = &H0

Public Const FF_DECORATIVE = 80
Public Const FF_DONTCARE = 0
Public Const FF_MODERN = 48
Public Const FF_ROMAN = 16
Public Const FF_SCRIPT = 64
Public Const FF_SWISS = 32

Public Const GWL_WNDPROC = -4

Public Const LF_FACESIZE = 32

Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&
Public Const MF_OWNERDRAW = &H100&
Public Const MF_POPUP = &H10&

Public Const ODS_SELECTED = &H1

Public Const ODT_MENU = 1

Public Const SYSTEM_FONT = 13

Public Const TRANSPARENT = 1

Public Const WM_COMMAND = &H111
Public Const WM_DRAWITEM = &H2B
Public Const WM_GETFONT = &H31
Public Const WM_MEASUREITEM = &H2C
Public Const WM_MENUSELECT = &H11F

Public Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type DRAWITEMSTRUCT
        CtlType As Long
        CtlID As Long
        itemID As Long
        itemAction As Long
        itemState As Long
        hwndItem As Long
        hdc As Long
        rcItem As RECT
        itemData As Long
End Type

Public Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(1 To LF_FACESIZE) As Byte
End Type

Public Type MEASUREITEMSTRUCT
        CtlType As Long
        CtlID As Long
        itemID As Long
        itemWidth As Long
        itemHeight As Long
        itemData As Long
End Type

Global gHW As Long
Global lpPrevWndProc As Long

'global variables to hold our menu handles
Global lngFile As Long
Global lngNew As Long
Global lngOpen As Long
Global lngSave As Long
Global lngSaveAs As Long
Global lngExit As Long
Global lngEdit As Long
Global lngUndo As Long
Global lngCut As Long
Global lngCopy As Long
Global lngPaste As Long
Global lngDelete As Long
Global lngSearch As Long
Global lngFind As Long
Global lngFindNext As Long
Global lngColors As Long
Global lngBlack As Long
Global lngGreen As Long
Global lngPurple As Long
Global lngRed As Long
Global lngYellow As Long
Global lngBlue As Long
Global lngWhite As Long

Public Sub Hook()
    'require all messages sent to our hwnd(gHW) to pass
    'through our WindowProc function before being processed
    lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub Unhook()
    'unhook our hwnd. this is vital.
    Call SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
End Sub

Private Function WindowProc(ByVal hW As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, retval As Long, nodef As Boolean) As Long
    Dim di As DRAWITEMSTRUCT
    Dim ms As MEASUREITEMSTRUCT
    Dim rc As RECT
    Dim strMenuItem As String * 128
    Dim intMenuHandle As Integer
    Select Case uMsg&
        Case WM_DRAWITEM
        'with our menus now being owner drawn, we can now
        'process the wm_drawitem message. wm_drawitem is
        'sent to our hwnd when the menu is being drawn.
        'by intercepting this message, we can change
        'various attributes of the menu including; the
        'text color and selection colors.
            Call CopyMemory(di, ByVal lParam, Len(di))
            Call DoMenuStuff(di)
        Case WM_MEASUREITEM
        'with our menus now being owner drawn, we can now
        'process the wm_measureitem message. this message
        'is also sent to our hwnd when the menu is being
        'drawn. we must intercept this message with owner
        'drawn menus in order to specify their size.
            Call CopyMemory(ms, ByVal lParam, Len(ms))
            'make sure our structure is owner drawn.
            If ms.CtlType = ODT_MENU Then
                'set initial values for the structure.
                ms.itemHeight = 20
                ms.itemWidth = 50
                'set size for the file menu's popup
                If ms.itemID > 70& And ms.itemID < 80& Then
                    ms.itemWidth = 65
                End If
                'set size for the edit menu's popup
                If ms.itemID > 80& And ms.itemID < 90& Then
                    ms.itemWidth = 60
                End If
                'set size for the search menu's popup
                If ms.itemID > 90& And ms.itemID < 100& Then
                    ms.itemWidth = 75
                End If
                'set size for the colors menu's popup
                If ms.itemID > 100& And ms.itemID < 110& Then
                    ms.itemWidth = 65
                End If
                'set size for the file and edit menus
                If ms.itemID = lngFile& Or ms.itemID = lngEdit& Then
                    ms.itemWidth = 20
                End If
                'set size for the search and colors menus
                If ms.itemID = lngSearch& Or ms.itemID = lngColors& Then
                    ms.itemWidth = 30
                End If
            End If
            Call CopyMemory(ByVal lParam, ms, Len(ms))
        Case WM_COMMAND
            'this is where we will handle the wm_command
            'message that is sent when a menu item is clicked.
            Select Case wParam
                Case lngNew&
                    MsgBox "selected menu - new"
                Case lngOpen&
                    MsgBox "selected menu - open"
                Case lngSave&
                    MsgBox "selected menu - save"
                Case lngSaveAs&
                    MsgBox "selected menu - save as"
                Case lngExit&
                    MsgBox "selected menu - exit"
                Case lngUndo&
                    MsgBox "selected menu - undo"
                Case lngCut&
                    MsgBox "selected menu - cut"
                Case lngCopy&
                    MsgBox "selected menu - copy"
                Case lngPaste&
                    MsgBox "selected menu - paste"
                Case lngDelete&
                    MsgBox "selected menu - delete"
                Case lngFind&
                    MsgBox "selected menu - find"
                Case lngFindNext&
                    MsgBox "selected menu - find next"
                Case lngBlack&
                    MsgBox "selected menu - black"
                Case lngGreen&
                    MsgBox "selected menu - green"
                Case lngPurple&
                    MsgBox "selected menu - purple"
                Case lngRed&
                    MsgBox "selected menu - red"
                Case lngYellow&
                    MsgBox "selected menu - yellow"
                Case lngBlue&
                    MsgBox "selected menu - blue"
                Case lngWhite&
                    MsgBox "selected menu - white"
            End Select
    End Select
    WindowProc = CallWindowProc(lpPrevWndProc, hW, uMsg, wParam, lParam)
End Function

Private Sub DoMenuStuff(ds As DRAWITEMSTRUCT)
    Dim UseBrush As Long, OldBrush As Long
    Dim UsePen As Long, OldPen As Long
    Dim CurFnt As Long, NewFnt As Long
    Dim Lf As LOGFONT, UseRect As RECT
    Dim TopRect As RECT
    'check to see if our menu item is selected
    If ds.itemState And ODS_SELECTED Then
        'set our menu selection highlight color
        Select Case ds.itemID
            Case lngBlack&
                UseBrush = CreateSolidBrush(&H0&)
            Case lngGreen&
                UseBrush = CreateSolidBrush(&HC000&)
            Case lngPurple&
                UseBrush = CreateSolidBrush(&HC000C0)
            Case lngYellow&
                UseBrush = CreateSolidBrush(&HFFFF&)
            Case lngBlue&
                UseBrush = CreateSolidBrush(&HFF0000)
            Case lngWhite&
                UseBrush = CreateSolidBrush(&HFFFFFF)
            Case Else
                UseBrush = CreateSolidBrush(&HFF)
        End Select
    Else
        'set our menu unselected highlight color
        UseBrush = CreateSolidBrush(GetSysColor(COLOR_MENU))
    End If
    'fill our selected area with color
    Call FillRect(ds.hdc, ds.rcItem, UseBrush)
    'then delete the brush we created
    If UseBrush Then
        Call DeleteObject(UseBrush)
    End If
    'again check to see if the menu item is selected
    If ds.itemState And ODS_SELECTED Then
        'set our menu selected text color
        Select Case ds.itemID
            Case lngYellow
                Call SetTextColor(ds.hdc, &H0&)
            Case lngWhite
                Call SetTextColor(ds.hdc, &H0&)
            Case Else
                Call SetTextColor(ds.hdc, &HFFFFFF)
        End Select
    Else
        'set our text unselected text color
        Call SetTextColor(ds.hdc, &H0)
    End If
    Call SetBkMode(ds.hdc, TRANSPARENT)
    'create two regions. userect will be for drop down
    'menu items and toprect will be used for top menus.
    LSet UseRect = ds.rcItem
    LSet TopRect = ds.rcItem
    UseRect.Left = UseRect.Left + 16
    'retrieve the system font
    CurFnt = SelectObject(ds.hdc, GetStockObject(SYSTEM_FONT))
    Call GetObjectAPIBynum(CurFnt, Len(Lf), VarPtr(Lf))
    'i set the font to ff_script, but you can also use;
    'ff_dontcare, ff_modern, ff_roman, ff_swiss.
    Lf.lfPitchAndFamily = FF_SCRIPT
    Lf.lfFaceName(1) = 1
    'i set the font weight to 400 for normal text. it would
    'be bolded at 600.
    Lf.lfWeight = 400
    'create and select our new font.
    NewFnt = CreateFontIndirect(Lf)
    Call SelectObject(ds.hdc, NewFnt)
    'set our top and left corners for our regions.
    UseRect.Left = UseRect.Left + 5
    UseRect.Top = UseRect.Top + 3
    TopRect.Left = TopRect.Left + 5
    TopRect.Top = TopRect.Top + 2
    'now we will draw our text according to the menu's id.
    'we must specify the text and the region we are drawing.
    Select Case ds.itemID
        Case lngFile&
            Call DrawText(ds.hdc, "file", 4, TopRect, DT_LEFT Or DT_TOP Or DT_SINGLELINE)
        Case lngNew&
            Call DrawText(ds.hdc, "new", 3, UseRect, DT_LEFT Or DT_TOP Or DT_SINGLELINE)
        Case lngOpen&
            Call DrawText(ds.hdc, "open", 4, UseRect, DT_LEFT Or DT_TOP Or DT_SINGLELINE)
        Case lngSave&
            Call DrawText(ds.hdc, "save", 4, UseRect, DT_LEFT Or DT_TOP Or DT_SINGLELINE)
        Case lngSaveAs&
            Call DrawText(ds.hdc, "save as", 7, UseRect, DT_LEFT Or DT_TOP Or DT_SINGLELINE)
        Case lngExit&
            Call DrawText(ds.hdc, "exit", 4, UseRect, DT_LEFT Or DT_TOP Or DT_SINGLELINE)
        Case lngEdit&
            Call DrawText(ds.hdc, "edit", 4, TopRect, DT_LEFT Or DT_TOP Or DT_SINGLELINE)
        Case lngUndo&
            Call DrawText(ds.hdc, "undo", 4, UseRect, DT_LEFT Or DT_TOP Or DT_SINGLELINE)
        Case lngCut&
            Call DrawText(ds.hdc, "cut", 3, UseRect, DT_LEFT Or DT_TOP Or DT_SINGLELINE)
        Case lngCopy&
            Call DrawText(ds.hdc, "copy", 4, UseRect, DT_LEFT Or DT_TOP Or DT_SINGLELINE)
        Case lngPaste&
            Call DrawText(ds.hdc, "paste", 5, UseRect, DT_LEFT Or DT_TOP Or DT_SINGLELINE)
        Case lngDelete&
            Call DrawText(ds.hdc, "delete", 6, UseRect, DT_LEFT Or DT_TOP Or DT_SINGLELINE)
        Case lngSearch&
            Call DrawText(ds.hdc, "search", 6, TopRect, DT_LEFT Or DT_TOP Or DT_SINGLELINE)
        Case lngFind&
            Call DrawText(ds.hdc, "find", 4, UseRect, DT_LEFT Or DT_TOP Or DT_SINGLELINE)
        Case lngFindNext&
            Call DrawText(ds.hdc, "find next", 9, UseRect, DT_LEFT Or DT_TOP Or DT_SINGLELINE)
        Case lngColors&
            Call DrawText(ds.hdc, "colors", 6, TopRect, DT_LEFT Or DT_TOP Or DT_SINGLELINE)
        Case lngBlack&
            Call DrawText(ds.hdc, "black", 5, UseRect, DT_LEFT Or DT_TOP Or DT_SINGLELINE)
        Case lngGreen&
            Call DrawText(ds.hdc, "green", 5, UseRect, DT_LEFT Or DT_TOP Or DT_SINGLELINE)
        Case lngPurple&
            Call DrawText(ds.hdc, "purple", 6, UseRect, DT_LEFT Or DT_TOP Or DT_SINGLELINE)
        Case lngRed&
            Call DrawText(ds.hdc, "red", 3, UseRect, DT_LEFT Or DT_TOP Or DT_SINGLELINE)
        Case lngYellow&
            Call DrawText(ds.hdc, "yellow", 6, UseRect, DT_LEFT Or DT_TOP Or DT_SINGLELINE)
        Case lngBlue&
            Call DrawText(ds.hdc, "blue", 4, UseRect, DT_LEFT Or DT_TOP Or DT_SINGLELINE)
        Case lngWhite&
            Call DrawText(ds.hdc, "white", 5, UseRect, DT_LEFT Or DT_TOP Or DT_SINGLELINE)
    End Select
    'select our old font and delete the one we created.
    Call SelectObject(ds.hdc, CurFnt)
    If NewFnt Then
        Call DeleteObject(NewFnt)
    End If
End Sub
