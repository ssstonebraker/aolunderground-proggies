Attribute VB_Name = "modGame"
'The BitBlt function allows for fast and smooth drawing to the form
'and to picture boxes, but isn't as fast as it should be for making games.
'It stands for bit-block transfer.
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal animX As Long, ByVal animY As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'same as bitblt, but allows stretching.
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

'allows the playing of wav files
Public Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'for the bitblt function
Public Const SRCCOPY = &HCC0020   'Copies the source over the destination
Public Const SRCINVERT = &H660046 'Copies and inverts the source over the destination
Public Const SRCAND = &H8800C6    'Adds the source to the destination

'holds the tile type for each tile (0 = walkable, 1 = non walkable, 2 = door)
Public Walkable(0 To 899) As Integer
'holds the texture number for each tile
Public Texture(0 To 899) As Integer
'holds the location of the top left corner of each tile
Public tileLeft(0 To 899) As Integer
Public tileTop(0 To 899) As Integer
'these hold the location of the map a tile with a walkable value of 2 leads to
Public mapXStored(0 To 899) As Integer
Public mapYStored(0 To 899) As Integer
Public mapAreaStored(0 To 899) As String

'shows whether the game has started or not
Public GameInProgress As Boolean

Public Character As New clsHero

'these are for changing the game speed
Public Speed As Integer    'holds the current speed
Public wait As Double      'holds the delay value

Public sndStep As String
Public sndButton As String

'symbolic constants - makes code easier to read
Public Const MAGE = 1
Public Const WARRIOR = 2
Public Const BARBARIAN = 3
Public Const CLARIC = 4

Public Const STR = 0
Public Const STA = 1
Public Const MAG = 2
Public Const INTEL = 3
Public Const VITAL = 4

'for buttons
Public Const CANCEL = 0
Public Const OK = 1

Public Function msg(ByVal message As String, ByRef callingFrm As Form)
    Call frmMessage.showMessage(message, callingFrm)
End Function
