Attribute VB_Name = "WinProcs"
' Modulo contenente le API di Windows (a 32 bit) per
' l'emissione dei triangoli e lo shade colorato.
' Viene gestita una palette personalizzata.


Type CornerRec
  x As Long
  Y As Long
End Type


Type PALETTEENTRY
  peRed As Byte
  peGreen As Byte
  peBlue As Byte
  peFlags As Byte
End Type

Type LOGPALETTE
  palVersion As Integer
  palNumEntries As Integer
  palPalEntry(16) As PALETTEENTRY
End Type

Declare Function CreatePalette Lib "GDI32" (LogicalPalette As LOGPALETTE) As Long
Declare Function CreatePen Lib "GDI32" (ByVal PenStyle As Long, ByVal Width As Long, ByVal Color As Long) As Long
Declare Function CreatePolygonRgn Lib "GDI32" (lpPoints As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Declare Function CreateSolidBrush Lib "GDI32" (ByVal rgbColor As Long) As Long
Declare Function DeleteObject Lib "GDI32" (ByVal hndobj As Long) As Long
Declare Function FillRgn Lib "GDI32" (ByVal hDC As Long, ByVal hRegion As Long, ByVal hBrush As Long) As Long
Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As Long, ByVal Index As Long) As Long
Declare Function GetSystemDirectory Lib "KERNEL32" Alias "GetSystemDirectoryA" (ByVal strBuffer As String, ByVal nBufLen As Long) As Long
Declare Function GetWindowsDirectory Lib "KERNEL32" Alias "GetWindowsDirectoryA" (ByVal strBuffer As String, ByVal nBufLen As Long) As Long
Declare Function LineTo Lib "GDI32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long) As Long
Declare Function MoveToEx Lib "GDI32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal NullPtr As Long) As Long
Declare Function PlaySound Lib "WINMM" (ByVal strName As String, ByVal hMod As Long, ByVal lFlags As Long) As Long
Declare Function RealizePalette Lib "GDI32" (ByVal hDC As Long) As Long
Declare Function SelectPalette Lib "GDI32" (ByVal hDC As Long, ByVal PaletteHandle As Long, ByVal Background As Long) As Long
Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal ObjectHandle As Long) As Long
Declare Function waveOutGetNumDevs Lib "WINMM" () As Long

Public Const PLANES = 14
Public Const BITSPIXEL = 12
Public Const PC_NOCOLLAPSE = 4
Public Const COLORS = 24
Public Const PS_SOLID = 0


Public Const NumColors = 16
Public Const TopColor = 12: ' all but last 3 colors are gray

Public OldPaletteHandle As Long
Public RedGreenBlue(16) As Long
Public Tilt As Double
Public UsePalette As Boolean
Public LogicalPalette As LOGPALETTE
Public NumRealized As Long
Sub SetColors(x As PictureBox)

  Dim ColorNum As Integer
  Dim NumBits As Long
  Dim NumColorsFree As Long
  Dim Tint As Integer
  
  
  OldPaletteHandle = 0
  
  NumColorsFree = 1
  NumBits = GetDeviceCaps(x.hDC, PLANES) * GetDeviceCaps(x.hDC, BITSPIXEL)
  If NumBits >= 31 Then
    UsePalette = False
  Else
    Do While (NumBits > 0)
      NumColorsFree = 2 * NumColorsFree
      NumBits = NumBits - 1
    Loop
    NumColorsFree = NumColorsFree - GetDeviceCaps(x.hDC, COLORS)
    If NumColorsFree < 16 Then
      UsePalette = False
    Else
      UsePalette = True
    End If
  End If
  LogicalPalette.palVersion = 3 * 256
  LogicalPalette.palNumEntries = 16
  For ColorNum = 0 To NumColors - 4
    ' Ciclo per definire l'ombra del colore
    Tint = (256 * ColorNum) \ (NumColors - 3)
    LogicalPalette.palPalEntry(ColorNum).peRed = Tint
    LogicalPalette.palPalEntry(ColorNum).peGreen = Tint
    LogicalPalette.palPalEntry(ColorNum).peBlue = Tint
    LogicalPalette.palPalEntry(ColorNum).peFlags = PC_NOCOLLAPSE
    RedGreenBlue(ColorNum) = RGB(Tint, Tint, 0)
  Next ColorNum
  
  If UsePalette Then
    PaletteHandle = CreatePalette(LogicalPalette)
    OldPaletteHandle = SelectPalette(x.hDC, PaletteHandle, 0)
    NumRealized = RealizePalette(x.hDC)
  End If

End Sub

Public Sub DrawTriangle(Pic As PictureBox, Box() As CornerRec, ColorNum As Integer)
  
  Dim Brush As Long
  Dim rc As Long
  Dim Region As Long
  Dim BaseCol As Long
  
 UsePalette = False
   
  BaseCol = 16777216
  
  If UsePalette Then
    Brush = CreateSolidBrush(BaseCol + ColorNum)
    If Brush Then
      Region = CreatePolygonRgn(Box(0), 3, 1)
      If Region Then
        rc = FillRgn(Pic.hDC, Region, Brush)
        rc = DeleteObject(Region)
      End If
      rc = DeleteObject(Brush)
    End If
  Else
    Brush = CreateSolidBrush(RedGreenBlue(ColorNum))
    If Brush Then
      Region = CreatePolygonRgn(Box(0), 3, 1)
      If Region Then
        rc = FillRgn(Pic.hDC, Region, Brush)
        rc = DeleteObject(Region)
      End If
      rc = DeleteObject(Brush)
    End If
  End If
 ' MsgBox "Ha disegnato un triangolo?"
End Sub


