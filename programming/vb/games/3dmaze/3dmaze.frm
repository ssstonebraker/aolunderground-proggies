VERSION 5.00
Begin VB.Form frm3DMaze 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Maze 'O' MaNiA!"
   ClientHeight    =   4470
   ClientLeft      =   1560
   ClientTop       =   1980
   ClientWidth     =   6000
   Icon            =   "3dmaze.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4470
   ScaleWidth      =   6000
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   288
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   4200
      Width           =   6012
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   4212
      LargeChange     =   5
      Left            =   5760
      Max             =   60
      Min             =   30
      TabIndex        =   0
      Top             =   0
      Value           =   45
      Width           =   252
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   55
      Left            =   0
      Top             =   0
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Begin VB.Menu mnuActionItem 
         Caption         =   "&New"
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuActionItem 
         Caption         =   "&Solve"
         Enabled         =   0   'False
         Index           =   1
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuActionItem 
         Caption         =   "&Clear"
         Enabled         =   0   'False
         Index           =   2
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuActionItem 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuActionItem 
         Caption         =   "E&xit"
         Index           =   4
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuStyle 
      Caption         =   "&Style"
      Begin VB.Menu mnuStyleItem 
         Caption         =   "&Hexagonal rooms"
         Index           =   0
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuStyleItem 
         Caption         =   "&Square rooms"
         Index           =   1
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&About..."
         Index           =   0
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frm3DMaze"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Type CornerRec
  X As Long
  Y As Long
End Type

Private Type VertexRec
  X As Double
  Y As Double
End Type

Private Type StackRec
  Index1 As Byte
  Index2 As Integer
End Type

Private Type PALETTEENTRY
  peRed As Byte
  peGreen As Byte
  peBlue As Byte
  peFlags As Byte
End Type

Private Type LOGPALETTE
  palVersion As Integer
  palNumEntries As Integer
  palPalEntry(16) As PALETTEENTRY
End Type

Private Declare Function CreatePalette Lib "GDI32" (LogicalPalette As LOGPALETTE) As Long
Private Declare Function CreatePen Lib "GDI32" (ByVal PenStyle As Long, ByVal Width As Long, ByVal Color As Long) As Long
Private Declare Function CreatePolygonRgn Lib "GDI32" (lpPoints As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateSolidBrush Lib "GDI32" (ByVal rgbColor As Long) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hndobj As Long) As Long
Private Declare Function FillRgn Lib "GDI32" (ByVal hDC As Long, ByVal hRegion As Long, ByVal hBrush As Long) As Long
Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As Long, ByVal Index As Long) As Long
Private Declare Function LineTo Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal NullPtr As Long) As Long
Private Declare Function RealizePalette Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function SelectPalette Lib "GDI32" (ByVal hDC As Long, ByVal PaletteHandle As Long, ByVal Background As Long) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal ObjectHandle As Long) As Long

Const PLANES = 14
Const BITSPIXEL = 12
Const PC_NOCOLLAPSE = 4
Const COLORS = 24
Const PS_SOLID = 0

Const NumColors = 16
Const TopColor = 12: ' all but last 3 colors are gray
Const RectangleSENWColor = 10
Const TriangleSSENNWColor = 9
Const TriangleSENWColor = 8
Const RectangleWEColor = 7
Const FloorColor = 6
Const TriangleSWNEColor = 5
Const RectangleSWNEColor = 4
Const TriangleSSWNNEColor = 3
Const BackoutColor = 13
Const AdvanceColor = 14
Const SolutionColor = 15

Const RelativeWidthOfWall = 0.25: ' relative to side of hexagon or square
Const RelativeHeightOfWall = 2#: ' relative to side of hexagon or square
Const MinWallLengthInInches = 0.25

Const SecondsForMazeSelection = 0.25

Dim AlreadyPainting As Boolean
Dim BaseRectangle(5, 3) As VertexRec
Dim BaseTriangle(3, 2) As VertexRec
Dim ComputerPage() As Byte
Dim CosTilt As Double
Dim CurrentColor As Integer
Dim HexDeltaX(5, 719) As Integer
Dim HexDeltaY(5, 719) As Integer
Dim MaxX As Integer
Dim MaxY As Integer
Dim Minimized As Boolean
Dim NumColumns As Integer
Dim NumRealized As Long
Dim NumRoomsInMaze As Integer
Dim NumRows As Integer
Dim OldPaletteHandle As Long
Dim Paint As Boolean
Dim PaletteHandle As Long
Dim PixelsPerX As Double
Dim PixelsPerZ As Double
Dim Rectangle(5, 3) As VertexRec
Dim RedGreenBlue(16) As Long
Dim RelDistOfUserFromScreen As Double
Dim Resize As Boolean
Dim Seed As String
Dim SinTilt As Double
Dim SolutionDisplayed As Boolean
Dim SqrDeltaX(3, 23) As Integer
Dim SqrDeltaY(3, 23) As Integer
Dim Sqrt3 As Double
Dim Stack() As StackRec
Dim State As Byte
Dim SubstitutionHigh(99) As Byte
Dim SubstitutionLow(99) As Byte
Dim Tilt As Double
Dim UsePalette As Boolean
Dim UserHasSolved As Boolean
Dim UserPage() As Byte
Dim UserX As Integer
Dim UserXRelative As Double
Dim UserY As Integer
Dim UserYRelative As Double
Dim X As Integer
Dim XMax As Double
Dim XOffset As Double
Dim Y As Integer
Dim YMax As Double
Dim YMod4 As Byte
Dim YOffset As Double

Private Sub DrawQuadrilateral(Box() As CornerRec, ColorNum As Integer)
  Dim Brush As Long
  Dim rc As Long
  Dim Region As Long
  If UsePalette Then
    Brush = CreateSolidBrush(16777216 + ColorNum)
    If Brush Then
      Region = CreatePolygonRgn(Box(0), 4, 1)
      If Region Then
        rc = FillRgn(frm3DMaze.hDC, Region, Brush)
        rc = DeleteObject(Region)
      End If
      rc = DeleteObject(Brush)
    End If
  Else
    Brush = CreateSolidBrush(RedGreenBlue(ColorNum))
    If Brush Then
      Region = CreatePolygonRgn(Box(0), 4, 1)
      If Region Then
        rc = FillRgn(frm3DMaze.hDC, Region, Brush)
        rc = DeleteObject(Region)
      End If
      rc = DeleteObject(Brush)
    End If
  End If
End Sub

Private Sub GetCorner(X#, Y#, Z#, PixelsPerX#, PixelsPerZ#, CosTilt#, SinTilt#, RelDistOfUserFromScreen#, XMax#, XOffset#, YMax#, Corner As CornerRec)
  Dim XAdjusted As Double
  Dim YPrime As Double
  Dim ZAdjusted As Double
  Dim ZPrime As Double

  YPrime = (YMax# - Y#) * CosTilt# - Z# * SinTilt#
  ZPrime = (YMax# - Y#) * SinTilt# + Z# * CosTilt#
  ZAdjusted = (YMax# / 2#) + RelDistOfUserFromScreen# * (ZPrime - (YMax# / 2#)) / (YPrime + RelDistOfUserFromScreen#)
  XAdjusted = (XMax# / 2#) + RelDistOfUserFromScreen# * (X# - (XMax# / 2#)) / (YPrime + RelDistOfUserFromScreen#)
  XAdjusted = XAdjusted + XOffset#
  Corner.X = Int(PixelsPerX# * XAdjusted)
  Corner.Y = (ScaleHeight - Text1.Height) - Int(PixelsPerZ# * ZAdjusted)
End Sub

Private Sub DisplayQuadrilateral(XMax#, XOffset#, YMax#, X0#, Y0#, Z0#, X1#, Y1#, Z1#, X2#, Y2#, Z2#, X3#, Y3#, Z3#, PixelsPerX#, PixelsPerZ#, CosTilt#, SinTilt#, RelDistOfUserFromScreen#, Shade%)
  Dim Quadrilateral(3) As CornerRec
  Dim TemQuad As CornerRec
  Call GetCorner(X0#, Y0#, Z0#, PixelsPerX#, PixelsPerZ#, CosTilt#, SinTilt#, RelDistOfUserFromScreen#, XMax#, XOffset#, YMax#, TemQuad)
  Quadrilateral(0).X = TemQuad.X
  Quadrilateral(0).Y = TemQuad.Y
  Call GetCorner(X1#, Y1#, Z1#, PixelsPerX#, PixelsPerZ#, CosTilt#, SinTilt#, RelDistOfUserFromScreen#, XMax#, XOffset#, YMax#, TemQuad)
  Quadrilateral(1).X = TemQuad.X
  Quadrilateral(1).Y = TemQuad.Y
  Call GetCorner(X2#, Y2#, Z2#, PixelsPerX#, PixelsPerZ#, CosTilt#, SinTilt#, RelDistOfUserFromScreen#, XMax#, XOffset#, YMax#, TemQuad)
  Quadrilateral(2).X = TemQuad.X
  Quadrilateral(2).Y = TemQuad.Y
  Call GetCorner(X3#, Y3#, Z3#, PixelsPerX#, PixelsPerZ#, CosTilt#, SinTilt#, RelDistOfUserFromScreen#, XMax#, XOffset#, YMax#, TemQuad)
  Quadrilateral(3).X = TemQuad.X
  Quadrilateral(3).Y = TemQuad.Y
  Call DrawQuadrilateral(Quadrilateral(), Shade%)
End Sub

Private Sub DrawTriangle(Box() As CornerRec, ColorNum As Integer)
  Dim Brush As Long
  Dim rc As Long
  Dim Region As Long
  If UsePalette Then
    Brush = CreateSolidBrush(16777216 + ColorNum)
    If Brush Then
      Region = CreatePolygonRgn(Box(0), 3, 1)
      If Region Then
        rc = FillRgn(frm3DMaze.hDC, Region, Brush)
        rc = DeleteObject(Region)
      End If
      rc = DeleteObject(Brush)
    End If
  Else
    Brush = CreateSolidBrush(RedGreenBlue(ColorNum))
    If Brush Then
      Region = CreatePolygonRgn(Box(0), 3, 1)
      If Region Then
        rc = FillRgn(frm3DMaze.hDC, Region, Brush)
        rc = DeleteObject(Region)
      End If
      rc = DeleteObject(Brush)
    End If
  End If
End Sub

Private Sub DisplayTriangle(XMax#, XOffset#, YMax#, X0#, Y0#, Z0#, X1#, Y1#, Z1#, X2#, Y2#, Z2#, PixelsPerX#, PixelsPerZ#, CosTilt#, SinTilt#, RelDistOfUserFromScreen#, Shade%)
  Dim Triangle(2) As CornerRec
  Dim TemTriangle As CornerRec
  Call GetCorner(X0#, Y0#, Z0#, PixelsPerX#, PixelsPerZ#, CosTilt#, SinTilt#, RelDistOfUserFromScreen#, XMax#, XOffset#, YMax#, TemTriangle)
  Triangle(0).X = TemTriangle.X
  Triangle(0).Y = TemTriangle.Y
  Call GetCorner(X1#, Y1#, Z1#, PixelsPerX#, PixelsPerZ#, CosTilt#, SinTilt#, RelDistOfUserFromScreen#, XMax#, XOffset#, YMax#, TemTriangle)
  Triangle(1).X = TemTriangle.X
  Triangle(1).Y = TemTriangle.Y
  Call GetCorner(X2#, Y2#, Z2#, PixelsPerX#, PixelsPerZ#, CosTilt#, SinTilt#, RelDistOfUserFromScreen#, XMax#, XOffset#, YMax#, TemTriangle)
  Triangle(2).X = TemTriangle.X
  Triangle(2).Y = TemTriangle.Y
  Call DrawTriangle(Triangle(), Shade%)
End Sub

Private Sub OutputTriangle(XMax#, XOffset#, YMax#, Triangle() As VertexRec, PixelsPerX#, PixelsPerZ#, CosTilt#, SinTilt#, RelDistOfUserFromScreen#, FirstPass%, FaceColor%)
  Dim X0 As Double
  Dim X1 As Double
  Dim X2 As Double
  Dim X3 As Double
  Dim Y0 As Double
  Dim Y1 As Double
  Dim Y2 As Double
  Dim Y3 As Double
  If FirstPass% Then
    If ((Triangle(1).X < XMax# / 2#) And (Triangle(1).X > Triangle(0).X)) Then
      X0 = Triangle(2).X
      Y0 = Triangle(2).Y
      X1 = Triangle(1).X
      Y1 = Triangle(1).Y
      X2 = Triangle(1).X
      Y2 = Triangle(1).Y
      X3 = Triangle(2).X
      Y3 = Triangle(2).Y
      Call DisplayQuadrilateral(XMax#, XOffset#, YMax#, X0, Y0, RelativeHeightOfWall, X1, Y1, RelativeHeightOfWall, X2, Y2, 0#, X3, Y3, 0#, PixelsPerX#, PixelsPerZ#, CosTilt#, SinTilt#, RelDistOfUserFromScreen#, TriangleSSWNNEColor)
    End If
    If ((Triangle(1).X > XMax# / 2#) And (Triangle(1).X < Triangle(2).X)) Then
      X0 = Triangle(1).X
      Y0 = Triangle(1).Y
      X1 = Triangle(0).X
      Y1 = Triangle(0).Y
      X2 = Triangle(0).X
      Y2 = Triangle(0).Y
      X3 = Triangle(1).X
      Y3 = Triangle(1).Y
      Call DisplayQuadrilateral(XMax#, XOffset#, YMax#, X0, Y0, RelativeHeightOfWall, X1, Y1, RelativeHeightOfWall, X2, Y2, 0#, X3, Y3, 0#, PixelsPerX#, PixelsPerZ#, CosTilt#, SinTilt#, RelDistOfUserFromScreen#, TriangleSSENNWColor)
    End If
  Else
    X0 = Triangle(0).X
    Y0 = Triangle(0).Y
    X1 = Triangle(2).X
    Y1 = Triangle(2).Y
    X2 = Triangle(2).X
    Y2 = Triangle(2).Y
    X3 = Triangle(0).X
    Y3 = Triangle(0).Y
    Call DisplayQuadrilateral(XMax#, XOffset#, YMax#, X0, Y0, RelativeHeightOfWall, X1, Y1, RelativeHeightOfWall, X2, Y2, 0#, X3, Y3, 0#, PixelsPerX#, PixelsPerZ#, CosTilt#, SinTilt#, RelDistOfUserFromScreen#, FaceColor%)
    X0 = Triangle(0).X
    Y0 = Triangle(0).Y
    X1 = Triangle(1).X
    Y1 = Triangle(1).Y
    X2 = Triangle(2).X
    Y2 = Triangle(2).Y
    Call DisplayTriangle(XMax#, XOffset#, YMax#, X0, Y0, RelativeHeightOfWall, X1, Y1, RelativeHeightOfWall, X2, Y2, RelativeHeightOfWall, PixelsPerX#, PixelsPerZ#, CosTilt#, SinTilt#, RelDistOfUserFromScreen#, TopColor)
  End If
End Sub

Private Sub OutputRectangle(XMax As Double, XOffset As Double, YMax As Double, Rectangle() As VertexRec, PixelsPerX As Double, PixelsPerZ As Double, CosTilt As Double, SinTilt As Double, RelDistOfUserFromScreen As Double, FaceColor As Integer)
  Dim X0 As Double
  Dim X1 As Double
  Dim X2 As Double
  Dim X3 As Double
  Dim Y0 As Double
  Dim Y1 As Double
  Dim Y2 As Double
  Dim Y3 As Double
  X0 = Rectangle(3).X
  Y0 = Rectangle(3).Y
  X1 = Rectangle(2).X
  Y1 = Rectangle(2).Y
  X2 = Rectangle(2).X
  Y2 = Rectangle(2).Y
  X3 = Rectangle(3).X
  Y3 = Rectangle(3).Y
  Call DisplayQuadrilateral(XMax, XOffset, YMax, X0, Y0, RelativeHeightOfWall, X1, Y1, RelativeHeightOfWall, X2, Y2, 0#, X3, Y3, 0#, PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen, FaceColor)
  X0 = Rectangle(0).X
  Y0 = Rectangle(0).Y
  X1 = Rectangle(1).X
  Y1 = Rectangle(1).Y
  X2 = Rectangle(2).X
  Y2 = Rectangle(2).Y
  X3 = Rectangle(3).X
  Y3 = Rectangle(3).Y
  Call DisplayQuadrilateral(XMax, XOffset, YMax, X0, Y0, RelativeHeightOfWall, X1, Y1, RelativeHeightOfWall, X2, Y2, RelativeHeightOfWall, X3, Y3, RelativeHeightOfWall, PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen, TopColor)
End Sub

Private Sub OutputLeftRight(XMax As Double, XOffset As Double, YMax As Double, Rectangle() As VertexRec, PixelsPerX As Double, PixelsPerZ As Double, CosTilt As Double, SinTilt As Double, RelDistOfUserFromScreen As Double)
  Dim X0 As Double
  Dim X1 As Double
  Dim X2 As Double
  Dim X3 As Double
  Dim Y0 As Double
  Dim Y1 As Double
  Dim Y2 As Double
  Dim Y3 As Double
  If 2# * Rectangle(0).X > XMax Then
    X0 = Rectangle(0).X
    Y0 = Rectangle(0).Y
    X1 = Rectangle(3).X
    Y1 = Rectangle(3).Y
    X2 = Rectangle(3).X
    Y2 = Rectangle(3).Y
    X3 = Rectangle(0).X
    Y3 = Rectangle(0).Y
    Call DisplayQuadrilateral(XMax, XOffset, YMax, X0, Y0, RelativeHeightOfWall, X1, Y1, RelativeHeightOfWall, X2, Y2, 0#, X3, Y3, 0#, PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen, RectangleSENWColor)
  End If
  If 2# * Rectangle(1).X < XMax Then
    X0 = Rectangle(2).X
    Y0 = Rectangle(2).Y
    X1 = Rectangle(1).X
    Y1 = Rectangle(1).Y
    X2 = Rectangle(1).X
    Y2 = Rectangle(1).Y
    X3 = Rectangle(2).X
    Y3 = Rectangle(2).Y
    Call DisplayQuadrilateral(XMax, XOffset, YMax, X0, Y0, RelativeHeightOfWall, X1, Y1, RelativeHeightOfWall, X2, Y2, 0#, X3, Y3, 0#, PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen, RectangleSWNEColor)
  End If
End Sub

Private Sub DrawLine(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, XMax As Double, XOffset As Double, YMax As Double, CosTilt As Double, SinTilt As Double, PixelsPerX As Double, PixelsPerZ As Double, RelDistOfUserFromScreen As Double)
  Dim LineX1 As Long
  Dim LineX2 As Long
  Dim LineY1 As Long
  Dim LineY2 As Long
  Dim Pen As Long
  Dim PreviousPen As Long
  Dim rc As Long
  Dim tem As CornerRec
  Call GetCorner(X1, Y1, RelativeHeightOfWall, PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen, XMax, XOffset, YMax, tem)
  LineX1 = tem.X
  LineY1 = tem.Y
  Call GetCorner(X2, Y2, RelativeHeightOfWall, PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen, XMax, XOffset, YMax, tem)
  LineX2 = tem.X
  LineY2 = tem.Y
  If UsePalette Then
    Pen = CreatePen(PS_SOLID, 2, 16777216 + CurrentColor)
    If Pen Then
      PreviousPen = SelectObject(frm3DMaze.hDC, Pen)
      rc = MoveToEx(frm3DMaze.hDC, LineX1, LineY1, 0)
      rc = LineTo(frm3DMaze.hDC, LineX2, LineY2)
      rc = SelectObject(frm3DMaze.hDC, PreviousPen)
      rc = DeleteObject(Pen)
    End If
  Else
    Pen = CreatePen(PS_SOLID, 2, RedGreenBlue(CurrentColor))
    If Pen Then
      PreviousPen = SelectObject(frm3DMaze.hDC, Pen)
      rc = MoveToEx(frm3DMaze.hDC, LineX1, LineY1, 0)
      rc = LineTo(frm3DMaze.hDC, LineX2, LineY2)
      rc = SelectObject(frm3DMaze.hDC, PreviousPen)
      rc = DeleteObject(Pen)
    End If
  End If
End Sub

Private Sub Hash(Counter0 As Byte, Counter1 As Byte, Counter2 As Byte, Counter3 As Byte, Counter4 As Byte, Counter5 As Byte, Counter6 As Byte, Counter7 As Byte)
  Dim Iteration As Byte
  Dim Seed0 As Byte
  Dim Seed1 As Byte
  Dim Seed2 As Byte
  Dim Seed3 As Byte
  Dim Seed4 As Byte
  Dim Seed5 As Byte
  Dim Seed6 As Byte
  Dim Seed7 As Byte
  Dim SubstitutionIndex As Byte
  Dim Tem0 As Byte
  Dim Tem1 As Byte
  Dim Tem2 As Byte
  Seed0 = Counter0
  Seed1 = Counter1
  Seed2 = Counter2
  Seed3 = Counter3
  Seed4 = Counter4
  Seed5 = Counter5
  Seed6 = Counter6
  Seed7 = Counter7
  For Iteration = 1 To 8
    SubstitutionIndex = 10 * Seed1 + Seed0
    Tem0 = SubstitutionLow(SubstitutionIndex)
    Tem1 = SubstitutionHigh(SubstitutionIndex)
    SubstitutionIndex = 10 * Seed3 + Seed2
    Seed0 = SubstitutionLow(SubstitutionIndex)
    Tem2 = SubstitutionHigh(SubstitutionIndex)
    SubstitutionIndex = 10 * Seed5 + Seed4
    Seed2 = SubstitutionLow(SubstitutionIndex)
    Seed1 = SubstitutionHigh(SubstitutionIndex)
    SubstitutionIndex = 10 * Seed7 + Seed6
    Seed5 = SubstitutionLow(SubstitutionIndex)
    Seed7 = SubstitutionHigh(SubstitutionIndex)
    Seed3 = Tem0
    Seed6 = Tem1
    Seed4 = Tem2
  Next Iteration
  Counter0 = Seed0
  Counter1 = Seed1
  Counter2 = Seed2
  Counter3 = Seed3
  Counter4 = Seed4
  Counter5 = Seed5
  Counter6 = Seed6
  Counter7 = Seed7
End Sub

Private Sub Increment(Counter0 As Byte, Counter1 As Byte, Counter2 As Byte, Counter3 As Byte, Counter4 As Byte, Counter5 As Byte, Counter6 As Byte, Counter7 As Byte)
  Dim tem As Byte
  tem = Counter0 + 1
  If tem <= 9 Then
    Counter0 = tem
  Else
    Counter0 = 0
    tem = Counter1 + 1
    If tem <= 9 Then
      Counter1 = tem
    Else
      Counter1 = 0
      tem = Counter2 + 1
      If tem <= 9 Then
        Counter2 = tem
      Else
        Counter2 = 0
        tem = Counter3 + 1
        If tem <= 9 Then
          Counter3 = tem
        Else
          Counter3 = 0
          tem = Counter4 + 1
          If tem <= 9 Then
            Counter4 = tem
          Else
            Counter4 = 0
            tem = Counter5 + 1
            If tem <= 9 Then
              Counter5 = tem
            Else
              Counter5 = 0
              tem = Counter6 + 1
              If tem <= 9 Then
                Counter6 = tem
              Else
                Counter6 = 0
                tem = Counter7 + 1
                If tem <= 9 Then
                  Counter7 = tem
                Else
                  Counter7 = 0
                End If
              End If
            End If
          End If
        End If
      End If
    End If
  End If
End Sub

Private Sub HexDisplaySolution(MaxY As Integer, Page() As Byte, XMax As Double, XOffset As Double, YMax As Double, CosTilt As Double, SinTilt As Double, PixelsPerX As Double, PixelsPerZ As Double, RelDistOfUserFromScreen As Double)
  Dim DeltaIndex As Byte
  Dim OldPaletteHandle As Long
  Dim PathFound As Integer
  Dim X As Integer
  Dim XNext As Integer
  Dim XPrevious As Integer
  Dim XRelative As Double
  Dim XRelativeNext As Double
  Dim Y As Integer
  Dim YNext As Integer
  Dim YPrevious As Integer
  Dim YRelative As Double
  Dim YRelativeNext As Double
  If UsePalette Then
    OldPaletteHandle = SelectPalette(frm3DMaze.hDC, PaletteHandle, 0)
    NumRealized = RealizePalette(frm3DMaze.hDC)
  End If
  XRelative = 1#
  YRelative = Sqrt3 / 2#
  CurrentColor = SolutionColor
  Call DrawLine(1#, 0#, XRelative, YRelative, XMax, XOffset, YMax, CosTilt, SinTilt, PixelsPerX, PixelsPerZ, RelDistOfUserFromScreen)
  XPrevious = 3
  YPrevious = -2
  X = 3
  Y = 2
  Do
    PathFound = False
    DeltaIndex = 0
    Do While (Not PathFound)
      XNext = X + HexDeltaX(DeltaIndex, 0)
      YNext = Y + HexDeltaY(DeltaIndex, 0)
      If Page(YNext, XNext) = 1 Then
        XNext = XNext + HexDeltaX(DeltaIndex, 0)
        YNext = YNext + HexDeltaY(DeltaIndex, 0)
        If (XNext <> XPrevious) Or (YNext <> YPrevious) Then
          PathFound = True
        Else
          DeltaIndex = DeltaIndex + 1
        End If
      Else
        DeltaIndex = DeltaIndex + 1
      End If
    Loop
    If YNext < MaxY Then
      Select Case YNext - Y
        Case -4
          XRelativeNext = XRelative
          YRelativeNext = YRelative - Sqrt3
        Case -2
          If XNext > X Then
            XRelativeNext = XRelative + 3# / 2#
            YRelativeNext = YRelative - Sqrt3 / 2#
          Else
            XRelativeNext = XRelative - 3# / 2#
            YRelativeNext = YRelative - Sqrt3 / 2#
          End If
        Case 2
          If XNext > X Then
            XRelativeNext = XRelative + 3# / 2#
            YRelativeNext = YRelative + Sqrt3 / 2#
          Else
            XRelativeNext = XRelative - 3# / 2#
            YRelativeNext = YRelative + Sqrt3 / 2#
          End If
        Case Else
          XRelativeNext = XRelative
          YRelativeNext = YRelative + Sqrt3
      End Select
      Call DrawLine(XRelative, YRelative, XRelativeNext, YRelativeNext, XMax, XOffset, YMax, CosTilt, SinTilt, PixelsPerX, PixelsPerZ, RelDistOfUserFromScreen)
    Else
      Call DrawLine(XRelative, YRelative, XRelative, YMax, XMax, XOffset, YMax, CosTilt, SinTilt, PixelsPerX, PixelsPerZ, RelDistOfUserFromScreen)
    End If
    XPrevious = X
    YPrevious = Y
    X = XNext
    Y = YNext
    XRelative = XRelativeNext
    YRelative = YRelativeNext
  Loop While YNext < MaxY
  If UsePalette Then
    NumRealized = SelectPalette(frm3DMaze.hDC, OldPaletteHandle, 0)
  End If
End Sub

Private Sub HexDisplayUserMoves(MaxX As Integer, MaxY As Integer, Page() As Byte, XMax As Double, XOffset As Double, YMax As Double, CosTilt As Double, SinTilt As Double, PixelsPerX As Double, PixelsPerZ As Double, RelDistOfUserFromScreen As Double)
  Dim DeltaIndex As Byte
  Dim EvenRow As Boolean
  Dim OldPaletteHandle As Long
  Dim X As Integer
  Dim XNext As Integer
  Dim XNextNext As Integer
  Dim XRelative As Double
  Dim XRelativeNext As Double
  Dim Y As Integer
  Dim YNext As Integer
  Dim YNextNext As Integer
  Dim YRelative As Double
  Dim YRelativeNext As Double
  If UsePalette Then
    OldPaletteHandle = SelectPalette(frm3DMaze.hDC, PaletteHandle, 0)
    NumRealized = RealizePalette(frm3DMaze.hDC)
  End If
  Y = 2
  YRelative = Sqrt3 / 2#
  EvenRow = False
  Do While (Y < MaxY)
    If EvenRow Then
      X = 7
      XRelative = 2.5
    Else
      X = 3
      XRelative = 1#
    End If
    Do While (X < MaxX)
      If ((Page(Y, X) = 1) Or (Page(Y, X) = 3)) Then
        For DeltaIndex = 0 To 5
          XNext = X + HexDeltaX(DeltaIndex, 0)
          YNext = Y + HexDeltaY(DeltaIndex, 0)
          If Page(YNext, XNext) <> 0 Then
            If YNext = 0 Then
              CurrentColor = AdvanceColor
              Call DrawLine(1#, 0#, XRelative, YRelative, XMax, XOffset, YMax, CosTilt, SinTilt, PixelsPerX, PixelsPerZ, RelDistOfUserFromScreen)
            Else
              If YNext = MaxY Then
                If UserHasSolved Then
                  CurrentColor = AdvanceColor
                  YRelativeNext = YRelative + Sqrt3 / 2#
                  Call DrawLine(XRelative, YRelative, XRelative, YRelativeNext, XMax, XOffset, YMax, CosTilt, SinTilt, PixelsPerX, PixelsPerZ, RelDistOfUserFromScreen)
                End If
              Else
                XNextNext = XNext + HexDeltaX(DeltaIndex, 0)
                If XNextNext > 0 Then
                  If XNextNext < MaxX Then
                    YNextNext = YNext + HexDeltaY(DeltaIndex, 0)
                    If YNextNext > 0 Then
                      If YNextNext < MaxY Then
                        If ((Page(YNextNext, XNextNext) = 1) Or (Page(YNextNext, XNextNext) = 3)) Then
                          If Page(Y, X) = Page(YNextNext, XNextNext) Then
                            If Page(Y, X) = 1 Then
                              CurrentColor = AdvanceColor
                            Else
                              CurrentColor = BackoutColor
                            End If
                          Else
                            CurrentColor = BackoutColor
                          End If
                          Select Case (YNext - Y)
                            Case -2
                              XRelativeNext = XRelative
                              YRelativeNext = YRelative - Sqrt3 / 2#
                            Case -1
                              If XNext > X Then
                                XRelativeNext = XRelative + 3# / 4#
                                YRelativeNext = YRelative - Sqrt3 / 4#
                              Else
                                XRelativeNext = XRelative - 3# / 4#
                                YRelativeNext = YRelative - Sqrt3 / 4#
                              End If
                            Case 1
                              If XNext > X Then
                                XRelativeNext = XRelative + 3# / 4#
                                YRelativeNext = YRelative + Sqrt3 / 4#
                              Else
                                XRelativeNext = XRelative - 3# / 4#
                                YRelativeNext = YRelative + Sqrt3 / 4#
                              End If
                            Case Else
                              XRelativeNext = XRelative
                              YRelativeNext = YRelative + Sqrt3 / 2#
                          End Select
                          Call DrawLine(XRelative, YRelative, XRelativeNext, YRelativeNext, XMax, XOffset, YMax, CosTilt, SinTilt, PixelsPerX, PixelsPerZ, RelDistOfUserFromScreen)
                        End If
                       End If
                    End If
                  End If
                End If
              End If
            End If
          End If
        Next DeltaIndex
      End If
      XRelative = XRelative + 3#
      X = X + 8
    Loop
    EvenRow = Not EvenRow
    YRelative = YRelative + Sqrt3 / 2#
    Y = Y + 2
  Loop
  If UsePalette Then
    NumRealized = SelectPalette(frm3DMaze.hDC, OldPaletteHandle, 0)
  End If
End Sub

Private Sub HexSolveMaze(Stack() As StackRec, Page() As Byte, NumRoomsInSolution As Integer, Adjacency As Integer, MaxX As Integer, MaxY As Integer)
  Dim DeltaIndex As Byte
  Dim PassageFound As Integer
  Dim StackHead As Integer
  Dim X As Integer
  Dim XNext As Integer
  Dim Y As Integer
  Dim YNext As Integer

  NumRoomsInSolution = 1
  Adjacency = 0
  X = 3
  Y = 2
  StackHead = -1
  Page(Y, X) = 1
  Do
    DeltaIndex = 0
    PassageFound = False
    Do
      Do While ((DeltaIndex < 6) And (Not PassageFound))
        XNext = X + HexDeltaX(DeltaIndex, 0)
        YNext = Y + HexDeltaY(DeltaIndex, 0)
        If Page(YNext, XNext) = 2 Then
          PassageFound = True
        Else
          DeltaIndex = DeltaIndex + 1
        End If
      Loop
      If Not PassageFound Then
        DeltaIndex = Stack(StackHead).Index1
        Page(Y, X) = 2
        X = X - HexDeltaX(DeltaIndex, 0)
        Y = Y - HexDeltaY(DeltaIndex, 0)
        Page(Y, X) = 2
        X = X - HexDeltaX(DeltaIndex, 0)
        Y = Y - HexDeltaY(DeltaIndex, 0)
        StackHead = StackHead - 1
        DeltaIndex = DeltaIndex + 1
      End If
    Loop While Not PassageFound
    Page(YNext, XNext) = 1
    XNext = XNext + HexDeltaX(DeltaIndex, 0)
    YNext = YNext + HexDeltaY(DeltaIndex, 0)
    If YNext <= MaxY Then
      StackHead = StackHead + 1
      Stack(StackHead).Index1 = DeltaIndex
      Page(YNext, XNext) = 1
      X = XNext
      Y = YNext
    End If
  Loop While YNext < MaxY
  X = MaxX - 3
  Y = MaxY - 2
  Adjacency = 0
  Do While (StackHead >= 0)
    For DeltaIndex = 0 To 5
      XNext = X + HexDeltaX(DeltaIndex, 0)
      YNext = Y + HexDeltaY(DeltaIndex, 0)
      If Page(YNext, XNext) <> 1 Then
        If Page(YNext, XNext) = 0 Then
          XNext = XNext + HexDeltaX(DeltaIndex, 0)
          YNext = YNext + HexDeltaY(DeltaIndex, 0)
          If XNext < 0 Then
            Adjacency = Adjacency + 1
          Else
            If XNext > MaxX Then
              Adjacency = Adjacency + 1
            Else
              If YNext < 0 Then
                Adjacency = Adjacency + 1
              Else
                If YNext > MaxY Then
                  Adjacency = Adjacency + 1
                Else
                  If Page(YNext, XNext) = 1 Then
                    Adjacency = Adjacency + 1
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
    Next DeltaIndex
    X = X - 2 * HexDeltaX(Stack(StackHead).Index1, 0)
    Y = Y - 2 * HexDeltaY(Stack(StackHead).Index1, 0)
    StackHead = StackHead - 1
    NumRoomsInSolution = NumRoomsInSolution + 1
  Loop
  For DeltaIndex = 0 To 5
    XNext = X + HexDeltaX(DeltaIndex, 0)
    YNext = X + HexDeltaY(DeltaIndex, 0)
    If Page(YNext, XNext) <> 2 Then
      If Page(YNext, XNext) = 0 Then
        XNext = XNext + HexDeltaX(DeltaIndex, 0)
        YNext = YNext + HexDeltaY(DeltaIndex, 0)
        If XNext < 0 Then
          Adjacency = Adjacency + 1
        Else
          If XNext > MaxX Then
            Adjacency = Adjacency + 1
          Else
            If YNext < 0 Then
              Adjacency = Adjacency + 1
            Else
              If YNext > MaxY Then
                Adjacency = Adjacency + 1
              Else
                If Page(YNext, XNext) = 1 Then
                  Adjacency = Adjacency + 1
                End If
              End If
            End If
          End If
        End If
      End If
    End If
  Next DeltaIndex
End Sub

Private Sub HexGenerateMaze(Page() As Byte, MaxX As Integer, MaxY As Integer, Stack() As StackRec, NumColumns As Integer, NumRows As Integer, Seed() As Byte)
  Dim ColumnNum As Integer
  Dim DeltaIndex1 As Integer
  Dim DeltaIndex2 As Integer
  Dim PassageFound As Integer
  Dim RN(7) As Integer
  Dim RNIndex1 As Integer
  Dim RNIndex2 As Integer
  Dim RowNum As Integer
  Dim SearchComplete As Integer
  Dim StackHead As Integer
  Dim TemInt As Integer
  Dim X As Integer
  Dim XMod8 As Byte
  Dim XNext As Integer
  Dim Y As Integer
  Dim YMod4 As Byte
  Dim YNext As Integer

  RN(0) = Seed(0) + 1
  RN(1) = Seed(1) + 1
  RN(2) = Seed(2) + 1
  RN(3) = Seed(3) + 1
  RN(4) = Seed(4) + 1
  RN(5) = Seed(5) + 1
  RN(6) = Seed(6) + 1
  RN(7) = Seed(7) + 1
  YMod4 = 1
  For Y = 0 To MaxY
    If YMod4 = 1 Then
      XMod8 = 1
      For X = 0 To MaxX
        If (((XMod8 = 0) And (Y <> 0) And (Y <> MaxY)) Or (XMod8 = 3) Or (XMod8 = 4) Or (XMod8 = 5)) Then
          Page(Y, X) = 0
        Else
          Page(Y, X) = 2
        End If
        XMod8 = XMod8 + 1
        If XMod8 >= 8 Then XMod8 = 0
      Next X
    Else
      If YMod4 = 0 Or YMod4 = 2 Then
        XMod8 = 1
        For X = 0 To MaxX
          If (XMod8 = 2) Or (XMod8 = 6) Then
            Page(Y, X) = 0
          Else
            Page(Y, X) = 2
          End If
          XMod8 = XMod8 + 1
          If XMod8 >= 8 Then XMod8 = 0
        Next X
      Else
        XMod8 = 1
        For X = 0 To MaxX
          If (XMod8 = 0) Or (XMod8 = 1) Or (XMod8 = 4) Or (XMod8 = 7) Then
            Page(Y, X) = 0
          Else
            Page(Y, X) = 2
          End If
          XMod8 = XMod8 + 1
          If XMod8 >= 8 Then XMod8 = 0
        Next X
      End If
    End If
    YMod4 = YMod4 + 1
    If YMod4 >= 4 Then YMod4 = 0
  Next Y
  ColumnNum = RN(0)
  RNIndex1 = 0
  RNIndex2 = 1
  Do While (RNIndex2 < 8)
    TemInt = RN(RNIndex2)
    RN(RNIndex1) = TemInt
    ColumnNum = ColumnNum + TemInt
    If ColumnNum >= 727 Then ColumnNum = ColumnNum - 727
    RNIndex1 = RNIndex2
    RNIndex2 = RNIndex2 + 1
  Loop
  RN(7) = ColumnNum
  ColumnNum = ColumnNum Mod NumColumns
  X = 4 * ColumnNum + 3
  RowNum = RN(0)
  RNIndex1 = 0
  RNIndex2 = 1
  Do While (RNIndex2 < 8)
    TemInt = RN(RNIndex2)
    RN(RNIndex1) = TemInt
    RowNum = RowNum + TemInt
    If RowNum >= 727 Then RowNum = RowNum - 727
    RNIndex1 = RNIndex2
    RNIndex2 = RNIndex2 + 1
  Loop
  RN(7) = RowNum
  If ColumnNum Mod 2 Then
    RowNum = RowNum Mod (NumRows - 1)
    Y = 4 * RowNum + 4
  Else
    RowNum = RowNum Mod NumRows
    Y = 4 * RowNum + 2
  End If
  Page(Y, X) = 2
  StackHead = -1
  Do
    DeltaIndex1 = 0
    Do
      DeltaIndex2 = RN(0)
      RNIndex1 = 0
      RNIndex2 = 1
      Do While (RNIndex2 < 8)
        TemInt = RN(RNIndex2)
        RN(RNIndex1) = TemInt
        DeltaIndex2 = DeltaIndex2 + TemInt
        If DeltaIndex2 >= 727 Then DeltaIndex2 = DeltaIndex2 - 727
        RNIndex1 = RNIndex2
        RNIndex2 = RNIndex2 + 1
      Loop
      RN(7) = DeltaIndex2
    Loop While DeltaIndex2 >= 720
    PassageFound = False
    SearchComplete = False
    Do While (Not SearchComplete)
      Do While ((DeltaIndex1 < 6) And (Not PassageFound))
        XNext = X + 2 * HexDeltaX(DeltaIndex1, DeltaIndex2)
        If XNext <= 0 Then
          DeltaIndex1 = DeltaIndex1 + 1
        Else
          If XNext > MaxX Then
            DeltaIndex1 = DeltaIndex1 + 1
          Else
            YNext = Y + 2 * HexDeltaY(DeltaIndex1, DeltaIndex2)
            If YNext <= 0 Then
              DeltaIndex1 = DeltaIndex1 + 1
            Else
              If YNext > MaxY Then
                DeltaIndex1 = DeltaIndex1 + 1
              Else
                If Page(YNext, XNext) = 0 Then
                  PassageFound = True
                Else
                  DeltaIndex1 = DeltaIndex1 + 1
                End If
              End If
            End If
          End If
        End If
      Loop
      If Not PassageFound Then
        If StackHead >= 0 Then
          DeltaIndex1 = Stack(StackHead).Index1
          DeltaIndex2 = Stack(StackHead).Index2
          X = X - 2 * HexDeltaX(DeltaIndex1, DeltaIndex2)
          Y = Y - 2 * HexDeltaY(DeltaIndex1, DeltaIndex2)
          StackHead = StackHead - 1
          DeltaIndex1 = DeltaIndex1 + 1
        End If
      End If
      If ((PassageFound) Or ((StackHead = -1) And (DeltaIndex1 >= 6))) Then
        SearchComplete = True
      Else
        SearchComplete = False
      End If
    Loop
    If PassageFound Then
      StackHead = StackHead + 1
      Stack(StackHead).Index1 = DeltaIndex1
      Stack(StackHead).Index2 = DeltaIndex2
      Page(YNext, XNext) = 2
      Page((Y + YNext) \ 2, (X + XNext) \ 2) = 2
      X = XNext
      Y = YNext
    End If
  Loop While StackHead <> -1
  Page(0, 3) = 1
  Page(MaxY, MaxX - 3) = 2
End Sub

Private Sub HexSelectMaze(Seed As String, Page() As Byte, MaxX As Integer, MaxY As Integer, Stack() As StackRec, NumRoomsInMaze As Integer, NumColumns As Integer, NumRows As Integer, SecondsForMazeSelection As Double)
  Dim Adjacency As Integer
  Dim Counter0 As Byte
  Dim Counter1 As Byte
  Dim Counter2 As Byte
  Dim Counter3 As Byte
  Dim Counter4 As Byte
  Dim Counter5 As Byte
  Dim Counter6 As Byte
  Dim Counter7 As Byte
  Dim ElapsedTime As Double
  Dim MinAdjacency As Integer
  Dim NumRoomsInSolution As Integer
  Dim NumRoomsInSolutionAtMin As Integer
  Dim RN(7) As Integer
  Dim RNIndex1 As Integer
  Dim RNIndex2 As Integer
  Dim SeedByte(7) As Byte
  Dim SeedByteAtMin(7) As Byte
  Dim SeedLength As Integer
  Dim StartTime As Double

  SeedLength = Len(Seed)
  If SeedLength > 8 Then SeedLength = 8
  RNIndex1 = 0
  For RNIndex2 = 1 To SeedLength
    RN(RNIndex1) = Asc(Mid$(Seed, RNIndex2, 1)) Mod 10
    RNIndex1 = RNIndex1 + 1
  Next RNIndex2
  RNIndex2 = 7
  Do While (RNIndex1 > 0)
    RNIndex1 = RNIndex1 - 1
    RN(RNIndex2) = RN(RNIndex1)
    RNIndex2 = RNIndex2 - 1
  Loop
  Do While (RNIndex2 >= 0)
    RN(RNIndex2) = 8
    RNIndex2 = RNIndex2 - 1
  Loop
  Counter0 = RN(0)
  Counter1 = RN(1)
  Counter2 = RN(2)
  Counter3 = RN(3)
  Counter4 = RN(4)
  Counter5 = RN(5)
  Counter6 = RN(6)
  Counter7 = RN(7)
  Call Hash(Counter0, Counter1, Counter2, Counter3, Counter4, Counter5, Counter6, Counter7)
  MinAdjacency = 4 * NumRoomsInMaze + 1
  NumRoomsInSolutionAtMin = 0
  SeedByteAtMin(0) = Counter0
  SeedByteAtMin(1) = Counter1
  SeedByteAtMin(2) = Counter2
  SeedByteAtMin(3) = Counter3
  SeedByteAtMin(4) = Counter4
  SeedByteAtMin(5) = Counter5
  SeedByteAtMin(6) = Counter6
  SeedByteAtMin(7) = Counter7
  StartTime = Timer
  Do
    SeedByte(0) = Counter0
    SeedByte(1) = Counter1
    SeedByte(2) = Counter2
    SeedByte(3) = Counter3
    SeedByte(4) = Counter4
    SeedByte(5) = Counter5
    SeedByte(6) = Counter6
    SeedByte(7) = Counter7
    Call HexGenerateMaze(Page(), MaxX, MaxY, Stack(), NumColumns, NumRows, SeedByte())
    Call HexSolveMaze(Stack(), Page(), NumRoomsInSolution, Adjacency, MaxX, MaxY)
    If 3 * NumRoomsInSolution >= NumRoomsInMaze Then
      If Adjacency < MinAdjacency Then
        MinAdjacency = Adjacency
        NumRoomsInSolutionAtMin = NumRoomsInSolution
        SeedByteAtMin(0) = SeedByte(0)
        SeedByteAtMin(1) = SeedByte(1)
        SeedByteAtMin(2) = SeedByte(2)
        SeedByteAtMin(3) = SeedByte(3)
        SeedByteAtMin(4) = SeedByte(4)
        SeedByteAtMin(5) = SeedByte(5)
        SeedByteAtMin(6) = SeedByte(6)
        SeedByteAtMin(7) = SeedByte(7)
      Else
        If Adjacency = MinAdjacency Then
          If NumRoomsInSolution > NumRoomsInSolutionAtMin Then
            NumRoomsInSolutionAtMin = NumRoomsInSolution
            SeedByteAtMin(0) = SeedByte(0)
            SeedByteAtMin(1) = SeedByte(1)
            SeedByteAtMin(2) = SeedByte(2)
            SeedByteAtMin(3) = SeedByte(3)
            SeedByteAtMin(4) = SeedByte(4)
            SeedByteAtMin(5) = SeedByte(5)
            SeedByteAtMin(6) = SeedByte(6)
            SeedByteAtMin(7) = SeedByte(7)
          End If
        End If
      End If
    End If
    Call Increment(Counter0, Counter1, Counter2, Counter3, Counter4, Counter5, Counter6, Counter7)
    ElapsedTime = Timer - StartTime
  Loop While ((ElapsedTime >= 0#) And (ElapsedTime < SecondsForMazeSelection))
  Call HexGenerateMaze(Page(), MaxX, MaxY, Stack(), NumColumns, NumRows, SeedByteAtMin())
  Call HexSolveMaze(Stack(), Page(), NumRoomsInSolution, Adjacency, MaxX, MaxY)
End Sub

Private Sub HexOutputMaze()
  Dim ObjectNum As Byte
  Dim Radians As Double
  Dim RadiansPerDegree As Double
  Dim SingleRectangle(3) As VertexRec
  Dim SingleTriangle(2) As VertexRec
  Dim TemDouble1 As Double
  Dim TemDouble2 As Double
  Dim TemDouble3 As Double
  Dim TemDouble4 As Double
  Dim Triangle(3, 2) As VertexRec
  Dim VertexNum As Byte
  Dim XMod8 As Byte
  Dim X0 As Double
  Dim X1 As Double
  Dim X2 As Double
  Dim X3 As Double
  Dim Y0 As Double
  Dim Y1 As Double
  Dim Y2 As Double
  Dim Y3 As Double

  Select Case State
    Case 0
      Text1.Text = ""
      ScaleMode = 1
      If (Resize) Then
        TemDouble1 = ScaleWidth - VScroll1.Width
        TemDouble2 = MinWallLengthInInches
        TemDouble2 = 1440# * TemDouble2
        TemDouble3 = RelativeWidthOfWall
        NumColumns = Int(2# * (TemDouble1 / TemDouble2 - 2# - TemDouble3 / Sqrt3) / 3# + 1#)
        If NumColumns Mod 2 = 0 Then NumColumns = NumColumns - 1
        If NumColumns < 3 Then NumColumns = 3
        TemDouble1 = ScaleHeight - Text1.Height
        TemDouble2 = ScaleWidth - VScroll1.Width
        ScaleMode = 3
        TemDouble3 = NumColumns
        TemDouble4 = RelativeWidthOfWall
        NumRows = Int(((TemDouble1 / TemDouble2) * (3# * (TemDouble3 - 1#) / 2# + 2# + TemDouble4 / Sqrt3) - TemDouble4) / Sqrt3)
        If NumRows < 2 Then NumRows = 2
        Tilt = 90 - VScroll1.Value
        MaxX = 8 * (NumColumns \ 2) + 6
        MaxY = 4 * NumRows
        NumRoomsInMaze = NumRows * NumColumns - (NumColumns \ 2)
        ReDim ComputerPage(MaxY, MaxX)
        ReDim UserPage(MaxY, MaxX)
        ReDim Stack(NumRoomsInMaze)
        Call HexSelectMaze(Seed, ComputerPage(), MaxX, MaxY, Stack(), NumRoomsInMaze, NumColumns, NumRows, SecondsForMazeSelection)
        For UserX = 0 To MaxX
          For UserY = 0 To MaxY
            If ComputerPage(UserY, UserX) = 0 Then
              UserPage(UserY, UserX) = 0
            Else
              UserPage(UserY, UserX) = 2
            End If
          Next UserY
        Next UserX
        UserX = 3
        UserXRelative = 1#
        UserY = 2
        UserYRelative = Sqrt3 / 2#
        UserPage(UserY, UserX) = 1
        Resize = False
      End If
      If (Paint) Then
        ScaleMode = 3
        Cls
        RadiansPerDegree = Atn(1#) / 45#
        Radians = Tilt * RadiansPerDegree
        SinTilt = Sin(Radians)
        CosTilt = Cos(Radians)
        TemDouble1 = NumColumns
        XMax = 3# * (TemDouble1 - 1#) / 2# + 2# + RelativeWidthOfWall / Sqrt3
        TemDouble1 = ScaleWidth - VScroll1.Width
        PixelsPerX = (TemDouble1 - 1#) / (XMax * (XMax / (XMax - RelativeHeightOfWall)))
        XOffset = (XMax / 2#) * (RelativeHeightOfWall / (XMax - RelativeHeightOfWall))
        TemDouble1 = NumRows
        YMax = TemDouble1 * Sqrt3 + RelativeWidthOfWall
        TemDouble1 = ScaleHeight - Text1.Height
        PixelsPerZ = (TemDouble1 - 1#) / Sqr(YMax * YMax + RelativeHeightOfWall * RelativeHeightOfWall)
        If YMax > XMax Then
          RelDistOfUserFromScreen = YMax
        Else
          RelDistOfUserFromScreen = XMax
        End If
        Paint = False
      End If
      If State = 0 Then
        State = 1
        DoEvents
        If State < 5 Then
          Timer1.Enabled = True
        End If
      End If
    Case 1
      BaseTriangle(0, 0).X = 0#
      BaseTriangle(0, 0).Y = RelativeWidthOfWall + Sqrt3 / 2#
      BaseTriangle(0, 1).X = 0#
      BaseTriangle(0, 1).Y = Sqrt3 / 2#
      BaseTriangle(0, 2).X = RelativeWidthOfWall * Sqrt3 / 2#
      BaseTriangle(0, 2).Y = (RelativeWidthOfWall + Sqrt3) / 2#
      BaseTriangle(1, 0).X = (1# - RelativeWidthOfWall / Sqrt3) / 2#
      BaseTriangle(1, 0).Y = RelativeWidthOfWall / 2#
      BaseTriangle(1, 1).X = 0.5 + RelativeWidthOfWall / Sqrt3
      BaseTriangle(1, 1).Y = 0#
      BaseTriangle(1, 2).X = BaseTriangle(1, 1).X
      BaseTriangle(1, 2).Y = RelativeWidthOfWall
      BaseTriangle(2, 0).X = 1.5
      BaseTriangle(2, 0).Y = RelativeWidthOfWall
      BaseTriangle(2, 1).X = 1.5
      BaseTriangle(2, 1).Y = 0#
      BaseTriangle(2, 2).X = 1.5 * (1# + RelativeWidthOfWall / Sqrt3)
      BaseTriangle(2, 2).Y = RelativeWidthOfWall / 2#
      BaseTriangle(3, 0).X = 2# - RelativeWidthOfWall / (2# * Sqrt3)
      BaseTriangle(3, 0).Y = BaseTriangle(0, 2).Y
      BaseTriangle(3, 1).X = 2# + RelativeWidthOfWall / Sqrt3
      BaseTriangle(3, 1).Y = BaseTriangle(0, 1).Y
      BaseTriangle(3, 2).X = BaseTriangle(3, 1).X
      BaseTriangle(3, 2).Y = BaseTriangle(0, 0).Y
      BaseRectangle(0, 0).X = BaseTriangle(0, 2).X
      BaseRectangle(0, 0).Y = BaseTriangle(0, 2).Y
      BaseRectangle(0, 1).X = BaseTriangle(1, 1).X
      BaseRectangle(0, 1).Y = Sqrt3
      BaseRectangle(0, 2).X = BaseTriangle(1, 0).X
      BaseRectangle(0, 2).Y = Sqrt3 + RelativeWidthOfWall / 2#
      BaseRectangle(0, 3).X = BaseTriangle(0, 0).X
      BaseRectangle(0, 3).Y = BaseTriangle(0, 0).Y
      BaseRectangle(1, 0).X = BaseTriangle(0, 1).X
      BaseRectangle(1, 0).Y = BaseTriangle(0, 1).Y
      BaseRectangle(1, 1).X = BaseTriangle(1, 0).X
      BaseRectangle(1, 1).Y = BaseTriangle(1, 0).Y
      BaseRectangle(1, 2).X = BaseTriangle(1, 2).X
      BaseRectangle(1, 2).Y = BaseTriangle(1, 2).Y
      BaseRectangle(1, 3).X = BaseTriangle(0, 2).X
      BaseRectangle(1, 3).Y = BaseTriangle(0, 2).Y
      BaseRectangle(2, 0).X = BaseTriangle(1, 1).X
      BaseRectangle(2, 0).Y = BaseTriangle(1, 1).Y
      BaseRectangle(2, 1).X = BaseTriangle(2, 1).X
      BaseRectangle(2, 1).Y = BaseTriangle(2, 1).Y
      BaseRectangle(2, 2).X = BaseTriangle(2, 0).X
      BaseRectangle(2, 2).Y = BaseTriangle(2, 0).Y
      BaseRectangle(2, 3).X = BaseTriangle(1, 2).X
      BaseRectangle(2, 3).Y = BaseTriangle(1, 2).Y
      BaseRectangle(3, 0).X = BaseTriangle(2, 2).X
      BaseRectangle(3, 0).Y = BaseTriangle(2, 2).Y
      BaseRectangle(3, 1).X = BaseTriangle(3, 1).X
      BaseRectangle(3, 1).Y = BaseTriangle(3, 1).Y
      BaseRectangle(3, 2).X = BaseTriangle(3, 0).X
      BaseRectangle(3, 2).Y = BaseTriangle(3, 0).Y
      BaseRectangle(3, 3).X = BaseTriangle(2, 0).X
      BaseRectangle(3, 3).Y = BaseTriangle(2, 0).Y
      BaseRectangle(4, 0).X = BaseTriangle(3, 1).X
      BaseRectangle(4, 0).Y = BaseTriangle(3, 1).Y
      BaseRectangle(4, 1).X = BaseTriangle(3, 1).X + (BaseTriangle(2, 1).X - BaseTriangle(1, 1).X)
      BaseRectangle(4, 1).Y = BaseTriangle(3, 1).Y
      BaseRectangle(4, 2).X = BaseRectangle(4, 1).X
      BaseRectangle(4, 2).Y = BaseTriangle(3, 2).Y
      BaseRectangle(4, 3).X = BaseTriangle(3, 2).X
      BaseRectangle(4, 3).Y = BaseTriangle(3, 2).Y
      BaseRectangle(5, 0).X = BaseRectangle(0, 1).X + (BaseTriangle(2, 1).X - BaseTriangle(1, 1).X)
      BaseRectangle(5, 0).Y = BaseRectangle(0, 1).Y
      BaseRectangle(5, 1).X = BaseTriangle(3, 0).X
      BaseRectangle(5, 1).Y = BaseTriangle(3, 0).Y
      BaseRectangle(5, 2).X = BaseTriangle(3, 2).X
      BaseRectangle(5, 2).Y = BaseTriangle(3, 2).Y
      BaseRectangle(5, 3).X = BaseRectangle(0, 2).X + (BaseTriangle(2, 2).X - BaseTriangle(1, 0).X)
      BaseRectangle(5, 3).Y = BaseRectangle(0, 2).Y
      Rectangle(0, 0).X = BaseTriangle(1, 1).X
      Rectangle(0, 0).Y = BaseTriangle(1, 1).Y
      Rectangle(0, 1).X = XMax - BaseTriangle(1, 1).X
      Rectangle(0, 1).Y = BaseTriangle(1, 1).Y
      Rectangle(0, 2).X = XMax - BaseTriangle(1, 2).X
      Rectangle(0, 2).Y = BaseTriangle(1, 2).Y
      Rectangle(0, 3).X = BaseTriangle(1, 2).X
      Rectangle(0, 3).Y = BaseTriangle(1, 2).Y
      Rectangle(1, 0).X = BaseTriangle(0, 1).X
      Rectangle(1, 0).Y = BaseTriangle(0, 1).Y
      Rectangle(1, 1).X = XMax - BaseTriangle(0, 1).X
      Rectangle(1, 1).Y = BaseTriangle(0, 1).Y
      Rectangle(1, 2).X = XMax - BaseTriangle(1, 2).X
      Rectangle(1, 2).Y = BaseTriangle(1, 2).Y
      Rectangle(1, 3).X = BaseTriangle(1, 2).X
      Rectangle(1, 3).Y = BaseTriangle(1, 2).Y
      Rectangle(2, 0).X = BaseTriangle(0, 1).X
      Rectangle(2, 0).Y = BaseTriangle(0, 1).Y
      Rectangle(2, 1).X = XMax - BaseTriangle(0, 1).X
      Rectangle(2, 1).Y = BaseTriangle(0, 1).Y
      Rectangle(2, 2).X = XMax - BaseTriangle(0, 0).X
      Rectangle(2, 2).Y = BaseTriangle(0, 0).Y
      Rectangle(2, 3).X = BaseTriangle(0, 0).X
      Rectangle(2, 3).Y = BaseTriangle(0, 0).Y
      Rectangle(3, 0).X = BaseTriangle(0, 0).X
      Rectangle(3, 0).Y = BaseTriangle(0, 0).Y
      Rectangle(3, 1).X = XMax - BaseTriangle(0, 0).X
      Rectangle(3, 1).Y = BaseTriangle(0, 0).Y
      Rectangle(3, 2).X = XMax - BaseRectangle(0, 1).X
      Rectangle(3, 2).Y = BaseRectangle(0, 1).Y
      Rectangle(3, 3).X = BaseRectangle(0, 1).X
      Rectangle(3, 3).Y = BaseRectangle(0, 1).Y
      Y = 0
      State = 2
      DoEvents
      If State < 5 Then
        Timer1.Enabled = True
      End If
    Case 2
      If UsePalette Then
        OldPaletteHandle = SelectPalette(frm3DMaze.hDC, PaletteHandle, 0)
        NumRealized = RealizePalette(frm3DMaze.hDC)
      End If
      If (Y <= MaxY - 1) Then
        If Y > 0 Then
          X0 = Rectangle(0, 0).X
          Y0 = Rectangle(0, 0).Y
          X1 = Rectangle(0, 1).X
          Y1 = Rectangle(0, 1).Y
          X2 = Rectangle(0, 2).X
          Y2 = Rectangle(0, 2).Y
          X3 = Rectangle(0, 3).X
          Y3 = Rectangle(0, 3).Y
          Call DisplayQuadrilateral(XMax, XOffset, YMax, X0, Y0, 0#, X1, Y1, 0#, X2, Y2, 0#, X3, Y3, 0#, PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen, FloorColor)
          X0 = Rectangle(1, 0).X
          Y0 = Rectangle(1, 0).Y
          X1 = Rectangle(1, 1).X
          Y1 = Rectangle(1, 1).Y
          X2 = Rectangle(1, 2).X
          Y2 = Rectangle(1, 2).Y
          X3 = Rectangle(1, 3).X
          Y3 = Rectangle(1, 3).Y
          Call DisplayQuadrilateral(XMax, XOffset, YMax, X0, Y0, 0#, X1, Y1, 0#, X2, Y2, 0#, X3, Y3, 0#, PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen, FloorColor)
        End If
        X0 = Rectangle(2, 0).X
        Y0 = Rectangle(2, 0).Y
        X1 = Rectangle(2, 1).X
        Y1 = Rectangle(2, 1).Y
        X2 = Rectangle(2, 2).X
        Y2 = Rectangle(2, 2).Y
        X3 = Rectangle(2, 3).X
        Y3 = Rectangle(2, 3).Y
        Call DisplayQuadrilateral(XMax, XOffset, YMax, X0, Y0, 0#, X1, Y1, 0#, X2, Y2, 0#, X3, Y3, 0#, PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen, FloorColor)
        If Y < MaxY - 4 Then
          X0 = Rectangle(3, 0).X
          Y0 = Rectangle(3, 0).Y
          X1 = Rectangle(3, 1).X
          Y1 = Rectangle(3, 1).Y
          X2 = Rectangle(3, 2).X
          Y2 = Rectangle(3, 2).Y
          X3 = Rectangle(3, 3).X
          Y3 = Rectangle(3, 3).Y
          Call DisplayQuadrilateral(XMax, XOffset, YMax, X0, Y0, 0#, X1, Y1, 0#, X2, Y2, 0#, X3, Y3, 0#, PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen, FloorColor)
        End If
        For ObjectNum = 0 To 3
          For VertexNum = 0 To 3
            Rectangle(ObjectNum, VertexNum).Y = Rectangle(ObjectNum, VertexNum).Y + Sqrt3
          Next VertexNum
        Next ObjectNum
        Y = Y + 4
      Else
        Rectangle(0, 0).X = BaseTriangle(1, 0).X
        Rectangle(0, 0).Y = BaseTriangle(1, 0).Y
        Rectangle(0, 1).X = BaseTriangle(1, 1).X
        Rectangle(0, 1).Y = BaseTriangle(1, 1).Y
        Rectangle(0, 2).X = BaseTriangle(2, 1).X
        Rectangle(0, 2).Y = BaseTriangle(2, 1).Y
        Rectangle(0, 3).X = BaseTriangle(2, 2).X
        Rectangle(0, 3).Y = BaseTriangle(2, 2).Y
        Rectangle(1, 0).X = BaseTriangle(0, 1).X
        Rectangle(1, 0).Y = BaseTriangle(0, 1).Y
        Rectangle(1, 1).X = BaseTriangle(1, 0).X
        Rectangle(1, 1).Y = BaseTriangle(1, 0).Y
        Rectangle(1, 2).X = BaseTriangle(2, 2).X
        Rectangle(1, 2).Y = BaseTriangle(2, 2).Y
        Rectangle(1, 3).X = BaseTriangle(3, 1).X
        Rectangle(1, 3).Y = BaseTriangle(3, 1).Y
        Rectangle(2, 0).X = BaseTriangle(0, 0).X
        Rectangle(2, 0).Y = BaseTriangle(0, 0).Y
        Rectangle(2, 1).X = BaseTriangle(0, 1).X
        Rectangle(2, 1).Y = BaseTriangle(0, 1).Y
        Rectangle(2, 2).X = BaseTriangle(3, 1).X
        Rectangle(2, 2).Y = BaseTriangle(3, 1).Y
        Rectangle(2, 3).X = BaseTriangle(3, 2).X
        Rectangle(2, 3).Y = BaseTriangle(3, 2).Y
        X = 0
        State = 3
      End If
      If UsePalette Then
        NumRealized = SelectPalette(frm3DMaze.hDC, OldPaletteHandle, 0)
      End If
      DoEvents
      If State < 5 Then
        Timer1.Enabled = True
      End If
    Case 3
      If UsePalette Then
        OldPaletteHandle = SelectPalette(frm3DMaze.hDC, PaletteHandle, 0)
        NumRealized = RealizePalette(frm3DMaze.hDC)
      End If
      If X <= MaxX Then
        For ObjectNum = 0 To 2
          X0 = Rectangle(ObjectNum, 0).X
          Y0 = Rectangle(ObjectNum, 0).Y
          X1 = Rectangle(ObjectNum, 1).X
          Y1 = Rectangle(ObjectNum, 1).Y
          X2 = Rectangle(ObjectNum, 2).X
          Y2 = Rectangle(ObjectNum, 2).Y
          X3 = Rectangle(ObjectNum, 3).X
          Y3 = Rectangle(ObjectNum, 3).Y
          Call DisplayQuadrilateral(XMax, XOffset, YMax, X0, Y0, 0#, X1, Y1, 0#, X2, Y2, 0#, X3, Y3, 0#, PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen, FloorColor)
          X0 = Rectangle(ObjectNum, 0).X
          Y0 = YMax - Rectangle(ObjectNum, 0).Y
          X1 = Rectangle(ObjectNum, 1).X
          Y1 = YMax - Rectangle(ObjectNum, 1).Y
          X2 = Rectangle(ObjectNum, 2).X
          Y2 = YMax - Rectangle(ObjectNum, 2).Y
          X3 = Rectangle(ObjectNum, 3).X
          Y3 = YMax - Rectangle(ObjectNum, 3).Y
          Call DisplayQuadrilateral(XMax, XOffset, YMax, X0, Y0, 0#, X1, Y1, 0#, X2, Y2, 0#, X3, Y3, 0#, PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen, FloorColor)
          For VertexNum = 0 To 3
            Rectangle(ObjectNum, VertexNum).X = Rectangle(ObjectNum, VertexNum).X + 3#
          Next VertexNum
        Next ObjectNum
        X = X + 8
      Else
        YMod4 = 0
        YOffset = 0#
        Y = 0
        State = 4
      End If
      If UsePalette Then
        NumRealized = SelectPalette(frm3DMaze.hDC, OldPaletteHandle, 0)
      End If
      DoEvents
      If State < 5 Then
        Timer1.Enabled = True
      End If
    Case 4
      If UsePalette Then
        OldPaletteHandle = SelectPalette(frm3DMaze.hDC, PaletteHandle, 0)
        NumRealized = RealizePalette(frm3DMaze.hDC)
      End If
      If Y <= MaxY Then
        Select Case YMod4
          Case 0
            XMod8 = 0
            For ObjectNum = 1 To 2
              For VertexNum = 0 To 2
                Triangle(ObjectNum, VertexNum).X = BaseTriangle(ObjectNum, VertexNum).X
                Triangle(ObjectNum, VertexNum).Y = BaseTriangle(ObjectNum, VertexNum).Y + YOffset
              Next VertexNum
            Next ObjectNum
            For VertexNum = 0 To 3
              Rectangle(2, VertexNum).X = BaseRectangle(2, VertexNum).X
              Rectangle(2, VertexNum).Y = BaseRectangle(2, VertexNum).Y + YOffset
            Next VertexNum
            For X = 0 To MaxX
              Select Case XMod8
                Case 2
                  SingleTriangle(0).X = Triangle(1, 0).X
                  SingleTriangle(0).Y = Triangle(1, 0).Y
                  SingleTriangle(1).X = Triangle(1, 1).X
                  SingleTriangle(1).Y = Triangle(1, 1).Y
                  SingleTriangle(2).X = Triangle(1, 2).X
                  SingleTriangle(2).Y = Triangle(1, 2).Y
                  Call OutputTriangle(XMax, XOffset, YMax, SingleTriangle(), PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen, True, TriangleSSWNNEColor)
                Case 4
                  SingleTriangle(0).X = Triangle(2, 0).X
                  SingleTriangle(0).Y = Triangle(2, 0).Y
                  SingleTriangle(1).X = Triangle(2, 1).X
                  SingleTriangle(1).Y = Triangle(2, 1).Y
                  SingleTriangle(2).X = Triangle(2, 2).X
                  SingleTriangle(2).Y = Triangle(2, 2).Y
                  Call OutputTriangle(XMax, XOffset, YMax, SingleTriangle(), PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen, True, TriangleSSENNWColor)
                Case Else
              End Select
              XMod8 = XMod8 + 1
              If XMod8 >= 8 Then
                XMod8 = 0
                For ObjectNum = 1 To 2
                  For VertexNum = 0 To 2
                    Triangle(ObjectNum, VertexNum).X = Triangle(ObjectNum, VertexNum).X + 3#
                  Next VertexNum
                Next ObjectNum
                For VertexNum = 0 To 3
                  Rectangle(2, VertexNum).X = Rectangle(2, VertexNum).X + 3#
                Next VertexNum
              End If
            Next X
            XMod8 = 0
            For ObjectNum = 1 To 2
              For VertexNum = 0 To 2
                Triangle(ObjectNum, VertexNum).X = BaseTriangle(ObjectNum, VertexNum).X
                Triangle(ObjectNum, VertexNum).Y = BaseTriangle(ObjectNum, VertexNum).Y + YOffset
              Next VertexNum
            Next ObjectNum
            For VertexNum = 0 To 3
              Rectangle(2, VertexNum).X = BaseRectangle(2, VertexNum).X
              Rectangle(2, VertexNum).Y = BaseRectangle(2, VertexNum).Y + YOffset
            Next VertexNum
            For X = 0 To MaxX
              Select Case XMod8
                Case 2
                  SingleTriangle(0).X = Triangle(1, 0).X
                  SingleTriangle(0).Y = Triangle(1, 0).Y
                  SingleTriangle(1).X = Triangle(1, 1).X
                  SingleTriangle(1).Y = Triangle(1, 1).Y
                  SingleTriangle(2).X = Triangle(1, 2).X
                  SingleTriangle(2).Y = Triangle(1, 2).Y
                  Call OutputTriangle(XMax, XOffset, YMax, SingleTriangle(), PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen, False, TriangleSENWColor)
                Case 3
                  If ComputerPage(Y, X) = 0 Then
                    SingleRectangle(0).X = Rectangle(2, 0).X
                    SingleRectangle(0).Y = Rectangle(2, 0).Y
                    SingleRectangle(1).X = Rectangle(2, 1).X
                    SingleRectangle(1).Y = Rectangle(2, 1).Y
                    SingleRectangle(2).X = Rectangle(2, 2).X
                    SingleRectangle(2).Y = Rectangle(2, 2).Y
                    SingleRectangle(3).X = Rectangle(2, 3).X
                    SingleRectangle(3).Y = Rectangle(2, 3).Y
                    Call OutputRectangle(XMax, XOffset, YMax, SingleRectangle(), PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen, RectangleWEColor)
                  End If
                Case 4
                  SingleTriangle(0).X = Triangle(2, 0).X
                  SingleTriangle(0).Y = Triangle(2, 0).Y
                  SingleTriangle(1).X = Triangle(2, 1).X
                  SingleTriangle(1).Y = Triangle(2, 1).Y
                  SingleTriangle(2).X = Triangle(2, 2).X
                  SingleTriangle(2).Y = Triangle(2, 2).Y
                  Call OutputTriangle(XMax, XOffset, YMax, SingleTriangle(), PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen, False, TriangleSWNEColor)
                Case Else
              End Select
              XMod8 = XMod8 + 1
              If XMod8 >= 8 Then
                XMod8 = 0
                For ObjectNum = 1 To 2
                  For VertexNum = 0 To 2
                    Triangle(ObjectNum, VertexNum).X = Triangle(ObjectNum, VertexNum).X + 3#
                  Next VertexNum
                Next ObjectNum
                For VertexNum = 0 To 3
                  Rectangle(2, VertexNum).X = Rectangle(2, VertexNum).X + 3#
                Next VertexNum
              End If
            Next X
          Case 1
            XMod8 = 0
            For ObjectNum = 1 To 3 Step 2
              For VertexNum = 0 To 3
                Rectangle(ObjectNum, VertexNum).X = BaseRectangle(ObjectNum, VertexNum).X
                Rectangle(ObjectNum, VertexNum).Y = BaseRectangle(ObjectNum, VertexNum).Y + YOffset
              Next VertexNum
            Next ObjectNum
            For X = 0 To MaxX
              Select Case XMod8
                Case 1
                  If ComputerPage(Y, X) = 0 Then
                    SingleRectangle(0).X = Rectangle(1, 0).X
                    SingleRectangle(0).Y = Rectangle(1, 0).Y
                    SingleRectangle(1).X = Rectangle(1, 1).X
                    SingleRectangle(1).Y = Rectangle(1, 1).Y
                    SingleRectangle(2).X = Rectangle(1, 2).X
                    SingleRectangle(2).Y = Rectangle(1, 2).Y
                    SingleRectangle(3).X = Rectangle(1, 3).X
                    SingleRectangle(3).Y = Rectangle(1, 3).Y
                    Call OutputRectangle(XMax, XOffset, YMax, SingleRectangle(), PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen, RectangleSWNEColor)
                  End If
                Case 5
                  If ComputerPage(Y, X) = 0 Then
                    SingleRectangle(0).X = Rectangle(3, 0).X
                    SingleRectangle(0).Y = Rectangle(3, 0).Y
                    SingleRectangle(1).X = Rectangle(3, 1).X
                    SingleRectangle(1).Y = Rectangle(3, 1).Y
                    SingleRectangle(2).X = Rectangle(3, 2).X
                    SingleRectangle(2).Y = Rectangle(3, 2).Y
                    SingleRectangle(3).X = Rectangle(3, 3).X
                    SingleRectangle(3).Y = Rectangle(3, 3).Y
                    Call OutputRectangle(XMax, XOffset, YMax, SingleRectangle(), PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen, RectangleSENWColor)
                  End If
                Case Else
              End Select
              XMod8 = XMod8 + 1
              If XMod8 >= 8 Then
                XMod8 = 0
                For ObjectNum = 1 To 3 Step 2
                  For VertexNum = 0 To 3
                    Rectangle(ObjectNum, VertexNum).X = Rectangle(ObjectNum, VertexNum).X + 3#
                  Next VertexNum
                Next ObjectNum
              End If
            Next X
          Case 2
            XMod8 = 0
            For ObjectNum = 0 To 3 Step 3
              For VertexNum = 0 To 2
                Triangle(ObjectNum, VertexNum).X = BaseTriangle(ObjectNum, VertexNum).X
                Triangle(ObjectNum, VertexNum).Y = BaseTriangle(ObjectNum, VertexNum).Y + YOffset
              Next VertexNum
            Next ObjectNum
            For VertexNum = 0 To 3
              Rectangle(4, VertexNum).X = BaseRectangle(4, VertexNum).X
              Rectangle(4, VertexNum).Y = BaseRectangle(4, VertexNum).Y + YOffset
            Next VertexNum
            For X = 0 To MaxX
              Select Case XMod8
                Case 0
                  SingleTriangle(0).X = Triangle(0, 0).X
                  SingleTriangle(0).Y = Triangle(0, 0).Y
                  SingleTriangle(1).X = Triangle(0, 1).X
                  SingleTriangle(1).Y = Triangle(0, 1).Y
                  SingleTriangle(2).X = Triangle(0, 2).X
                  SingleTriangle(2).Y = Triangle(0, 2).Y
                  Call OutputTriangle(XMax, XOffset, YMax, SingleTriangle(), PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen, True, TriangleSSWNNEColor)
                Case 6
                  SingleTriangle(0).X = Triangle(3, 0).X
                  SingleTriangle(0).Y = Triangle(3, 0).Y
                  SingleTriangle(1).X = Triangle(3, 1).X
                  SingleTriangle(1).Y = Triangle(3, 1).Y
                  SingleTriangle(2).X = Triangle(3, 2).X
                  SingleTriangle(2).Y = Triangle(3, 2).Y
                  Call OutputTriangle(XMax, XOffset, YMax, SingleTriangle(), PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen, True, TriangleSSENNWColor)
                Case Else
              End Select
              XMod8 = XMod8 + 1
              If XMod8 >= 8 Then
                XMod8 = 0
                For ObjectNum = 0 To 3 Step 3
                  For VertexNum = 0 To 2
                    Triangle(ObjectNum, VertexNum).X = Triangle(ObjectNum, VertexNum).X + 3#
                  Next VertexNum
                Next ObjectNum
                For VertexNum = 0 To 3
                  Rectangle(4, VertexNum).X = Rectangle(4, VertexNum).X + 3#
                Next VertexNum
              End If
            Next X
            XMod8 = 0
            For ObjectNum = 0 To 3 Step 3
              For VertexNum = 0 To 2
                Triangle(ObjectNum, VertexNum).X = BaseTriangle(ObjectNum, VertexNum).X
                Triangle(ObjectNum, VertexNum).Y = BaseTriangle(ObjectNum, VertexNum).Y + YOffset
              Next VertexNum
            Next ObjectNum
            For VertexNum = 0 To 3
              Rectangle(4, VertexNum).X = BaseRectangle(4, VertexNum).X
              Rectangle(4, VertexNum).Y = BaseRectangle(4, VertexNum).Y + YOffset
            Next VertexNum
            For X = 0 To MaxX
              Select Case XMod8
                Case 0
                  SingleTriangle(0).X = Triangle(0, 0).X
                  SingleTriangle(0).Y = Triangle(0, 0).Y
                  SingleTriangle(1).X = Triangle(0, 1).X
                  SingleTriangle(1).Y = Triangle(0, 1).Y
                  SingleTriangle(2).X = Triangle(0, 2).X
                  SingleTriangle(2).Y = Triangle(0, 2).Y
                  Call OutputTriangle(XMax, XOffset, YMax, SingleTriangle(), PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen, False, TriangleSWNEColor)
                Case 6
                  SingleTriangle(0).X = Triangle(3, 0).X
                  SingleTriangle(0).Y = Triangle(3, 0).Y
                  SingleTriangle(1).X = Triangle(3, 1).X
                  SingleTriangle(1).Y = Triangle(3, 1).Y
                  SingleTriangle(2).X = Triangle(3, 2).X
                  SingleTriangle(2).Y = Triangle(3, 2).Y
                  Call OutputTriangle(XMax, XOffset, YMax, SingleTriangle(), PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen, False, TriangleSENWColor)
                Case 7
                  If ComputerPage(Y, X) = 0 Then
                    SingleRectangle(0).X = Rectangle(4, 0).X
                    SingleRectangle(0).Y = Rectangle(4, 0).Y
                    SingleRectangle(1).X = Rectangle(4, 1).X
                    SingleRectangle(1).Y = Rectangle(4, 1).Y
                    SingleRectangle(2).X = Rectangle(4, 2).X
                    SingleRectangle(2).Y = Rectangle(4, 2).Y
                    SingleRectangle(3).X = Rectangle(4, 3).X
                    SingleRectangle(3).Y = Rectangle(4, 3).Y
                    Call OutputRectangle(XMax, XOffset, YMax, SingleRectangle(), PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen, RectangleWEColor)
                  End If
                Case Else
              End Select
              XMod8 = XMod8 + 1
              If XMod8 >= 8 Then
                XMod8 = 0
                For ObjectNum = 0 To 3 Step 3
                  For VertexNum = 0 To 2
                    Triangle(ObjectNum, VertexNum).X = Triangle(ObjectNum, VertexNum).X + 3#
                  Next VertexNum
                Next ObjectNum
                For VertexNum = 0 To 3
                  Rectangle(4, VertexNum).X = Rectangle(4, VertexNum).X + 3#
                Next VertexNum
              End If
            Next X
          Case Else
            XMod8 = 0
            For ObjectNum = 0 To 5 Step 5
              For VertexNum = 0 To 3
                Rectangle(ObjectNum, VertexNum).X = BaseRectangle(ObjectNum, VertexNum).X
                Rectangle(ObjectNum, VertexNum).Y = BaseRectangle(ObjectNum, VertexNum).Y + YOffset
              Next VertexNum
            Next ObjectNum
            For X = 0 To MaxX
              Select Case XMod8
                Case 1
                  If ComputerPage(Y, X) = 0 Then
                    SingleRectangle(0).X = Rectangle(0, 0).X
                    SingleRectangle(0).Y = Rectangle(0, 0).Y
                    SingleRectangle(1).X = Rectangle(0, 1).X
                    SingleRectangle(1).Y = Rectangle(0, 1).Y
                    SingleRectangle(2).X = Rectangle(0, 2).X
                    SingleRectangle(2).Y = Rectangle(0, 2).Y
                    SingleRectangle(3).X = Rectangle(0, 3).X
                    SingleRectangle(3).Y = Rectangle(0, 3).Y
                    Call OutputRectangle(XMax, XOffset, YMax, SingleRectangle(), PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen, RectangleSENWColor)
                  End If
                Case 5
                  If ComputerPage(Y, X) = 0 Then
                    SingleRectangle(0).X = Rectangle(5, 0).X
                    SingleRectangle(0).Y = Rectangle(5, 0).Y
                    SingleRectangle(1).X = Rectangle(5, 1).X
                    SingleRectangle(1).Y = Rectangle(5, 1).Y
                    SingleRectangle(2).X = Rectangle(5, 2).X
                    SingleRectangle(2).Y = Rectangle(5, 2).Y
                    SingleRectangle(3).X = Rectangle(5, 3).X
                    SingleRectangle(3).Y = Rectangle(5, 3).Y
                    Call OutputRectangle(XMax, XOffset, YMax, SingleRectangle(), PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen, RectangleSWNEColor)
                  End If
                Case Else
              End Select
              XMod8 = XMod8 + 1
              If XMod8 >= 8 Then
                XMod8 = 0
                For ObjectNum = 0 To 5 Step 5
                  For VertexNum = 0 To 3
                    Rectangle(ObjectNum, VertexNum).X = Rectangle(ObjectNum, VertexNum).X + 3#
                  Next VertexNum
                Next ObjectNum
              End If
            Next X
        End Select
        YMod4 = YMod4 + 1
        If YMod4 >= 4 Then
          YMod4 = 0
          YOffset = YOffset + Sqrt3
        End If
        Y = Y + 1
      Else
        State = 5
      End If
      If UsePalette Then
        NumRealized = SelectPalette(frm3DMaze.hDC, OldPaletteHandle, 0)
      End If
      DoEvents
      If State < 5 Then
        Timer1.Enabled = True
      Else
        If State = 5 Then
          AlreadyPainting = False
          Call HexDisplayUserMoves(MaxX, MaxY, UserPage, XMax, XOffset, YMax, CosTilt, SinTilt, PixelsPerX, PixelsPerZ, RelDistOfUserFromScreen)
          If SolutionDisplayed Then
            Call HexDisplaySolution(MaxY, ComputerPage, XMax, XOffset, YMax, CosTilt, SinTilt, PixelsPerX, PixelsPerZ, RelDistOfUserFromScreen)
            Text1.Text = ""
          Else
            If UserHasSolved Then
              Text1.Text = "Congratulations!"
            Else
              Text1.Text = "Use Home, Up Arrow, PgUp, End, Down Arrow, and PgDn to solve."
            End If
          End If
          mnuActionItem(1).Enabled = True
          mnuActionItem(2).Enabled = True
        End If
      End If
    Case Else
      DoEvents
  End Select
End Sub

Private Sub SqrDisplaySolution(MaxY As Integer, Page() As Byte, XMax As Double, XOffset As Double, YMax As Double, CosTilt As Double, SinTilt As Double, PixelsPerX As Double, PixelsPerZ As Double, RelDistOfUserFromScreen As Double)
  Dim DeltaIndex As Byte
  Dim OldPaletteHandle As Long
  Dim PathFound As Integer
  Dim TemDouble As Double
  Dim X As Integer
  Dim XNext As Integer
  Dim XPrevious As Integer
  Dim XRelative As Double
  Dim XRelativeNext As Double
  Dim Y As Integer
  Dim YNext As Integer
  Dim YPrevious As Integer
  Dim YRelative As Double
  Dim YRelativeNext As Double
  If UsePalette Then
    OldPaletteHandle = SelectPalette(frm3DMaze.hDC, PaletteHandle, 0)
    NumRealized = RealizePalette(frm3DMaze.hDC)
  End If
  XRelative = (RelativeWidthOfWall + 1#) / 2#
  YRelative = (RelativeWidthOfWall + 1#) / 2#
  CurrentColor = SolutionColor
  Call DrawLine(XRelative, RelativeWidthOfWall / 2#, XRelative, YRelative, XMax, XOffset, YMax, CosTilt, SinTilt, PixelsPerX, PixelsPerZ, RelDistOfUserFromScreen)
  XPrevious = 1
  YPrevious = -1
  X = 1
  Y = 1
  Do
    PathFound = False
    DeltaIndex = 0
    Do While (Not PathFound)
      XNext = X + SqrDeltaX(DeltaIndex, 0)
      YNext = Y + SqrDeltaY(DeltaIndex, 0)
      If Page(YNext, XNext) = 1 Then
        XNext = XNext + SqrDeltaX(DeltaIndex, 0)
        YNext = YNext + SqrDeltaY(DeltaIndex, 0)
        If (XNext <> XPrevious) Or (YNext <> YPrevious) Then
          PathFound = True
        Else
          DeltaIndex = DeltaIndex + 1
        End If
      Else
        DeltaIndex = DeltaIndex + 1
      End If
    Loop
    If YNext < MaxY Then
      TemDouble = SqrDeltaX(DeltaIndex, 0)
      XRelativeNext = XRelative + TemDouble
      TemDouble = SqrDeltaY(DeltaIndex, 0)
      YRelativeNext = YRelative + TemDouble
      Call DrawLine(XRelative, YRelative, XRelativeNext, YRelativeNext, XMax, XOffset, YMax, CosTilt, SinTilt, PixelsPerX, PixelsPerZ, RelDistOfUserFromScreen)
    Else
      Call DrawLine(XRelative, YRelative, XRelative, YMax, XMax, XOffset, YMax, CosTilt, SinTilt, PixelsPerX, PixelsPerZ, RelDistOfUserFromScreen)
    End If
    XPrevious = X
    YPrevious = Y
    X = XNext
    Y = YNext
    XRelative = XRelativeNext
    YRelative = YRelativeNext
  Loop While YNext < MaxY
  If UsePalette Then
    NumRealized = SelectPalette(frm3DMaze.hDC, OldPaletteHandle, 0)
  End If
End Sub

Private Sub SqrDisplayUserMoves(MaxX As Integer, MaxY As Integer, Page() As Byte, XMax As Double, XOffset As Double, YMax As Double, CosTilt As Double, SinTilt As Double, PixelsPerX As Double, PixelsPerZ As Double, RelDistOfUserFromScreen As Double)
  Dim DeltaIndex As Byte
  Dim OldPaletteHandle As Long
  Dim TemDouble As Double
  Dim X As Integer
  Dim XNext As Integer
  Dim XNextNext As Integer
  Dim XRelative As Double
  Dim XRelativeNext As Double
  Dim Y As Integer
  Dim YNext As Integer
  Dim YNextNext As Integer
  Dim YRelative As Double
  Dim YRelativeNext As Double
  If UsePalette Then
    OldPaletteHandle = SelectPalette(frm3DMaze.hDC, PaletteHandle, 0)
    NumRealized = RealizePalette(frm3DMaze.hDC)
  End If
  Y = 1
  YRelative = (RelativeWidthOfWall + 1#) / 2#
  Do While (Y < MaxY)
    X = 1
    XRelative = (RelativeWidthOfWall + 1#) / 2#
    Do While (X < MaxX)
      If ((Page(Y, X) = 1) Or (Page(Y, X) = 3)) Then
        For DeltaIndex = 0 To 3
          XNext = X + SqrDeltaX(DeltaIndex, 0)
          YNext = Y + SqrDeltaY(DeltaIndex, 0)
          If Page(YNext, XNext) <> 0 Then
            If YNext = 0 Then
              CurrentColor = AdvanceColor
              Call DrawLine(XRelative, RelativeWidthOfWall / 2#, XRelative, YRelative, XMax, XOffset, YMax, CosTilt, SinTilt, PixelsPerX, PixelsPerZ, RelDistOfUserFromScreen)
            Else
              If YNext = MaxY Then
                If UserHasSolved Then
                  CurrentColor = AdvanceColor
                  Call DrawLine(XRelative, YRelative, XRelative, YMax, XMax, XOffset, YMax, CosTilt, SinTilt, PixelsPerX, PixelsPerZ, RelDistOfUserFromScreen)
                End If
              Else
                XNextNext = XNext + SqrDeltaX(DeltaIndex, 0)
                If XNextNext > 0 Then
                  If XNextNext < MaxX Then
                    YNextNext = YNext + SqrDeltaY(DeltaIndex, 0)
                    If YNextNext > 0 Then
                      If YNextNext < MaxY Then
                        If ((Page(YNextNext, XNextNext) = 1) Or (Page(YNextNext, XNextNext) = 3)) Then
                          If Page(Y, X) = Page(YNextNext, XNextNext) Then
                            If Page(Y, X) = 1 Then
                              CurrentColor = AdvanceColor
                            Else
                              CurrentColor = BackoutColor
                            End If
                          Else
                            CurrentColor = BackoutColor
                          End If
                          TemDouble = SqrDeltaX(DeltaIndex, 0)
                          XRelativeNext = XRelative + TemDouble / 2#
                          TemDouble = SqrDeltaY(DeltaIndex, 0)
                          YRelativeNext = YRelative + TemDouble / 2#
                          Call DrawLine(XRelative, YRelative, XRelativeNext, YRelativeNext, XMax, XOffset, YMax, CosTilt, SinTilt, PixelsPerX, PixelsPerZ, RelDistOfUserFromScreen)
                        End If
                       End If
                    End If
                  End If
                End If
              End If
            End If
          End If
        Next DeltaIndex
      End If
      XRelative = XRelative + 1#
      X = X + 2
    Loop
    YRelative = YRelative + 1#
    Y = Y + 2
  Loop
  If UsePalette Then
    NumRealized = SelectPalette(frm3DMaze.hDC, OldPaletteHandle, 0)
  End If
End Sub

Private Sub SqrSolveMaze(Stack() As StackRec, Page() As Byte, NumRoomsInSolution As Integer, Adjacency As Integer, MaxX As Integer, MaxY As Integer)
  Dim DeltaIndex As Byte
  Dim PassageFound As Integer
  Dim StackHead As Integer
  Dim X As Integer
  Dim XNext As Integer
  Dim Y As Integer
  Dim YNext As Integer

  NumRoomsInSolution = 1
  Adjacency = 0
  X = 1
  Y = 1
  StackHead = -1
  Page(Y, X) = 1
  Do
    DeltaIndex = 0
    PassageFound = False
    Do
      Do While ((DeltaIndex < 4) And (Not PassageFound))
        XNext = X + SqrDeltaX(DeltaIndex, 0)
        YNext = Y + SqrDeltaY(DeltaIndex, 0)
        If Page(YNext, XNext) = 2 Then
          PassageFound = True
        Else
          DeltaIndex = DeltaIndex + 1
        End If
      Loop
      If Not PassageFound Then
        DeltaIndex = Stack(StackHead).Index1
        Page(Y, X) = 2
        X = X - SqrDeltaX(DeltaIndex, 0)
        Y = Y - SqrDeltaY(DeltaIndex, 0)
        Page(Y, X) = 2
        X = X - SqrDeltaX(DeltaIndex, 0)
        Y = Y - SqrDeltaY(DeltaIndex, 0)
        StackHead = StackHead - 1
        DeltaIndex = DeltaIndex + 1
      End If
    Loop While Not PassageFound
    Page(YNext, XNext) = 1
    XNext = XNext + SqrDeltaX(DeltaIndex, 0)
    YNext = YNext + SqrDeltaY(DeltaIndex, 0)
    If YNext <= MaxY Then
      StackHead = StackHead + 1
      Stack(StackHead).Index1 = DeltaIndex
      Page(YNext, XNext) = 1
      X = XNext
      Y = YNext
    End If
  Loop While YNext < MaxY
  X = MaxX - 1
  Y = MaxY - 1
  Adjacency = 0
  Do While (StackHead >= 0)
    For DeltaIndex = 0 To 3
      XNext = X + SqrDeltaX(DeltaIndex, 0)
      YNext = Y + SqrDeltaY(DeltaIndex, 0)
      If Page(YNext, XNext) <> 1 Then
        If Page(YNext, XNext) = 0 Then
          XNext = XNext + SqrDeltaX(DeltaIndex, 0)
          YNext = YNext + SqrDeltaY(DeltaIndex, 0)
          If XNext < 0 Then
            Adjacency = Adjacency + 1
          Else
            If XNext > MaxX Then
              Adjacency = Adjacency + 1
            Else
              If YNext < 0 Then
                Adjacency = Adjacency + 1
              Else
                If YNext > MaxY Then
                  Adjacency = Adjacency + 1
                Else
                  If Page(YNext, XNext) = 1 Then
                    Adjacency = Adjacency + 1
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
    Next DeltaIndex
    X = X - 2 * SqrDeltaX(Stack(StackHead).Index1, 0)
    Y = Y - 2 * SqrDeltaY(Stack(StackHead).Index1, 0)
    StackHead = StackHead - 1
    NumRoomsInSolution = NumRoomsInSolution + 1
  Loop
  For DeltaIndex = 0 To 3
    XNext = X + SqrDeltaX(DeltaIndex, 0)
    YNext = X + SqrDeltaY(DeltaIndex, 0)
    If Page(YNext, XNext) <> 2 Then
      If Page(YNext, XNext) = 0 Then
        XNext = XNext + SqrDeltaX(DeltaIndex, 0)
        YNext = YNext + SqrDeltaY(DeltaIndex, 0)
        If XNext < 0 Then
          Adjacency = Adjacency + 1
        Else
          If XNext > MaxX Then
            Adjacency = Adjacency + 1
          Else
            If YNext < 0 Then
              Adjacency = Adjacency + 1
            Else
              If YNext > MaxY Then
                Adjacency = Adjacency + 1
              Else
                If Page(YNext, XNext) = 1 Then
                  Adjacency = Adjacency + 1
                End If
              End If
            End If
          End If
        End If
      End If
    End If
  Next DeltaIndex
End Sub

Private Sub SqrGenerateMaze(Page() As Byte, MaxX As Integer, MaxY As Integer, Stack() As StackRec, NumColumns As Integer, NumRows As Integer, Seed() As Byte)
  Dim DeltaIndex1 As Byte
  Dim DeltaIndex2 As Integer
  Dim Digit As Integer
  Dim DigitNum As Byte
  Dim PassageFound As Integer
  Dim RN(7) As Integer
  Dim RNIndex1 As Integer
  Dim RNIndex2 As Integer
  Dim SearchComplete As Integer
  Dim StackHead As Integer
  Dim Sum As Integer
  Dim TemInt As Integer
  Dim X As Integer
  Dim XNext As Integer
  Dim Y As Integer
  Dim YNext As Integer

  RN(0) = Seed(0) + 1
  RN(1) = Seed(1) + 1
  RN(2) = Seed(2) + 1
  RN(3) = Seed(3) + 1
  RN(4) = Seed(4) + 1
  RN(5) = Seed(5) + 1
  RN(6) = Seed(6) + 1
  RN(7) = Seed(7) + 1
  For Y = 0 To MaxY
    For X = 0 To MaxX
      Page(Y, X) = 0
    Next X
  Next Y
  Sum = 0
  For DigitNum = 1 To 3
    Digit = RN(0)
    RNIndex1 = 0
    RNIndex2 = 1
    Do While (RNIndex2 < 8)
      TemInt = RN(RNIndex2)
      RN(RNIndex1) = TemInt
      Digit = Digit + TemInt
      If Digit >= 29 Then Digit = Digit - 29
      RNIndex1 = RNIndex2
      RNIndex2 = RNIndex2 + 1
    Loop
    RN(7) = Digit
    Sum = 29 * Sum + Digit
  Next DigitNum
  X = 2 * (Sum Mod NumColumns) + 1
  Sum = 0
  For DigitNum = 1 To 3
    Digit = RN(0)
    RNIndex1 = 0
    RNIndex2 = 1
    Do While (RNIndex2 < 8)
      TemInt = RN(RNIndex2)
      RN(RNIndex1) = TemInt
      Digit = Digit + TemInt
      If Digit >= 29 Then Digit = Digit - 29
      RNIndex1 = RNIndex2
      RNIndex2 = RNIndex2 + 1
    Loop
    RN(7) = Digit
    Sum = 29 * Sum + Digit
  Next DigitNum
  Y = 2 * (Sum Mod NumRows) + 1
  Page(Y, X) = 2
  StackHead = -1
  Do
    DeltaIndex1 = 0
    Do
      DeltaIndex2 = RN(0)
      RNIndex1 = 0
      RNIndex2 = 1
      Do While (RNIndex2 < 8)
        TemInt = RN(RNIndex2)
        RN(RNIndex1) = TemInt
        DeltaIndex2 = DeltaIndex2 + TemInt
        If DeltaIndex2 >= 29 Then DeltaIndex2 = DeltaIndex2 - 29
        RNIndex1 = RNIndex2
        RNIndex2 = RNIndex2 + 1
      Loop
      RN(7) = DeltaIndex2
    Loop While DeltaIndex2 >= 24
    PassageFound = False
    SearchComplete = False
    Do While (Not SearchComplete)
      Do While ((DeltaIndex1 < 4) And (Not PassageFound))
        XNext = X + 2 * SqrDeltaX(DeltaIndex1, DeltaIndex2)
        If XNext <= 0 Then
          DeltaIndex1 = DeltaIndex1 + 1
        Else
          If XNext > MaxX Then
            DeltaIndex1 = DeltaIndex1 + 1
          Else
            YNext = Y + 2 * SqrDeltaY(DeltaIndex1, DeltaIndex2)
            If YNext <= 0 Then
              DeltaIndex1 = DeltaIndex1 + 1
            Else
              If YNext > MaxY Then
                DeltaIndex1 = DeltaIndex1 + 1
              Else
                If Page(YNext, XNext) = 0 Then
                  PassageFound = True
                Else
                  DeltaIndex1 = DeltaIndex1 + 1
                End If
              End If
            End If
          End If
        End If
      Loop
      If Not PassageFound Then
        If StackHead >= 0 Then
          DeltaIndex1 = Stack(StackHead).Index1
          DeltaIndex2 = Stack(StackHead).Index2
          X = X - 2 * SqrDeltaX(DeltaIndex1, DeltaIndex2)
          Y = Y - 2 * SqrDeltaY(DeltaIndex1, DeltaIndex2)
          StackHead = StackHead - 1
          DeltaIndex1 = DeltaIndex1 + 1
        End If
      End If
      If ((PassageFound) Or ((StackHead = -1) And (DeltaIndex1 >= 4))) Then
        SearchComplete = True
      Else
        SearchComplete = False
      End If
    Loop
    If PassageFound Then
      StackHead = StackHead + 1
      Stack(StackHead).Index1 = DeltaIndex1
      Stack(StackHead).Index2 = DeltaIndex2
      Page(YNext, XNext) = 2
      Page((Y + YNext) \ 2, (X + XNext) \ 2) = 2
      X = XNext
      Y = YNext
    End If
  Loop While StackHead <> -1
  Page(0, 1) = 1
  Page(MaxY, MaxX - 1) = 2
End Sub

Private Sub SqrSelectMaze(Seed As String, Page() As Byte, MaxX As Integer, MaxY As Integer, Stack() As StackRec, NumRoomsInMaze As Integer, NumColumns As Integer, NumRows As Integer, SecondsForMazeSelection As Double)
  Dim Adjacency As Integer
  Dim Counter0 As Byte
  Dim Counter1 As Byte
  Dim Counter2 As Byte
  Dim Counter3 As Byte
  Dim Counter4 As Byte
  Dim Counter5 As Byte
  Dim Counter6 As Byte
  Dim Counter7 As Byte
  Dim ElapsedTime As Double
  Dim MinAdjacency As Integer
  Dim NumRoomsInSolution As Integer
  Dim NumRoomsInSolutionAtMin As Integer
  Dim RN(7) As Integer
  Dim RNIndex1 As Integer
  Dim RNIndex2 As Integer
  Dim SeedByte(7) As Byte
  Dim SeedByteAtMin(7) As Byte
  Dim SeedLength As Integer
  Dim StartTime As Double

  SeedLength = Len(Seed)
  If SeedLength > 8 Then SeedLength = 8
  RNIndex1 = 0
  For RNIndex2 = 1 To SeedLength
    RN(RNIndex1) = Asc(Mid$(Seed, RNIndex2, 1)) Mod 10
    RNIndex1 = RNIndex1 + 1
  Next RNIndex2
  RNIndex2 = 7
  Do While (RNIndex1 > 0)
    RNIndex1 = RNIndex1 - 1
    RN(RNIndex2) = RN(RNIndex1)
    RNIndex2 = RNIndex2 - 1
  Loop
  Do While (RNIndex2 >= 0)
    RN(RNIndex2) = 8
    RNIndex2 = RNIndex2 - 1
  Loop
  Counter0 = RN(0)
  Counter1 = RN(1)
  Counter2 = RN(2)
  Counter3 = RN(3)
  Counter4 = RN(4)
  Counter5 = RN(5)
  Counter6 = RN(6)
  Counter7 = RN(7)
  Call Hash(Counter0, Counter1, Counter2, Counter3, Counter4, Counter5, Counter6, Counter7)
  MinAdjacency = 2 * NumRoomsInMaze + 1
  NumRoomsInSolutionAtMin = 0
  SeedByteAtMin(0) = Counter0
  SeedByteAtMin(1) = Counter1
  SeedByteAtMin(2) = Counter2
  SeedByteAtMin(3) = Counter3
  SeedByteAtMin(4) = Counter4
  SeedByteAtMin(5) = Counter5
  SeedByteAtMin(6) = Counter6
  SeedByteAtMin(7) = Counter7
  StartTime = Timer
  Do
    SeedByte(0) = Counter0
    SeedByte(1) = Counter1
    SeedByte(2) = Counter2
    SeedByte(3) = Counter3
    SeedByte(4) = Counter4
    SeedByte(5) = Counter5
    SeedByte(6) = Counter6
    SeedByte(7) = Counter7
    Call SqrGenerateMaze(Page(), MaxX, MaxY, Stack(), NumColumns, NumRows, SeedByte())
    Call SqrSolveMaze(Stack(), Page(), NumRoomsInSolution, Adjacency, MaxX, MaxY)
    If 3 * NumRoomsInSolution >= NumRoomsInMaze Then
      If Adjacency < MinAdjacency Then
        MinAdjacency = Adjacency
        NumRoomsInSolutionAtMin = NumRoomsInSolution
        SeedByteAtMin(0) = SeedByte(0)
        SeedByteAtMin(1) = SeedByte(1)
        SeedByteAtMin(2) = SeedByte(2)
        SeedByteAtMin(3) = SeedByte(3)
        SeedByteAtMin(4) = SeedByte(4)
        SeedByteAtMin(5) = SeedByte(5)
        SeedByteAtMin(6) = SeedByte(6)
        SeedByteAtMin(7) = SeedByte(7)
      Else
        If Adjacency = MinAdjacency Then
          If NumRoomsInSolution > NumRoomsInSolutionAtMin Then
            NumRoomsInSolutionAtMin = NumRoomsInSolution
            SeedByteAtMin(0) = SeedByte(0)
            SeedByteAtMin(1) = SeedByte(1)
            SeedByteAtMin(2) = SeedByte(2)
            SeedByteAtMin(3) = SeedByte(3)
            SeedByteAtMin(4) = SeedByte(4)
            SeedByteAtMin(5) = SeedByte(5)
            SeedByteAtMin(6) = SeedByte(6)
            SeedByteAtMin(7) = SeedByte(7)
          End If
        End If
      End If
    End If
    Call Increment(Counter0, Counter1, Counter2, Counter3, Counter4, Counter5, Counter6, Counter7)
    ElapsedTime = Timer - StartTime
  Loop While ((ElapsedTime >= 0#) And (ElapsedTime < SecondsForMazeSelection))
  Call SqrGenerateMaze(Page(), MaxX, MaxY, Stack(), NumColumns, NumRows, SeedByteAtMin())
  Call SqrSolveMaze(Stack(), Page(), NumRoomsInSolution, Adjacency, MaxX, MaxY)
End Sub

Private Sub SqrOutputMaze()
  Dim ObjectNum As Byte
  Dim Radians As Double
  Dim RadiansPerDegree As Double
  Dim SingleRectangle(3) As VertexRec
  Dim SingleTriangle(2) As VertexRec
  Dim TemDouble1 As Double
  Dim TemDouble2 As Double
  Dim TemDouble3 As Double
  Dim TemDouble4 As Double
  Dim Triangle(3, 2) As VertexRec
  Dim VertexNum As Byte
  Dim XMod8 As Byte
  Dim X0 As Double
  Dim X1 As Double
  Dim X2 As Double
  Dim X3 As Double
  Dim Y0 As Double
  Dim Y1 As Double
  Dim Y2 As Double
  Dim Y3 As Double

  Select Case State
    Case 0
      Text1.Text = ""
      ScaleMode = 1
      If (Resize) Then
        TemDouble1 = ScaleWidth - VScroll1.Width
        TemDouble2 = MinWallLengthInInches
        TemDouble2 = 1440# * TemDouble2
        TemDouble3 = RelativeWidthOfWall
        NumColumns = Int(TemDouble1 / TemDouble2 - TemDouble3)
        If NumColumns < 2 Then NumColumns = 2
        TemDouble1 = ScaleHeight - Text1.Height
        TemDouble2 = ScaleWidth - VScroll1.Width
        ScaleMode = 3
        TemDouble3 = NumColumns
        NumRows = Int((TemDouble1 * TemDouble3) / TemDouble2)
        If NumRows < 2 Then NumRows = 2
        Tilt = 90 - VScroll1.Value
        MaxX = 2 * NumColumns
        MaxY = 2 * NumRows
        NumRoomsInMaze = NumRows * NumColumns
        ReDim ComputerPage(MaxY, MaxX)
        ReDim UserPage(MaxY, MaxX)
        ReDim Stack(NumRoomsInMaze)
        Call SqrSelectMaze(Seed, ComputerPage(), MaxX, MaxY, Stack(), NumRoomsInMaze, NumColumns, NumRows, SecondsForMazeSelection)
        For UserX = 0 To MaxX
          For UserY = 0 To MaxY
            If ComputerPage(UserY, UserX) = 0 Then
              UserPage(UserY, UserX) = 0
            Else
              UserPage(UserY, UserX) = 2
            End If
          Next UserY
        Next UserX
        UserX = 1
        UserXRelative = (RelativeWidthOfWall + 1#) / 2#
        UserY = 1
        UserYRelative = (RelativeWidthOfWall + 1#) / 2#
        UserPage(UserY, UserX) = 1
        Resize = False
      End If
      If (Paint) Then
        ScaleMode = 3
        Cls
        RadiansPerDegree = Atn(1#) / 45#
        Radians = Tilt * RadiansPerDegree
        SinTilt = Sin(Radians)
        CosTilt = Cos(Radians)
        TemDouble1 = NumColumns
        XMax = TemDouble1 + RelativeWidthOfWall
        TemDouble1 = ScaleWidth - VScroll1.Width
        PixelsPerX = (TemDouble1 - 1#) / (XMax * (XMax / (XMax - RelativeHeightOfWall)))
        XOffset = (XMax / 2#) * (RelativeHeightOfWall / (XMax - RelativeHeightOfWall))
        TemDouble1 = NumRows
        YMax = TemDouble1 + RelativeWidthOfWall
        TemDouble1 = ScaleHeight - Text1.Height
        PixelsPerZ = (TemDouble1 - 1#) / Sqr(YMax * YMax + RelativeHeightOfWall * RelativeHeightOfWall)
        If YMax > XMax Then
          RelDistOfUserFromScreen = YMax
        Else
          RelDistOfUserFromScreen = XMax
        End If
        Paint = False
      End If
      If State = 0 Then
        State = 1
        DoEvents
        If State < 5 Then
          Timer1.Enabled = True
        End If
      End If
    Case 1
      BaseRectangle(0, 0).X = 0#
      BaseRectangle(0, 0).Y = 0#
      BaseRectangle(0, 1).X = RelativeWidthOfWall
      BaseRectangle(0, 1).Y = 0#
      BaseRectangle(0, 2).X = RelativeWidthOfWall
      BaseRectangle(0, 2).Y = RelativeWidthOfWall
      BaseRectangle(0, 3).X = 0#
      BaseRectangle(0, 3).Y = RelativeWidthOfWall
      BaseRectangle(1, 0).X = RelativeWidthOfWall
      BaseRectangle(1, 0).Y = 0#
      BaseRectangle(1, 1).X = 1#
      BaseRectangle(1, 1).Y = 0#
      BaseRectangle(1, 2).X = 1#
      BaseRectangle(1, 2).Y = RelativeWidthOfWall
      BaseRectangle(1, 3).X = RelativeWidthOfWall
      BaseRectangle(1, 3).Y = RelativeWidthOfWall
      BaseRectangle(2, 0).X = RelativeWidthOfWall
      BaseRectangle(2, 0).Y = RelativeWidthOfWall
      BaseRectangle(2, 1).X = 1#
      BaseRectangle(2, 1).Y = RelativeWidthOfWall
      BaseRectangle(2, 2).X = 1#
      BaseRectangle(2, 2).Y = 1#
      BaseRectangle(2, 3).X = RelativeWidthOfWall
      BaseRectangle(2, 3).Y = 1#
      BaseRectangle(3, 0).X = 0#
      BaseRectangle(3, 0).Y = RelativeWidthOfWall
      BaseRectangle(3, 1).X = RelativeWidthOfWall
      BaseRectangle(3, 1).Y = RelativeWidthOfWall
      BaseRectangle(3, 2).X = RelativeWidthOfWall
      BaseRectangle(3, 2).Y = 1#
      BaseRectangle(3, 3).X = 0#
      BaseRectangle(3, 3).Y = 1#
      Rectangle(0, 0).X = 0#
      Rectangle(0, 0).Y = 0#
      Rectangle(0, 1).X = XMax
      Rectangle(0, 1).Y = 0#
      Rectangle(0, 2).X = XMax
      Rectangle(0, 2).Y = YMax
      Rectangle(0, 3).X = 0#
      Rectangle(0, 3).Y = YMax
      If UsePalette Then
        OldPaletteHandle = SelectPalette(frm3DMaze.hDC, PaletteHandle, 0)
        NumRealized = RealizePalette(frm3DMaze.hDC)
      End If
      X0 = Rectangle(0, 0).X
      Y0 = Rectangle(0, 0).Y
      X1 = Rectangle(0, 1).X
      Y1 = Rectangle(0, 1).Y
      X2 = Rectangle(0, 2).X
      Y2 = Rectangle(0, 2).Y
      X3 = Rectangle(0, 3).X
      Y3 = Rectangle(0, 3).Y
      Call DisplayQuadrilateral(XMax, XOffset, YMax, X0, Y0, 0#, X1, Y1, 0#, X2, Y2, 0#, X3, Y3, 0#, PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen, FloorColor)
      Y = 0
      YOffset = 0
      State = 4
      If UsePalette Then
        NumRealized = SelectPalette(frm3DMaze.hDC, OldPaletteHandle, 0)
      End If
      DoEvents
      If State < 5 Then
        Timer1.Enabled = True
      End If
    Case 4
      If UsePalette Then
        OldPaletteHandle = SelectPalette(frm3DMaze.hDC, PaletteHandle, 0)
        NumRealized = RealizePalette(frm3DMaze.hDC)
      End If
      If Y <= MaxY Then
        For VertexNum = 0 To 3
          Rectangle(0, VertexNum).X = BaseRectangle(0, VertexNum).X
          Rectangle(0, VertexNum).Y = BaseRectangle(0, VertexNum).Y + YOffset
        Next VertexNum
        X = 0
        Do While X <= MaxX
          If ComputerPage(Y, X) = 0 Then
            SingleRectangle(0).X = Rectangle(0, 0).X
            SingleRectangle(0).Y = Rectangle(0, 0).Y
            SingleRectangle(1).X = Rectangle(0, 1).X
            SingleRectangle(1).Y = Rectangle(0, 1).Y
            SingleRectangle(2).X = Rectangle(0, 2).X
            SingleRectangle(2).Y = Rectangle(0, 2).Y
            SingleRectangle(3).X = Rectangle(0, 3).X
            SingleRectangle(3).Y = Rectangle(0, 3).Y
            Call OutputLeftRight(XMax, XOffset, YMax, SingleRectangle(), PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen)
          End If
          For VertexNum = 0 To 3
            Rectangle(0, VertexNum).X = Rectangle(0, VertexNum).X + 1
          Next VertexNum
          X = X + 2
        Loop
        For VertexNum = 0 To 3
          Rectangle(0, VertexNum).X = BaseRectangle(0, VertexNum).X
          Rectangle(0, VertexNum).Y = BaseRectangle(0, VertexNum).Y + YOffset
        Next VertexNum
        For VertexNum = 0 To 3
          Rectangle(1, VertexNum).X = BaseRectangle(1, VertexNum).X
          Rectangle(1, VertexNum).Y = BaseRectangle(1, VertexNum).Y + YOffset
        Next VertexNum
        X = 0
        Do While X <= MaxX
          If ComputerPage(Y, X) = 0 Then
            SingleRectangle(0).X = Rectangle(0, 0).X
            SingleRectangle(0).Y = Rectangle(0, 0).Y
            SingleRectangle(1).X = Rectangle(0, 1).X
            SingleRectangle(1).Y = Rectangle(0, 1).Y
            SingleRectangle(2).X = Rectangle(0, 2).X
            SingleRectangle(2).Y = Rectangle(0, 2).Y
            SingleRectangle(3).X = Rectangle(0, 3).X
            SingleRectangle(3).Y = Rectangle(0, 3).Y
            Call OutputRectangle(XMax, XOffset, YMax, SingleRectangle(), PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen, RectangleWEColor)
          End If
          For VertexNum = 0 To 3
            Rectangle(0, VertexNum).X = Rectangle(0, VertexNum).X + 1
          Next VertexNum
          X = X + 1
          If X <= MaxX Then
            If ComputerPage(Y, X) = 0 Then
              SingleRectangle(0).X = Rectangle(1, 0).X
              SingleRectangle(0).Y = Rectangle(1, 0).Y
              SingleRectangle(1).X = Rectangle(1, 1).X
              SingleRectangle(1).Y = Rectangle(1, 1).Y
              SingleRectangle(2).X = Rectangle(1, 2).X
              SingleRectangle(2).Y = Rectangle(1, 2).Y
              SingleRectangle(3).X = Rectangle(1, 3).X
              SingleRectangle(3).Y = Rectangle(1, 3).Y
              Call OutputRectangle(XMax, XOffset, YMax, SingleRectangle(), PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen, RectangleWEColor)
            End If
            For VertexNum = 0 To 3
              Rectangle(1, VertexNum).X = Rectangle(1, VertexNum).X + 1
            Next VertexNum
            X = X + 1
          End If
        Loop
        Y = Y + 1
        If Y <= MaxY Then
          For VertexNum = 0 To 3
            Rectangle(3, VertexNum).X = BaseRectangle(3, VertexNum).X
            Rectangle(3, VertexNum).Y = BaseRectangle(3, VertexNum).Y + YOffset
          Next VertexNum
          X = 0
          Do While X <= MaxX
            If ComputerPage(Y, X) = 0 Then
              SingleRectangle(0).X = Rectangle(3, 0).X
              SingleRectangle(0).Y = Rectangle(3, 0).Y
              SingleRectangle(1).X = Rectangle(3, 1).X
              SingleRectangle(1).Y = Rectangle(3, 1).Y
              SingleRectangle(2).X = Rectangle(3, 2).X
              SingleRectangle(2).Y = Rectangle(3, 2).Y
              SingleRectangle(3).X = Rectangle(3, 3).X
              SingleRectangle(3).Y = Rectangle(3, 3).Y
              Call OutputLeftRight(XMax, XOffset, YMax, SingleRectangle(), PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen)
            End If
            For VertexNum = 0 To 3
              Rectangle(3, VertexNum).X = Rectangle(3, VertexNum).X + 1
            Next VertexNum
            X = X + 2
          Loop
          For VertexNum = 0 To 3
            Rectangle(3, VertexNum).X = BaseRectangle(3, VertexNum).X
            Rectangle(3, VertexNum).Y = BaseRectangle(3, VertexNum).Y + YOffset
          Next VertexNum
          X = 0
          Do While X <= MaxX
            If ComputerPage(Y, X) = 0 Then
              SingleRectangle(0).X = Rectangle(3, 0).X
              SingleRectangle(0).Y = Rectangle(3, 0).Y
              SingleRectangle(1).X = Rectangle(3, 1).X
              SingleRectangle(1).Y = Rectangle(3, 1).Y
              SingleRectangle(2).X = Rectangle(3, 2).X
              SingleRectangle(2).Y = Rectangle(3, 2).Y
              SingleRectangle(3).X = Rectangle(3, 3).X
              SingleRectangle(3).Y = Rectangle(3, 3).Y
              Call OutputRectangle(XMax, XOffset, YMax, SingleRectangle(), PixelsPerX, PixelsPerZ, CosTilt, SinTilt, RelDistOfUserFromScreen, RectangleWEColor)
            End If
            For VertexNum = 0 To 3
              Rectangle(3, VertexNum).X = Rectangle(3, VertexNum).X + 1
            Next VertexNum
            X = X + 2
          Loop
          Y = Y + 1
        End If
        YOffset = YOffset + 1
      Else
        State = 5
      End If
      If UsePalette Then
        NumRealized = SelectPalette(frm3DMaze.hDC, OldPaletteHandle, 0)
      End If
      DoEvents
      If State < 5 Then
        Timer1.Enabled = True
      Else
        If State = 5 Then
          AlreadyPainting = False
          Call SqrDisplayUserMoves(MaxX, MaxY, UserPage, XMax, XOffset, YMax, CosTilt, SinTilt, PixelsPerX, PixelsPerZ, RelDistOfUserFromScreen)
          If SolutionDisplayed Then
            Call SqrDisplaySolution(MaxY, ComputerPage, XMax, XOffset, YMax, CosTilt, SinTilt, PixelsPerX, PixelsPerZ, RelDistOfUserFromScreen)
            Text1.Text = ""
          Else
            If UserHasSolved Then
              Text1.Text = "Congratulations!"
            Else
              Text1.Text = "Use the arrow keys to solve."
            End If
          End If
          mnuActionItem(1).Enabled = True
          mnuActionItem(2).Enabled = True
        End If
      End If
    Case Else
      DoEvents
  End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If ((State = 5) And (Not SolutionDisplayed) And (Not UserHasSolved)) Then
    Dim DeltaIndex1 As Integer
    Dim OldPaletteHandle As Long
    Dim PassageFound As Integer
    Dim TemDouble As Double
    Dim XNext As Integer
    Dim XRelativeNext As Double
    Dim YNext As Integer
    Dim YRelativeNext As Double
    PassageFound = True
    DeltaIndex1 = -1
    If mnuStyleItem(0).Checked Then
      Select Case KeyCode
        Case vbKeyPageDown, vbKeyNumpad3
          DeltaIndex1 = 5
          KeyCode = 0
        Case vbKeyHome, vbKeyNumpad7
          DeltaIndex1 = 0
          KeyCode = 0
        Case vbKeyLeft, vbKeyNumpad4
          Beep
          KeyCode = 0
        Case vbKeyUp, vbKeyNumpad8
          DeltaIndex1 = 2
          KeyCode = 0
        Case vbKeyRight, vbKeyNumpad6
          Beep
          KeyCode = 0
        Case vbKeyDown, vbKeyNumpad2
          DeltaIndex1 = 3
          KeyCode = 0
        Case vbKeyPageUp, vbKeyNumpad9
          DeltaIndex1 = 4
          KeyCode = 0
        Case vbKeyEnd, vbKeyNumpad1
          DeltaIndex1 = 1
          KeyCode = 0
      End Select
      If DeltaIndex1 >= 0 Then
        XNext = UserX + HexDeltaX(DeltaIndex1, 0)
        If XNext <= 0 Then
          PassageFound = False
        Else
          If XNext >= MaxX Then
            PassageFound = False
          Else
            YNext = UserY + HexDeltaY(DeltaIndex1, 0)
            If YNext <= 0 Then
              PassageFound = False
            Else
              If YNext > MaxY Then
                PassageFound = False
              Else
                If UserPage(YNext, XNext) = 0 Then
                  PassageFound = False
                End If
              End If
            End If
          End If
        End If
        If PassageFound Then
          XNext = XNext + HexDeltaX(DeltaIndex1, 0)
          YNext = YNext + HexDeltaY(DeltaIndex1, 0)
          If YNext < MaxY Then
            If UserPage(YNext, XNext) = 1 Then
              CurrentColor = BackoutColor
              UserPage(UserY, UserX) = 3
            Else
              CurrentColor = AdvanceColor
              UserPage(YNext, XNext) = 1
            End If
            Select Case (YNext - UserY)
              Case -4
                XRelativeNext = UserXRelative
                YRelativeNext = UserYRelative - Sqrt3
              Case -2
                If XNext > UserX Then
                  XRelativeNext = UserXRelative + 3# / 2#
                  YRelativeNext = UserYRelative - Sqrt3 / 2#
                Else
                  XRelativeNext = UserXRelative - 3# / 2#
                  YRelativeNext = UserYRelative - Sqrt3 / 2#
                End If
              Case 2
                If XNext > UserX Then
                  XRelativeNext = UserXRelative + 3# / 2#
                  YRelativeNext = UserYRelative + Sqrt3 / 2#
                Else
                  XRelativeNext = UserXRelative - 3# / 2#
                  YRelativeNext = UserYRelative + Sqrt3 / 2#
                End If
              Case Else
                XRelativeNext = UserXRelative
                YRelativeNext = UserYRelative + Sqrt3
            End Select
            If UsePalette Then
              OldPaletteHandle = SelectPalette(frm3DMaze.hDC, PaletteHandle, 0)
              NumRealized = RealizePalette(frm3DMaze.hDC)
            End If
            Call DrawLine(UserXRelative, UserYRelative, XRelativeNext, YRelativeNext, XMax, XOffset, YMax, CosTilt, SinTilt, PixelsPerX, PixelsPerZ, RelDistOfUserFromScreen)
            If UsePalette Then
              NumRealized = SelectPalette(frm3DMaze.hDC, OldPaletteHandle, 0)
            End If
          Else
            CurrentColor = AdvanceColor
            If UsePalette Then
              OldPaletteHandle = SelectPalette(frm3DMaze.hDC, PaletteHandle, 0)
              NumRealized = RealizePalette(frm3DMaze.hDC)
            End If
            Call DrawLine(UserXRelative, UserYRelative, UserXRelative, YMax, XMax, XOffset, YMax, CosTilt, SinTilt, PixelsPerX, PixelsPerZ, RelDistOfUserFromScreen)
            If UsePalette Then
              NumRealized = SelectPalette(frm3DMaze.hDC, OldPaletteHandle, 0)
            End If
            UserHasSolved = True
            Text1.Text = "Congratulations!"
          End If
          UserX = XNext
          UserY = YNext
          UserXRelative = XRelativeNext
          UserYRelative = YRelativeNext
        Else
          Beep
        End If
      End If
    Else
      Select Case KeyCode
        Case vbKeyPageDown, vbKeyNumpad3
          Beep
          KeyCode = 0
        Case vbKeyHome, vbKeyNumpad7
          Beep
          KeyCode = 0
        Case vbKeyLeft, vbKeyNumpad4
          DeltaIndex1 = 0
          KeyCode = 0
        Case vbKeyUp, vbKeyNumpad8
          DeltaIndex1 = 3
          KeyCode = 0
        Case vbKeyRight, vbKeyNumpad6
          DeltaIndex1 = 2
          KeyCode = 0
        Case vbKeyDown, vbKeyNumpad2
          DeltaIndex1 = 1
          KeyCode = 0
        Case vbKeyPageUp, vbKeyNumpad9
          Beep
          KeyCode = 0
        Case vbKeyEnd, vbKeyNumpad1
          Beep
          KeyCode = 0
      End Select
      If DeltaIndex1 >= 0 Then
        XNext = UserX + SqrDeltaX(DeltaIndex1, 0)
        If XNext <= 0 Then
          PassageFound = False
        Else
          If XNext >= MaxX Then
            PassageFound = False
          Else
            YNext = UserY + SqrDeltaY(DeltaIndex1, 0)
            If YNext <= 0 Then
              PassageFound = False
            Else
              If YNext > MaxY Then
                PassageFound = False
              Else
                If UserPage(YNext, XNext) = 0 Then
                  PassageFound = False
                End If
              End If
            End If
          End If
        End If
        If PassageFound Then
          XNext = XNext + SqrDeltaX(DeltaIndex1, 0)
          YNext = YNext + SqrDeltaY(DeltaIndex1, 0)
          If YNext < MaxY Then
            If UserPage(YNext, XNext) = 1 Then
              CurrentColor = BackoutColor
              UserPage(UserY, UserX) = 3
            Else
              CurrentColor = AdvanceColor
              UserPage(YNext, XNext) = 1
            End If
            TemDouble = SqrDeltaX(DeltaIndex1, 0)
            XRelativeNext = UserXRelative + TemDouble
            TemDouble = SqrDeltaY(DeltaIndex1, 0)
            YRelativeNext = UserYRelative + TemDouble
            If UsePalette Then
              OldPaletteHandle = SelectPalette(frm3DMaze.hDC, PaletteHandle, 0)
              NumRealized = RealizePalette(frm3DMaze.hDC)
            End If
            Call DrawLine(UserXRelative, UserYRelative, XRelativeNext, YRelativeNext, XMax, XOffset, YMax, CosTilt, SinTilt, PixelsPerX, PixelsPerZ, RelDistOfUserFromScreen)
            If UsePalette Then
              NumRealized = SelectPalette(frm3DMaze.hDC, OldPaletteHandle, 0)
            End If
          Else
            CurrentColor = AdvanceColor
            If UsePalette Then
              OldPaletteHandle = SelectPalette(frm3DMaze.hDC, PaletteHandle, 0)
              NumRealized = RealizePalette(frm3DMaze.hDC)
            End If
            Call DrawLine(UserXRelative, UserYRelative, UserXRelative, YMax, XMax, XOffset, YMax, CosTilt, SinTilt, PixelsPerX, PixelsPerZ, RelDistOfUserFromScreen)
            If UsePalette Then
              NumRealized = SelectPalette(frm3DMaze.hDC, OldPaletteHandle, 0)
            End If
            UserHasSolved = True
            Text1.Text = "Congratulations!"
          End If
          UserX = XNext
          UserY = YNext
          UserXRelative = XRelativeNext
          UserYRelative = YRelativeNext
        Else
          Beep
        End If
      End If
    End If
  End If
End Sub

Private Sub Form_Load()
  Dim ColorNum As Integer
  Dim DeltaIndex1a As Byte
  Dim DeltaIndex1b As Byte
  Dim DeltaIndex1c As Byte
  Dim DeltaIndex1d As Byte
  Dim DeltaIndex1e As Byte
  Dim DeltaIndex1f As Byte
  Dim DeltaIndex2 As Integer
  Dim LogicalPalette As LOGPALETTE
  Dim NumBits As Long
  Dim NumColorsFree As Long
  Dim Tint As Integer
  OldPaletteHandle = 0
  AlreadyPainting = False
  SolutionDisplayed = False
  UserHasSolved = False
  State = 0
  Minimized = False
  mnuStyleItem(0).Checked = False
  mnuStyleItem(1).Checked = True
  SubstitutionHigh(0) = 4
  SubstitutionHigh(1) = 1
  SubstitutionHigh(2) = 2
  SubstitutionHigh(3) = 8
  SubstitutionHigh(4) = 8
  SubstitutionHigh(5) = 9
  SubstitutionHigh(6) = 9
  SubstitutionHigh(7) = 6
  SubstitutionHigh(8) = 5
  SubstitutionHigh(9) = 7
  SubstitutionHigh(10) = 2
  SubstitutionHigh(11) = 1
  SubstitutionHigh(12) = 2
  SubstitutionHigh(13) = 9
  SubstitutionHigh(14) = 8
  SubstitutionHigh(15) = 8
  SubstitutionHigh(16) = 6
  SubstitutionHigh(17) = 3
  SubstitutionHigh(18) = 5
  SubstitutionHigh(19) = 1
  SubstitutionHigh(20) = 9
  SubstitutionHigh(21) = 5
  SubstitutionHigh(22) = 4
  SubstitutionHigh(23) = 4
  SubstitutionHigh(24) = 9
  SubstitutionHigh(25) = 8
  SubstitutionHigh(26) = 6
  SubstitutionHigh(27) = 0
  SubstitutionHigh(28) = 8
  SubstitutionHigh(29) = 0
  SubstitutionHigh(30) = 6
  SubstitutionHigh(31) = 0
  SubstitutionHigh(32) = 2
  SubstitutionHigh(33) = 4
  SubstitutionHigh(34) = 1
  SubstitutionHigh(35) = 9
  SubstitutionHigh(36) = 2
  SubstitutionHigh(37) = 0
  SubstitutionHigh(38) = 7
  SubstitutionHigh(39) = 4
  SubstitutionHigh(40) = 7
  SubstitutionHigh(41) = 3
  SubstitutionHigh(42) = 0
  SubstitutionHigh(43) = 0
  SubstitutionHigh(44) = 2
  SubstitutionHigh(45) = 6
  SubstitutionHigh(46) = 8
  SubstitutionHigh(47) = 9
  SubstitutionHigh(48) = 4
  SubstitutionHigh(49) = 0
  SubstitutionHigh(50) = 8
  SubstitutionHigh(51) = 3
  SubstitutionHigh(52) = 2
  SubstitutionHigh(53) = 3
  SubstitutionHigh(54) = 2
  SubstitutionHigh(55) = 5
  SubstitutionHigh(56) = 2
  SubstitutionHigh(57) = 4
  SubstitutionHigh(58) = 6
  SubstitutionHigh(59) = 9
  SubstitutionHigh(60) = 7
  SubstitutionHigh(61) = 9
  SubstitutionHigh(62) = 1
  SubstitutionHigh(63) = 3
  SubstitutionHigh(64) = 5
  SubstitutionHigh(65) = 7
  SubstitutionHigh(66) = 1
  SubstitutionHigh(67) = 1
  SubstitutionHigh(68) = 4
  SubstitutionHigh(69) = 5
  SubstitutionHigh(70) = 8
  SubstitutionHigh(71) = 1
  SubstitutionHigh(72) = 6
  SubstitutionHigh(73) = 0
  SubstitutionHigh(74) = 5
  SubstitutionHigh(75) = 7
  SubstitutionHigh(76) = 8
  SubstitutionHigh(77) = 2
  SubstitutionHigh(78) = 3
  SubstitutionHigh(79) = 3
  SubstitutionHigh(80) = 7
  SubstitutionHigh(81) = 3
  SubstitutionHigh(82) = 5
  SubstitutionHigh(83) = 1
  SubstitutionHigh(84) = 7
  SubstitutionHigh(85) = 5
  SubstitutionHigh(86) = 4
  SubstitutionHigh(87) = 0
  SubstitutionHigh(88) = 3
  SubstitutionHigh(89) = 6
  SubstitutionHigh(90) = 3
  SubstitutionHigh(91) = 7
  SubstitutionHigh(92) = 7
  SubstitutionHigh(93) = 1
  SubstitutionHigh(94) = 9
  SubstitutionHigh(95) = 4
  SubstitutionHigh(96) = 0
  SubstitutionHigh(97) = 5
  SubstitutionHigh(98) = 6
  SubstitutionHigh(99) = 6
  SubstitutionLow(0) = 1
  SubstitutionLow(1) = 2
  SubstitutionLow(2) = 2
  SubstitutionLow(3) = 1
  SubstitutionLow(4) = 5
  SubstitutionLow(5) = 5
  SubstitutionLow(6) = 4
  SubstitutionLow(7) = 6
  SubstitutionLow(8) = 4
  SubstitutionLow(9) = 6
  SubstitutionLow(10) = 4
  SubstitutionLow(11) = 4
  SubstitutionLow(12) = 5
  SubstitutionLow(13) = 6
  SubstitutionLow(14) = 6
  SubstitutionLow(15) = 3
  SubstitutionLow(16) = 0
  SubstitutionLow(17) = 9
  SubstitutionLow(18) = 6
  SubstitutionLow(19) = 5
  SubstitutionLow(20) = 7
  SubstitutionLow(21) = 2
  SubstitutionLow(22) = 0
  SubstitutionLow(23) = 9
  SubstitutionLow(24) = 3
  SubstitutionLow(25) = 4
  SubstitutionLow(26) = 2
  SubstitutionLow(27) = 3
  SubstitutionLow(28) = 9
  SubstitutionLow(29) = 1
  SubstitutionLow(30) = 9
  SubstitutionLow(31) = 9
  SubstitutionLow(32) = 9
  SubstitutionLow(33) = 3
  SubstitutionLow(34) = 8
  SubstitutionLow(35) = 9
  SubstitutionLow(36) = 3
  SubstitutionLow(37) = 4
  SubstitutionLow(38) = 1
  SubstitutionLow(39) = 5
  SubstitutionLow(40) = 0
  SubstitutionLow(41) = 5
  SubstitutionLow(42) = 2
  SubstitutionLow(43) = 7
  SubstitutionLow(44) = 0
  SubstitutionLow(45) = 8
  SubstitutionLow(46) = 8
  SubstitutionLow(47) = 0
  SubstitutionLow(48) = 4
  SubstitutionLow(49) = 5
  SubstitutionLow(50) = 0
  SubstitutionLow(51) = 3
  SubstitutionLow(52) = 6
  SubstitutionLow(53) = 8
  SubstitutionLow(54) = 1
  SubstitutionLow(55) = 7
  SubstitutionLow(56) = 8
  SubstitutionLow(57) = 8
  SubstitutionLow(58) = 7
  SubstitutionLow(59) = 1
  SubstitutionLow(60) = 3
  SubstitutionLow(61) = 2
  SubstitutionLow(62) = 7
  SubstitutionLow(63) = 7
  SubstitutionLow(64) = 1
  SubstitutionLow(65) = 8
  SubstitutionLow(66) = 0
  SubstitutionLow(67) = 3
  SubstitutionLow(68) = 7
  SubstitutionLow(69) = 5
  SubstitutionLow(70) = 2
  SubstitutionLow(71) = 6
  SubstitutionLow(72) = 4
  SubstitutionLow(73) = 0
  SubstitutionLow(74) = 9
  SubstitutionLow(75) = 9
  SubstitutionLow(76) = 7
  SubstitutionLow(77) = 7
  SubstitutionLow(78) = 4
  SubstitutionLow(79) = 6
  SubstitutionLow(80) = 2
  SubstitutionLow(81) = 0
  SubstitutionLow(82) = 0
  SubstitutionLow(83) = 1
  SubstitutionLow(84) = 7
  SubstitutionLow(85) = 3
  SubstitutionLow(86) = 6
  SubstitutionLow(87) = 6
  SubstitutionLow(88) = 1
  SubstitutionLow(89) = 1
  SubstitutionLow(90) = 2
  SubstitutionLow(91) = 4
  SubstitutionLow(92) = 5
  SubstitutionLow(93) = 9
  SubstitutionLow(94) = 8
  SubstitutionLow(95) = 2
  SubstitutionLow(96) = 8
  SubstitutionLow(97) = 8
  SubstitutionLow(98) = 3
  SubstitutionLow(99) = 5
  NumColorsFree = 1
  NumBits = GetDeviceCaps(frm3DMaze.hDC, PLANES) * GetDeviceCaps(frm3DMaze.hDC, BITSPIXEL)
  If NumBits >= 31 Then
    UsePalette = False
  Else
    Do While (NumBits > 0)
      NumColorsFree = 2 * NumColorsFree
      NumBits = NumBits - 1
    Loop
    NumColorsFree = NumColorsFree - GetDeviceCaps(frm3DMaze.hDC, COLORS)
    If NumColorsFree < 16 Then
      UsePalette = False
    Else
      UsePalette = True
    End If
  End If
  LogicalPalette.palVersion = 3 * 256
  LogicalPalette.palNumEntries = 16
  For ColorNum = 0 To NumColors - 4
    ' evenly spaced shades of gray
    Tint = (256 * ColorNum) \ (NumColors - 3)
    LogicalPalette.palPalEntry(ColorNum).peRed = Tint
    LogicalPalette.palPalEntry(ColorNum).peGreen = Tint
    LogicalPalette.palPalEntry(ColorNum).peBlue = Tint
    LogicalPalette.palPalEntry(ColorNum).peFlags = PC_NOCOLLAPSE
    RedGreenBlue(ColorNum) = RGB(Tint, Tint, Tint)
  Next ColorNum
  LogicalPalette.palPalEntry(BackoutColor).peRed = 255
  LogicalPalette.palPalEntry(BackoutColor).peGreen = 255
  LogicalPalette.palPalEntry(BackoutColor).peBlue = 0
  LogicalPalette.palPalEntry(BackoutColor).peFlags = PC_NOCOLLAPSE
  RedGreenBlue(BackoutColor) = RGB(255, 255, 0)
  LogicalPalette.palPalEntry(AdvanceColor).peRed = 0
  LogicalPalette.palPalEntry(AdvanceColor).peGreen = 255
  LogicalPalette.palPalEntry(AdvanceColor).peBlue = 0
  LogicalPalette.palPalEntry(AdvanceColor).peFlags = PC_NOCOLLAPSE
  RedGreenBlue(AdvanceColor) = RGB(0, 255, 0)
  LogicalPalette.palPalEntry(SolutionColor).peRed = 255
  LogicalPalette.palPalEntry(SolutionColor).peGreen = 0
  LogicalPalette.palPalEntry(SolutionColor).peBlue = 0
  LogicalPalette.palPalEntry(SolutionColor).peFlags = PC_NOCOLLAPSE
  RedGreenBlue(SolutionColor) = RGB(255, 0, 0)
  If UsePalette Then
    PaletteHandle = CreatePalette(LogicalPalette)
  End If

  HexDeltaY(0, 0) = -1
  HexDeltaX(0, 0) = -2
  HexDeltaY(1, 0) = 1
  HexDeltaX(1, 0) = -2
  HexDeltaY(2, 0) = -2
  HexDeltaX(2, 0) = 0
  HexDeltaY(3, 0) = 2
  HexDeltaX(3, 0) = 0
  HexDeltaY(4, 0) = -1
  HexDeltaX(4, 0) = 2
  HexDeltaY(5, 0) = 1
  HexDeltaX(5, 0) = 2
  DeltaIndex2 = 0
  For DeltaIndex1a = 0 To 5
    For DeltaIndex1b = 0 To 5
      If DeltaIndex1a <> DeltaIndex1b Then
        For DeltaIndex1c = 0 To 5
          If (DeltaIndex1a <> DeltaIndex1c) And (DeltaIndex1b <> DeltaIndex1c) Then
            For DeltaIndex1d = 0 To 5
              If (DeltaIndex1a <> DeltaIndex1d) And (DeltaIndex1b <> DeltaIndex1d) And (DeltaIndex1c <> DeltaIndex1d) Then
                For DeltaIndex1e = 0 To 5
                  If (DeltaIndex1a <> DeltaIndex1e) And (DeltaIndex1b <> DeltaIndex1e) And (DeltaIndex1c <> DeltaIndex1e) And (DeltaIndex1d <> DeltaIndex1e) Then
                    For DeltaIndex1f = 0 To 5
                      If (DeltaIndex1a <> DeltaIndex1f) And (DeltaIndex1b <> DeltaIndex1f) And (DeltaIndex1c <> DeltaIndex1f) And (DeltaIndex1d <> DeltaIndex1f) And (DeltaIndex1e <> DeltaIndex1f) Then
                        HexDeltaX(DeltaIndex1a, DeltaIndex2) = HexDeltaX(0, 0)
                        HexDeltaY(DeltaIndex1a, DeltaIndex2) = HexDeltaY(0, 0)
                        HexDeltaX(DeltaIndex1b, DeltaIndex2) = HexDeltaX(1, 0)
                        HexDeltaY(DeltaIndex1b, DeltaIndex2) = HexDeltaY(1, 0)
                        HexDeltaX(DeltaIndex1c, DeltaIndex2) = HexDeltaX(2, 0)
                        HexDeltaY(DeltaIndex1c, DeltaIndex2) = HexDeltaY(2, 0)
                        HexDeltaX(DeltaIndex1d, DeltaIndex2) = HexDeltaX(3, 0)
                        HexDeltaY(DeltaIndex1d, DeltaIndex2) = HexDeltaY(3, 0)
                        HexDeltaX(DeltaIndex1e, DeltaIndex2) = HexDeltaX(4, 0)
                        HexDeltaY(DeltaIndex1e, DeltaIndex2) = HexDeltaY(4, 0)
                        HexDeltaX(DeltaIndex1f, DeltaIndex2) = HexDeltaX(5, 0)
                        HexDeltaY(DeltaIndex1f, DeltaIndex2) = HexDeltaY(5, 0)
                        DeltaIndex2 = DeltaIndex2 + 1
                      End If
                    Next DeltaIndex1f
                  End If
                Next DeltaIndex1e
              End If
            Next DeltaIndex1d
          End If
        Next DeltaIndex1c
      End If
    Next DeltaIndex1b
  Next DeltaIndex1a
  SqrDeltaY(0, 0) = 0
  SqrDeltaX(0, 0) = -1
  SqrDeltaY(1, 0) = 1
  SqrDeltaX(1, 0) = 0
  SqrDeltaY(2, 0) = 0
  SqrDeltaX(2, 0) = 1
  SqrDeltaY(3, 0) = -1
  SqrDeltaX(3, 0) = 0
  DeltaIndex2 = 0
  For DeltaIndex1a = 0 To 3
    For DeltaIndex1b = 0 To 3
      If DeltaIndex1a <> DeltaIndex1b Then
        For DeltaIndex1c = 0 To 3
          If (DeltaIndex1a <> DeltaIndex1c) And (DeltaIndex1b <> DeltaIndex1c) Then
            For DeltaIndex1d = 0 To 3
              If (DeltaIndex1a <> DeltaIndex1d) And (DeltaIndex1b <> DeltaIndex1d) And (DeltaIndex1c <> DeltaIndex1d) Then
                SqrDeltaX(DeltaIndex1a, DeltaIndex2) = SqrDeltaX(0, 0)
                SqrDeltaY(DeltaIndex1a, DeltaIndex2) = SqrDeltaY(0, 0)
                SqrDeltaX(DeltaIndex1b, DeltaIndex2) = SqrDeltaX(1, 0)
                SqrDeltaY(DeltaIndex1b, DeltaIndex2) = SqrDeltaY(1, 0)
                SqrDeltaX(DeltaIndex1c, DeltaIndex2) = SqrDeltaX(2, 0)
                SqrDeltaY(DeltaIndex1c, DeltaIndex2) = SqrDeltaY(2, 0)
                SqrDeltaX(DeltaIndex1d, DeltaIndex2) = SqrDeltaX(3, 0)
                SqrDeltaY(DeltaIndex1d, DeltaIndex2) = SqrDeltaY(3, 0)
                DeltaIndex2 = DeltaIndex2 + 1
              End If
            Next DeltaIndex1d
          End If
        Next DeltaIndex1c
      End If
    Next DeltaIndex1b
  Next DeltaIndex1a
  Sqrt3 = Sqr(3#)
End Sub


Private Sub Form_Paint()
  mnuActionItem(1).Enabled = False
  mnuActionItem(2).Enabled = False
  Paint = True
  State = 0
  If Not AlreadyPainting Then
    AlreadyPainting = True
    Timer1.Enabled = True
  End If
End Sub

Private Sub Form_Resize()
  If WindowState = 1 Then
    Minimized = True
    Cls
    State = 5
    AlreadyPainting = False
  Else
    If ScaleHeight < 3 * Text1.Height Then
      Minimized = False
      Cls
      State = 5
      AlreadyPainting = False
      Text1.Text = "This window is too small!"
    Else
      VScroll1.Height = ScaleHeight - Text1.Height
      VScroll1.Left = ScaleWidth - VScroll1.Width
      Text1.Top = ScaleHeight - Text1.Height
      Text1.Width = ScaleWidth
      Paint = True
      State = 0
      If (Not Minimized) Then
        Resize = True
        UserHasSolved = False
        SolutionDisplayed = False
        Seed = Str(Timer)
      End If
      Minimized = False
      Refresh
    End If
  End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
  State = 6
  Timer1.Enabled = False
  Erase Stack
  Erase UserPage
  Erase ComputerPage
End Sub



Private Sub mnuActionItem_Click(Index As Integer)
  Select Case Index
    Case 0
      SolutionDisplayed = False
      Call Form_Resize
    Case 1
      SolutionDisplayed = True
      Text1.Text = ""
      If mnuStyleItem(0).Checked Then
        Call HexDisplaySolution(MaxY, ComputerPage, XMax, XOffset, YMax, CosTilt, SinTilt, PixelsPerX, PixelsPerZ, RelDistOfUserFromScreen)
      Else
        Call SqrDisplaySolution(MaxY, ComputerPage, XMax, XOffset, YMax, CosTilt, SinTilt, PixelsPerX, PixelsPerZ, RelDistOfUserFromScreen)
      End If
    Case 2
      UserHasSolved = False
      SolutionDisplayed = False
      Paint = True
      State = 0
      For UserX = 0 To MaxX
        For UserY = 0 To MaxY
          If ComputerPage(UserY, UserX) = 0 Then
            UserPage(UserY, UserX) = 0
          Else
            UserPage(UserY, UserX) = 2
          End If
        Next UserY
      Next UserX
      If mnuStyleItem(0).Checked Then
        UserX = 3
        UserXRelative = 1#
        UserY = 2
        UserYRelative = Sqrt3 / 2#
      Else
        UserX = 1
        UserXRelative = (RelativeWidthOfWall + 1#) / 2#
        UserY = 1
        UserYRelative = (RelativeWidthOfWall + 1#) / 2#
      End If
      UserPage(UserY, UserX) = 1
      Refresh
    Case 4
      State = 6
      Timer1.Enabled = False
      Erase Stack
      Erase UserPage
      Erase ComputerPage
      End
    Case Else
  End Select
End Sub

Private Sub mnuHelpItem_Click(Index As Integer)
  Dim rc As Integer
  rc = MsgBox("Maze 'O' MaNiA!" + Chr(13) + Chr(13) + "Copyright " + Chr(169) + " 1998 Christopher D. Fennell (webmaster203@juno.com)" + Chr(13) + Chr(13) + "This application may be distributed without payments to Christopher D. Fennell." + Chr(13) + Chr(13) + "As per Microsoft's license for Visual Basic 5.0, the end-user may not distribute the components having names starting with other than " + Chr(34) + "Maze 'O' MaNiA!" + Chr(34) + ".", vbOKOnly, "About MaZe 'O' MaNiA")
End Sub

Private Sub mnuStyleItem_Click(Index As Integer)
  Select Case Index
    Case 0
      If mnuStyleItem(1).Checked Then
        mnuStyleItem(0).Checked = True
        mnuStyleItem(1).Checked = False
        SolutionDisplayed = False
        Call Form_Resize
      End If
    Case 1
      If mnuStyleItem(0).Checked Then
        mnuStyleItem(0).Checked = False
        mnuStyleItem(1).Checked = True
        SolutionDisplayed = False
        Call Form_Resize
      End If
  End Select
End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  If mnuStyleItem(0).Checked Then
    Call HexOutputMaze
  Else
    Call SqrOutputMaze
  End If
End Sub

Private Sub VScroll1_Change()
  Tilt = 90 - VScroll1.Value
  Paint = True
  State = 0
  If Not AlreadyPainting Then Call Form_Paint
End Sub

Private Sub VScroll1_Scroll()
  If AlreadyPainting Then
    Tilt = 90 - VScroll1.Value
    Paint = True
    State = 0
  End If
End Sub



