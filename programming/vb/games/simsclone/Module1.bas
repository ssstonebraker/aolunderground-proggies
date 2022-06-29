Attribute VB_Name = "Module1"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCAND = &H8800C6
Public Const SRCPAINT = &HEE0086
Public Const SRCCOPY = &HCC0020

Public Const MAPROWS = 28
Public Const MAPCOLS = 24

Public Type TILE
    StructureID As Integer
    EarthTile As Integer
    LandValue As Long
    Population As Long
    Growth As Integer
    ColorFlag As Integer
    Name As String * 20
    ClassFlag As Integer
End Type: Public T(0 To MAPROWS, 0 To MAPCOLS) As TILE, Cash As Long

'Selection Memory
Public Type MOUSESTAT
    selectedPurchase As Integer
    price As Long
End Type: Public MS As MOUSESTAT

'Date system
Public CurrentSeason As Integer, CurMonth As Integer, CurYear As Integer

'Mechanix Vars (for loops, mouse, stats, ect...)
Public CURS As Integer, CURC As Integer, CURL As Integer, MouseOUT As Boolean
Public TotalPOP As Long, SafetyCount As Integer, rn As Integer, rn2 As Integer, Crime As Integer
Public i As Integer, ii As Integer, iii As Integer, iiii As Integer, Drawing As Boolean
Public CX As Single, CY As Single, NX As Integer, NY As Integer, NX1 As Integer, NY1 As Integer, NX2 As Integer, NY2 As Integer, W As Integer, H As Integer

Sub filesave()
Open App.Path & "\save.bin" For Binary As #1
Put #1, , Cash
Put #1, , CurrentSeason
Put #1, , CurYear
Put #1, , CurMonth

For i = 0 To MAPROWS
For ii = 0 To MAPCOLS
Put #1, , T(i, ii)
Next
Next

Close #1
End Sub
Sub fileload()
Open App.Path & "\save.bin" For Binary As #1
Get #1, , Cash
Get #1, , CurrentSeason
Get #1, , CurYear
Get #1, , CurMonth

For i = 0 To MAPROWS
For ii = 0 To MAPCOLS
Get #1, , T(i, ii)
Next
Next

Close #1
End Sub
Public Function RndRange(ByVal intMin As Integer, ByVal intMax As Integer) As Integer
RndRange = Int(Rnd * (intMax - intMin + 1)) + intMin
End Function
Sub initTILES() 'Sets Default Tile Values
For i = 0 To MAPROWS
For ii = 0 To MAPCOLS
T(i, ii).StructureID = 100
T(i, ii).EarthTile = Rnd * 8
T(i, ii).LandValue = 100
T(i, ii).Population = 0
T(i, ii).Growth = 0
T(i, ii).ColorFlag = 0
T(i, ii).Name = "Open Space"
Next
Next
CurrentSeason = 1
Cash = 1000000
CurMonth = 1
CurYear = 1900
End Sub
Function ReturnMstr(inte As Integer) As String
'Returns Month + Changes Seasons

Select Case inte 'Process input
Case 1: ReturnMstr = "JAN": CurrentSeason = 1
Case 2: ReturnMstr = "FEB": CurrentSeason = 1
Case 3: ReturnMstr = "MAR": CurrentSeason = 1
Case 4: ReturnMstr = "APR": CurrentSeason = 2
Case 5: ReturnMstr = "MAY": CurrentSeason = 2
Case 6: ReturnMstr = "JUNE": CurrentSeason = 3
Case 7: ReturnMstr = "JULY": CurrentSeason = 3
Case 8: ReturnMstr = "AUG": CurrentSeason = 3
Case 9: ReturnMstr = "SEP": CurrentSeason = 3
Case 10: ReturnMstr = "OCT": CurrentSeason = 4
Case 11: ReturnMstr = "NOV": CurrentSeason = 4
Case 12: ReturnMstr = "DEC": CurrentSeason = 1
End Select
End Function
Sub DrawBacks() 'Draw Turf backgrounds.
For i = 0 To MAPROWS
For ii = 0 To MAPCOLS
BitBlt GFX.Turf(1).hDC, i * 13, ii * 13, 13, 13, Form1.TURFWinter(T(i, ii).EarthTile).hDC, 0, 0, SRCCOPY
BitBlt GFX.Turf(2).hDC, i * 13, ii * 13, 13, 13, Form1.TURFSpring(T(i, ii).EarthTile).hDC, 0, 0, SRCCOPY
BitBlt GFX.Turf(3).hDC, i * 13, ii * 13, 13, 13, Form1.TURFSummer(T(i, ii).EarthTile).hDC, 0, 0, SRCCOPY
BitBlt GFX.Turf(4).hDC, i * 13, ii * 13, 13, 13, Form1.TURFFall(T(i, ii).EarthTile).hDC, 0, 0, SRCCOPY
Next
Next
End Sub
Sub DrawBoard() 'Draw Structures Sprite
Drawing = True
On Error Resume Next
For i = 0 To MAPROWS
For ii = 0 To MAPCOLS
Select Case T(i, ii).StructureID
Case 0
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 13, 14, GFX.h1M.hDC, 0, 0, SRCCOPY
Select Case T(i, ii).ColorFlag
Case 0: BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 14, GFX.h1Sbr.hDC, 0, 0, SRCCOPY
Case 1: BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 14, GFX.h1SMon.hDC, 0, 0, SRCCOPY
Case 2: BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 14, GFX.h1Sgr.hDC, 0, 0, SRCCOPY
End Select
Case 1
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 13, 13, GFX.Picture1.hDC, 0, 0, SRCCOPY
Select Case T(i, ii).ColorFlag
Case 0: BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 13, GFX.h2Sbr.hDC, 0, 0, SRCCOPY
Case 1: BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 13, GFX.h2SMon.hDC, 0, 0, SRCCOPY
Case 2: BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 13, GFX.h2Sgr.hDC, 0, 0, SRCCOPY
End Select
Case 2
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 13, 14, GFX.c1M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 14, GFX.c1S.hDC, 0, 0, SRCCOPY
Case 3
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 13, 14, GFX.c2M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 14, GFX.c2S.hDC, 0, 0, SRCCOPY
Case 4
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 13, 14, GFX.c3M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 14, GFX.c3S.hDC, 0, 0, SRCCOPY
Case 5
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 13, 14, GFX.c4M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 14, GFX.c4S.hDC, 0, 0, SRCCOPY
Case 6
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 13, 14, GFX.i1M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 14, GFX.i1S.hDC, 0, 0, SRCCOPY
Case 7
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 13, 14, GFX.i2M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 14, GFX.i2S.hDC, 0, 0, SRCCOPY
Case 8
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 13, 14, GFX.i3M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 14, GFX.i3S.hDC, 0, 0, SRCCOPY
Case 9
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 13, 14, GFX.i4M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 14, GFX.i4S.hDC, 0, 0, SRCCOPY
Case 10
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 13, 14, GFX.rT4M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 14, GFX.rT4s.hDC, 0, 0, SRCCOPY
Case 11
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 13, 14, GFX.rT1M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 14, GFX.rT1s.hDC, 0, 0, SRCCOPY
Case 12
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 13, 14, GFX.rC3M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 14, GFX.rC3s.hDC, 0, 0, SRCCOPY
Case 13
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 13, 14, GFX.rC4M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 14, GFX.rC4s.hDC, 0, 0, SRCCOPY
Case 14
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 13, 14, GFX.rLRM.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 14, GFX.rLRs.hDC, 0, 0, SRCCOPY
Case 15
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 13, 14, GFX.rT2M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 14, GFX.rT2s.hDC, 0, 0, SRCCOPY
Case 16
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 13, 14, GFX.rT3M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 14, GFX.rT3s.hDC, 0, 0, SRCCOPY
Case 17
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 13, 14, GFX.rC1M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 14, GFX.rC1s.hDC, 0, 0, SRCCOPY
Case 18
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 13, 14, GFX.rC2M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 14, GFX.rC2s.hDC, 0, 0, SRCCOPY
Case 19
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 13, 14, GFX.rUDM.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 14, GFX.rUDs.hDC, 0, 0, SRCCOPY
Case 20
BitBlt Form1.BGPB2.hDC, i * 13 - 3, ii * 13 - 3, 26, 16, GFX.RoadIm.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13 - 3, ii * 13 - 3, 26, 16, GFX.RoadIs.hDC, 0, 0, SRCCOPY

Select Case T(i - 1, ii).StructureID
Case 10
BitBlt Form1.BGPB2.hDC, (i - 1) * 13, ii * 13, 13, 14, GFX.rT4M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, (i - 1) * 13, ii * 13, 13, 14, GFX.rT4s.hDC, 0, 0, SRCCOPY
Case 11
BitBlt Form1.BGPB2.hDC, (i - 1) * 13, ii * 13, 13, 14, GFX.rT1M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, (i - 1) * 13, ii * 13, 13, 14, GFX.rT1s.hDC, 0, 0, SRCCOPY
Case 12
BitBlt Form1.BGPB2.hDC, (i - 1) * 13, ii * 13, 13, 14, GFX.rC3M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, (i - 1) * 13, ii * 13, 13, 14, GFX.rC3s.hDC, 0, 0, SRCCOPY
Case 13
BitBlt Form1.BGPB2.hDC, (i - 1) * 13, ii * 13, 13, 14, GFX.rC4M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, (i - 1) * 13, ii * 13, 13, 14, GFX.rC4s.hDC, 0, 0, SRCCOPY
Case 14
BitBlt Form1.BGPB2.hDC, (i - 1) * 13, ii * 13, 13, 14, GFX.rLRM.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, (i - 1) * 13, ii * 13, 13, 14, GFX.rLRs.hDC, 0, 0, SRCCOPY
Case 15
BitBlt Form1.BGPB2.hDC, (i - 1) * 13, ii * 13, 13, 14, GFX.rT2M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, (i - 1) * 13, ii * 13, 13, 14, GFX.rT2s.hDC, 0, 0, SRCCOPY
Case 16
BitBlt Form1.BGPB2.hDC, (i - 1) * 13, ii * 13, 13, 14, GFX.rT3M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, (i - 1) * 13, ii * 13, 13, 14, GFX.rT3s.hDC, 0, 0, SRCCOPY
Case 17
BitBlt Form1.BGPB2.hDC, (i - 1) * 13, ii * 13, 13, 14, GFX.rC1M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, (i - 1) * 13, ii * 13, 13, 14, GFX.rC1s.hDC, 0, 0, SRCCOPY
Case 18
BitBlt Form1.BGPB2.hDC, (i - 1) * 13, ii * 13, 13, 14, GFX.rC2M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, (i - 1) * 13, ii * 13, 13, 14, GFX.rC2s.hDC, 0, 0, SRCCOPY
Case 19
BitBlt Form1.BGPB2.hDC, (i - 1) * 13, ii * 13, 13, 14, GFX.rUDM.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, (i - 1) * 13, ii * 13, 13, 14, GFX.rUDs.hDC, 0, 0, SRCCOPY
Case 20
BitBlt Form1.BGPB2.hDC, (i - 1) * 13, ii * 13, 13, 14, GFX.RoadIm.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, (i - 1) * 13, ii * 13, 13, 14, GFX.RoadIs.hDC, 0, 0, SRCCOPY
End Select
Select Case T(i, ii - 1).StructureID
Case 10
BitBlt Form1.BGPB2.hDC, i * 13, (ii - 1) * 13, 13, 14, GFX.rT4M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, (ii - 1) * 13, 13, 14, GFX.rT4s.hDC, 0, 0, SRCCOPY
Case 11
BitBlt Form1.BGPB2.hDC, i * 13, (ii - 1) * 13, 13, 14, GFX.rT1M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, (ii - 1) * 13, 13, 14, GFX.rT1s.hDC, 0, 0, SRCCOPY
Case 12
BitBlt Form1.BGPB2.hDC, i * 13, (ii - 1) * 13, 13, 14, GFX.rC3M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, (ii - 1) * 13, 13, 14, GFX.rC3s.hDC, 0, 0, SRCCOPY
Case 13
BitBlt Form1.BGPB2.hDC, i * 13, (ii - 1) * 13, 13, 14, GFX.rC4M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, (ii - 1) * 13, 13, 14, GFX.rC4s.hDC, 0, 0, SRCCOPY
Case 14
BitBlt Form1.BGPB2.hDC, i * 13, (ii - 1) * 13, 13, 14, GFX.rLRM.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, (ii - 1) * 13, 13, 14, GFX.rLRs.hDC, 0, 0, SRCCOPY
Case 15
BitBlt Form1.BGPB2.hDC, i * 13, (ii - 1) * 13, 13, 14, GFX.rT2M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, (ii - 1) * 13, 13, 14, GFX.rT2s.hDC, 0, 0, SRCCOPY
Case 16
BitBlt Form1.BGPB2.hDC, i * 13, (ii - 1) * 13, 13, 14, GFX.rT3M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, (ii - 1) * 13, 13, 14, GFX.rT3s.hDC, 0, 0, SRCCOPY
Case 17
BitBlt Form1.BGPB2.hDC, i * 13, (ii - 1) * 13, 13, 14, GFX.rC1M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, (ii - 1) * 13, 13, 14, GFX.rC1s.hDC, 0, 0, SRCCOPY
Case 18
BitBlt Form1.BGPB2.hDC, i * 13, (ii - 1) * 13, 13, 14, GFX.rC2M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, (ii - 1) * 13, 13, 14, GFX.rC2s.hDC, 0, 0, SRCCOPY
Case 19
BitBlt Form1.BGPB2.hDC, i * 13, (ii - 1) * 13, 13, 14, GFX.rUDM.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, (ii - 1) * 13, 13, 14, GFX.rUDs.hDC, 0, 0, SRCCOPY
Case 20
BitBlt Form1.BGPB2.hDC, i * 13, (ii - 1) * 13, 13, 14, GFX.RoadIm.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, (ii - 1) * 13, 13, 14, GFX.RoadIs.hDC, 0, 0, SRCCOPY
End Select

Case 21
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 29, 29, GFX.BR1M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 29, 29, GFX.BR1s.hDC, 0, 0, SRCCOPY
Case 22
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 29, 29, GFX.BR2M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 29, 29, GFX.BR2s.hDC, 0, 0, SRCCOPY
Case 23
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 29, 29, GFX.BR3M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 29, 29, GFX.BR3s.hDC, 0, 0, SRCCOPY
Case 24
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 29, 29, GFX.BR4M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 29, 29, GFX.BR4s.hDC, 0, 0, SRCCOPY
Case 25
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 29, 29, GFX.BC1M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 29, 29, GFX.BC1s.hDC, 0, 0, SRCCOPY
Case 26
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 29, 29, GFX.BC2M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 29, 29, GFX.BC2s.hDC, 0, 0, SRCCOPY
Case 27
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 29, 29, GFX.BC3M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 29, 29, GFX.BC3s.hDC, 0, 0, SRCCOPY
Case 28
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 29, 29, GFX.BI1M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 29, 29, GFX.BI1s.hDC, 0, 0, SRCCOPY
Case 29
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 29, 29, GFX.BI2M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 29, 29, GFX.BI2s.hDC, 0, 0, SRCCOPY
Case 30
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 29, 29, GFX.BI3M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 29, 29, GFX.BI3s.hDC, 0, 0, SRCCOPY
Case 31
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 29, 29, GFX.ESM.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 29, 29, GFX.ES1.hDC, 0, 0, SRCCOPY
Case 32
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 29, 29, GFX.ESM.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 29, 29, GFX.ES2.hDC, 0, 0, SRCCOPY
Case 33
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 26, 27, GFX.Tree1M.hDC, 0, 0, SRCCOPY
Select Case CurrentSeason
Case 1
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 13, GFX.Tree1sW.hDC, 0, 0, SRCCOPY
Case 2, 3
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 13, GFX.Tree1sSS.hDC, 0, 0, SRCCOPY
Case 4
Select Case T(i, ii).ColorFlag
Case 1: BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 13, GFX.Tree1sF.hDC, 0, 0, SRCCOPY
Case Else: BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 14, GFX.Tree1sF2.hDC, 0, 0, SRCCOPY
End Select
End Select
Case 34
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 26, 27, GFX.Tree2M.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 13, GFX.Tree2s.hDC, 0, 0, SRCCOPY
Case 35
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 26, 27, GFX.BusM.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 13, GFX.BusS.hDC, 0, 0, SRCCOPY
Case 36
Select Case T(i, ii).ClassFlag
Case 0
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 13, 13, GFX.GC1(5).hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 13, GFX.GC1(0).hDC, 0, 0, SRCCOPY
Case 1
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 13, 13, GFX.GC1(4).hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 13, GFX.GC1(1).hDC, 0, 0, SRCCOPY
Case 2
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 13, 13, GFX.GC1(3).hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 13, GFX.GC1(2).hDC, 0, 0, SRCCOPY
End Select
Case 100
BitBlt Form1.BGPB2.hDC, i * 13, ii * 13, 13, 13, Form1.Blank.hDC, 0, 0, SRCCOPY
BitBlt Form1.BGPB.hDC, i * 13, ii * 13, 13, 13, Form1.BlackBird.hDC, 0, 0, SRCCOPY
End Select
Next
Next

Drawing = False
End Sub
Sub DoCURS(X As Integer, Y As Integer, Dest As Object, CNum As Integer, Clevel As Integer, CSize As Integer)
On Error Resume Next
NX2 = (13 * Int(X / 13)) - 1
NY2 = (13 * Int(Y / 13)) - 1

Select Case CSize
Case 1

    BitBlt Dest.hDC, NX2, NY2, 29, 29, GFX.cur2M.hDC, 0, 0, SRCAND
    
    Select Case CNum
    Case 1
        Select Case Clevel
            Case 1: BitBlt Dest.hDC, NX2, NY2, 15, 15, GFX.CA2.hDC, 0, 0, SRCPAINT
            Case 2: BitBlt Dest.hDC, NX2, NY2, 15, 15, GFX.c31.hDC, 0, 0, SRCPAINT
            Case 3: BitBlt Dest.hDC, NX2, NY2, 15, 15, GFX.c32.hDC, 0, 0, SRCPAINT
        End Select
    Case 2
        Select Case Clevel
            Case 1: BitBlt Dest.hDC, NX2, NY2, 15, 15, GFX.CA2.hDC, 0, 0, SRCPAINT
            Case 2: BitBlt Dest.hDC, NX2, NY2, 15, 15, GFX.c41.hDC, 0, 0, SRCPAINT
            Case 3: BitBlt Dest.hDC, NX2, NY2, 15, 15, GFX.c42.hDC, 0, 0, SRCPAINT
        End Select
    End Select
    
Case 2

    BitBlt Dest.hDC, NX2, NY2, 29, 29, GFX.curM.hDC, 0, 0, SRCAND
    Select Case CNum
    Case 1
        Select Case Clevel
            Case 1: BitBlt Dest.hDC, NX2, NY2, 29, 29, GFX.CA1.hDC, 0, 0, SRCPAINT
            Case 2: BitBlt Dest.hDC, NX2, NY2, 29, 29, GFX.C11.hDC, 0, 0, SRCPAINT
            Case 3: BitBlt Dest.hDC, NX2, NY2, 29, 29, GFX.C12.hDC, 0, 0, SRCPAINT
        End Select
    Case 2
        Select Case Clevel
            Case 1: BitBlt Dest.hDC, NX2, NY2, 29, 29, GFX.CA1.hDC, 0, 0, SRCPAINT
            Case 2: BitBlt Dest.hDC, NX2, NY2, 29, 29, GFX.C21.hDC, 0, 0, SRCPAINT
            Case 3: BitBlt Dest.hDC, NX2, NY2, 29, 29, GFX.C22.hDC, 0, 0, SRCPAINT
        End Select
    End Select
    
End Select
End Sub

Sub bmp_rotate(pic1 As PictureBox, pic2 As PictureBox, ByVal theta As Single)
On Error Resume Next
    'Rotate the image in a picture box.

    'pic1 is the picture box with the bitmap to rotate

    'pic2 is the picture box to receive the rotated bitmap

    'theta is the angle of rotation

    
    Dim c1x As Integer, c1y As Integer
    Dim c2x As Integer, c2y As Integer
    Dim a As Single
    Dim p1x As Integer, p1y As Integer
    Dim p2x As Integer, p2y As Integer
    Dim n As Integer, R As Integer
    Dim c0 As Long, c1 As Long, C2 As Long, c3 As Long
    
    c1x = pic1.ScaleWidth \ 2
    c1y = pic1.ScaleHeight \ 2
    c2x = pic2.ScaleWidth \ 2
    c2y = pic2.ScaleHeight \ 2
    If c2x < c2y Then n = c2y Else n = c2x
    n = n - 1
    pic1hDC = pic1.hDC
    pic2hDC = pic2.hDC


    For p2x = 0 To n / 2 ': DoEvents


        For p2y = 0 To n / 2
            If p2x = 0 Then a = Pi / 2 Else a = Atn(p2y / p2x)
            R = Sqr(1& * p2x * p2x + 1& * p2y * p2y)
            p1x = R * Cos(a + theta!)
            p1y = R * Sin(a + theta!)
            c0& = pic1.Point(c1x + p1x, c1y + p1y)
            c1& = pic1.Point(c1x - p1x, c1y - p1y)
            C2& = pic1.Point(c1x + p1y, c1y - p1x)
            c3& = pic1.Point(c1x - p1y, c1y + p1x)
            If c0& <> -1 Then pic2.PSet (c2x + p2x, c2y + p2y), c0&
            If c1& <> -1 Then pic2.PSet (c2x - p2x, c2y - p2y), c1&
            If C2& <> -1 Then pic2.PSet (c2x + p2y, c2y - p2x), C2&
            If c3& <> -1 Then pic2.PSet (c2x - p2y, c2y + p2x), c3&
        Next

        
        
    Next

End Sub
