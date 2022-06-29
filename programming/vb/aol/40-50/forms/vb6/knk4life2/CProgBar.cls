VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Copyright (C)Karl E. Peterson, 1995
' All Rights Reserved.
'
' Written for publication in March 1996 issue
' of Visual Basic Programmer's Journal.
' May be freely used in your applications.
'
Option Explicit
'
' Set aside storage for private member variables.
'
Private m_Bar As PictureBox
Private m_Val As Long
Private m_Min As Long
Private m_Max As Long
Private m_fColor As Long
Private m_bColor As Long
Private m_CellWidth As Integer
Private m_CellHeight As Integer
Private m_xMargin As Integer
Private m_yMargin As Integer
Private m_Cells As Integer

' ********************************************
'  Initialize
' ********************************************
Private Sub Class_Initialize()
   '
   ' Set default values for class properties.
   '
   Set m_Bar = Nothing
   m_Val = 0
   m_Min = 0
   m_Max = 100
   m_fColor = &HFF&
   m_bColor = &H0&
End Sub

' ********************************************
'  Public Properties
' ********************************************
Public Property Set Canvas(NewObj As Object)
   '
   ' Set new PictureBox as Canvas property.
   '
   If TypeOf NewObj Is PictureBox Then
      Set m_Bar = NewObj
      m_Bar.ForeColor = m_fColor
      m_Bar.BackColor = m_bColor
      m_Bar.ScaleMode = vbPixels
      ResizeEx
   Else
      Err.Raise Number:=vbObjectError + 1, _
                source:="CProgressBar.Canvas", _
                Description:="Canvas property must be of type PictureBox."
   End If
End Property

Public Property Get Canvas() As Object
   '
   ' Return PictureBox as Canvas property.
   '
   Set Canvas = m_Bar
End Property

Public Property Let Value(NewVal As Long)
   Dim PrevVal As Long
   '
   ' Set aside previous value.
   '
   PrevVal = m_Val
   '
   ' Set Value, checking for in-range.
   '
   Select Case NewVal
      Case Is > m_Max
         m_Val = m_Max
      Case Is < m_Min
         m_Val = m_Min
      Case Else
         m_Val = NewVal
   End Select
   '
   ' Update display
   '
   Refresh
End Property

Public Property Get Value() As Long
   '
   ' Return Value
   '
   Value = m_Val
End Property

Public Property Let min(NewVal As Long)
   '
   ' Set minimum value, and refresh display.
   '
   If NewVal < m_Max Then
      m_Min = NewVal
      Refresh
   Else
      Err.Raise Number:=vbObjectError + 2, _
                source:="CProgressBar.Min", _
                Description:="Min property must be less than Max."
   End If
End Property

Public Property Get min() As Long
   '
   ' Return minimum value
   '
   min = m_Min
End Property

Public Property Let max(NewVal As Long)
   '
   ' Set maximum value, and refresh display.
   '
   If NewVal > m_Min Then
      m_Max = NewVal
      Refresh
   Else
      Err.Raise Number:=vbObjectError + 3, _
                source:="CProgressBar.Max", _
                Description:="Max property must be greater than Min."
   End If
End Property

Public Property Get max() As Long
   '
   ' Return maximum value
   '
   max = m_Max
End Property

Public Property Let ForeColor(NewVal As Long)
   '
   ' Set foreground color
   '
   m_fColor = NewVal
   If Not m_Bar Is Nothing Then
      m_Bar.ForeColor = m_fColor
      Refresh
   End If
End Property

Public Property Get ForeColor() As Long
   '
   ' Return foreground color
   '
   ForeColor = m_fColor
End Property

Public Property Let BackColor(NewVal As Long)
   '
   ' Set background color
   '
   m_bColor = NewVal
   If Not m_Bar Is Nothing Then
      m_Bar.BackColor = m_bColor
      Refresh
   End If
End Property

Public Property Get BackColor() As Long
   '
   ' Return background color
   '
   BackColor = m_bColor
End Property

' ********************************************
'  Public Methods
' ********************************************
Public Sub Refresh(Optional ClearFirst)
   '
   ' Update display, clearing it if requested.
   '
   If Not m_Bar Is Nothing Then
      If Not IsMissing(ClearFirst) Then
         If ClearFirst Then
            m_Bar.Cls
         End If
      End If
      RedrawMe
   End If
End Sub

Public Sub Resize()
   '
   ' Calc new cell dimensions.
   '
   CalcCellSize
   '
   ' Control has been resized, clear and update.
   '
   Refresh True
End Sub

Public Sub ResizeEx()
   Dim xActual As Integer
   Dim xPicture As Integer
   Dim dLarge As Integer
   Dim dSmall As Integer
   Dim xExtra As Integer
   Static Recursing As Boolean
   '
   ' Bail out if recursing back into this routine.
   '
   If Recursing Then Exit Sub
   '
   ' Calc new cell dimensions.
   '
   CalcCellSize
   '
   ' Calc differences between one more and one less cell.
   '
   xActual = m_xMargin + m_Cells * (m_CellWidth + m_xMargin)
   xPicture = m_Bar.ScaleWidth
   dLarge = xActual - xPicture
   dSmall = xPicture - (xActual - m_CellWidth - m_xMargin)
   '
   ' If less diff to shrink, adjust number of cells.
   '
   If dSmall <= dLarge Then
      m_Cells = m_Cells - 1
      xActual = m_xMargin + m_Cells * (m_CellWidth + m_xMargin)
   End If
   '
   ' Resize progress bar Canvas
   '
   xExtra = m_Bar.Width - (xPicture * Screen.TwipsPerPixelX)
   Recursing = True
      m_Bar.Width = xActual * Screen.TwipsPerPixelX + xExtra
   Recursing = False
   '
   ' Control has been resized, clear and update.
   '
   Refresh True
End Sub

' ********************************************
'  Private Methods
' ********************************************
Private Sub RedrawMe()
   Dim i As Long
   Dim x As Long
   Dim y As Long
   Dim n As Long
   '
   ' Calc number of live cells to draw.
   '
   n = (m_Val / (m_Max - m_Min)) * m_Cells
   '
   ' Draw live cells.
   '
   m_Bar.ForeColor = m_fColor
   y = m_yMargin + m_CellHeight - 1
   x = m_yMargin
   For i = 1 To n
      m_Bar.Line (x, m_yMargin)-(x + m_CellWidth, y), , BF
      x = x + m_xMargin + m_CellWidth
   Next i
   '
   ' Draw dead cells.
   '
   If n < m_Cells Then
      m_Bar.ForeColor = m_bColor
      For i = n + 1 To m_Cells
         m_Bar.Line (x, m_yMargin)-(x + m_CellWidth, y), , BF
         x = x + m_xMargin + m_CellWidth
      Next i
   End If
End Sub

Private Sub CalcCellSize()
   '
   ' Calc new cell dimensions.
   '
   m_yMargin = 2
   m_xMargin = 3
   m_CellHeight = m_Bar.ScaleHeight - m_yMargin * 2
   m_CellWidth = m_CellHeight * (2 / 3)
   m_Cells = m_Bar.ScaleWidth \ (m_CellWidth + m_xMargin) + 1
End Sub
