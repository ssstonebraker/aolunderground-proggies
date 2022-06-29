Attribute VB_Name = "Module1"
Option Explicit


     Type RECT
         Left As Long
         Top As Long
         Right As Long
         Bottom As Long
         End Type


     Declare Function ClipCursor Lib "user32" _
         (lpRect As Any) As Long


     Public Sub DisableTrap(CurForm As Form)


         Dim erg As Long
         Dim NewRect As RECT
         

         With NewRect
             .Left = 0&
             .Top = 0&
             .Right = Screen.Width / Screen.TwipsPerPixelX
             .Bottom = Screen.Height / Screen.TwipsPerPixelY
         End With


         erg& = ClipCursor(NewRect)
     End Sub



     Public Sub EnableTrap(CurForm As Form)


         Dim x As Long, y As Long, erg As Long
         Dim NewRect As RECT
         x& = Screen.TwipsPerPixelX
         y& = Screen.TwipsPerPixelY
         

         With NewRect
             .Left = CurForm.Left / x&
             .Top = CurForm.Top / y&
             .Right = .Left + CurForm.Width / x&
             .Bottom = .Top + CurForm.Height / y&
         End With


         erg& = ClipCursor(NewRect)
     End Sub



