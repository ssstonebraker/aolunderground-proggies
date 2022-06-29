Attribute VB_Name = "Module1"
Option Explicit
'Only one API declaration
'from The Visual Basic 5.0 API declarations document
'Win32API.txt
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Public Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Public Const NOTSRCCOPY = &H330008      ' (DWORD) dest = (NOT source)
Public Const NOTSRCERASE = &H1100A6     ' (DWORD) dest = (NOT src) AND (NOT dest)

Public Type ONE_PIECE
   Width As Long
   Height As Long
   OnBmp_x As Long 'x pos of the piece on the Pieces.BMP file
   OnBmp_y As Long 'y pos
   MaskOnBmp_x As Long 'x pos of the mask on the Pieces.BMP file
   MaskOnBmp_y As Long 'y pos
   Creating_x As Long 'Center the piece on release
   Creating_Piece_No As Long 'Allways release the good one.
   Next_Piece_Pos_x As Long 'Center the piece on the next piece window
   Next_Piece_Pos_y As Long 'Same
   After_Turn_Piece_No As Long 'Which one will be the next piece after turn
   After_Turn_dx As Long 'if it will turn, how much will
   After_Turn_dy As Long 'be the x,y differences on turn
   Number_Of_Positions_To_Check_Left As Long 'Number of positions To check on the left, right and down
   Number_Of_Positions_To_Check_Right As Long
   Number_Of_Positions_To_Check_Down As Long
   Check_These_Positions_Left_x(9) As Long 'X and Y coordinates of the Positions to check if empty or not
   Check_These_Positions_Left_y(9) As Long
   Check_These_Positions_Right_x(9) As Long
   Check_These_Positions_Right_y(9) As Long
   Check_These_Positions_Down_x(9) As Long
   Check_These_Positions_Down_y(9) As Long
   Number_Of_Squares As Long 'It is allways equal to 4
   Piece_Is_This_x(4) As Long
   Piece_Is_This_y(4) As Long
End Type

Type PIECE_PERMISSIONS
   Left(1 To 18, 0 To 11) As Boolean
   Right(1 To 18, 0 To 11) As Boolean
   Vertical(1 To 18, 0 To 11) As Boolean
   Left2(1 To 18, 0 To 11) As Boolean
   Right2(1 To 18, 0 To 11) As Boolean
   Vertical2(1 To 18, 0 To 11) As Boolean
End Type

Public Piece(19) As ONE_PIECE
Public Piece_No_Permission As PIECE_PERMISSIONS
Public Score As Long
Public Vertical_Stop_Time As Single
Public Vx As Single
Public Vy As Single
Public Vy_Level As Single
Public Level As Long
Public x As Single
Public y As Single
Public Prv_x As Single
Public Prv_y As Single
Public Piece_No As Long
Public Prv_Piece_No As Long
Public Next_Piece_No As Long
Public Position_Empty(1 To 18, 0 To 11) As Boolean
Public Copied_To_BMP As Boolean
Public Game_Is_Started As Boolean
Public Game_Is_About_To_End As Boolean
Public Key_Up As Boolean
Public Vertical_Stop_Status As Boolean
Public Left_Stop_Status As Boolean
Public Right_Stop_Status As Boolean
Public Prv_Vertical_Stop_Status As Boolean
Public Prv_Left_Stop_Status As Boolean
Public Prv_Right_Stop_Status As Boolean
Public Right_Move_Requested As Boolean
Public Left_Move_Requested As Boolean
Public Piece_Stopped As Boolean
Sub Draw_The_Piece(DTP_Prv_Piece_No As Long, DTP_Piece_No As Long, DTP_x As Long, DTP_y As Long, DTP_Px As Long, DTP_Py As Long)
   'First, copy the previous background back
   'Then copy the new position content to a blank area on picture box BMP
   'Draw the mask
   'and draw the piece to its new position.
      If Copied_To_BMP = True Then BitBlt Form1.Picture1.hDC, DTP_Px, DTP_Py, Piece(DTP_Prv_Piece_No).Width, Piece(DTP_Prv_Piece_No).Height, Form1.BMP.hDC, 408, 240, SRCCOPY
      BitBlt Form1.BMP.hDC, 408, 240, Piece(DTP_Piece_No).Width, Piece(DTP_Piece_No).Height, Form1.Picture1.hDC, DTP_x, DTP_y, SRCCOPY
      Copied_To_BMP = True
      BitBlt Form1.Picture1.hDC, DTP_x, DTP_y, Piece(DTP_Piece_No).Width, Piece(DTP_Piece_No).Height, Form1.BMP.hDC, Piece(DTP_Piece_No).MaskOnBmp_x, Piece(DTP_Piece_No).MaskOnBmp_y, SRCAND
      BitBlt Form1.Picture1.hDC, DTP_x, DTP_y, Piece(DTP_Piece_No).Width, Piece(DTP_Piece_No).Height, Form1.BMP.hDC, Piece(DTP_Piece_No).OnBmp_x, Piece(DTP_Piece_No).OnBmp_y, SRCPAINT
      Form1.Picture1.Refresh
End Sub

Public Function Get_A_Piece() As Long
   Get_A_Piece = 1 + Int(Rnd * 19)
End Function

Sub Check_The_Permissions(CTP_Piece_No As Long, CTP_x As Single, CTP_y As Single, CTP_Vx As Single, CTP_Vy As Single)
Dim CTP_x_Pos
Dim CTP_y_Pos
CTP_x_Pos = Int(CTP_x / 24) + 1
CTP_y_Pos = Int((CTP_y + CTP_Vy) / 24) + 1
Vertical_Stop_Status = False
Right_Stop_Status = False
Left_Stop_Status = False
If Int((CTP_y + CTP_Vy) / 24) > Int(CTP_y / 24) Or Int(CTP_y) Mod 24 = 0 Then
   Vertical_Stop_Status = False
   If Int(CTP_x) Mod 24 = 0 Then
      If Piece_No_Permission.Vertical(CTP_y_Pos, CTP_x_Pos) = True Then
         y = Int((CTP_y + CTP_Vy) / 24) * 24: Vy = 0: Vertical_Stop_Status = True
         If Vertical_Stop_Time = 0 Then Vertical_Stop_Time = Timer
         If Vertical_Stop_Time > 0 And Timer - Vertical_Stop_Time > 0.5 Then Stop_The_Piece CTP_Piece_No, CTP_x, CTP_y: Exit Sub
      End If
   Else
      If Piece_No_Permission.Vertical(CTP_y_Pos, CTP_x_Pos) = True Or Piece_No_Permission.Vertical2(CTP_y_Pos, CTP_x_Pos) = True Then
         y = Int((CTP_y + CTP_Vy) / 24) * 24: Vy = 0: Vertical_Stop_Status = True
         If Vertical_Stop_Time > 0 And Timer - Vertical_Stop_Time > 0.5 Then Stop_The_Piece CTP_Piece_No, CTP_x, CTP_y: Exit Sub
         If Vertical_Stop_Time = 0 Then Vertical_Stop_Time = Timer
      End If
   End If
   If Vertical_Stop_Status = False And Vy = 0 Then Vy = Vy_Level
End If
If Int((CTP_x + CTP_Vx) / 24) < Int(CTP_x / 24) Or Int(CTP_x) Mod 24 = 0 Then
   Left_Stop_Status = False
   If Int(CTP_y) Mod 24 = 0 Then
      If Piece_No_Permission.Left(CTP_y_Pos, CTP_x_Pos - 1) = True Then
         x = Int(CTP_x / 24) * 24: Vx = 0: Left_Stop_Status = True
      End If
   Else
      If Piece_No_Permission.Left(CTP_y_Pos, CTP_x_Pos - 1) = True Or Piece_No_Permission.Left2(CTP_y_Pos, CTP_x_Pos - 1) = True Then
         x = Int(CTP_x / 24) * 24: Vx = 0: Left_Stop_Status = True
      End If
   End If
End If
If Int((CTP_x + CTP_Vx) / 24) > Int(CTP_x / 24) Or Int(CTP_x) Mod 24 = 0 Then
   Right_Stop_Status = False
   If Int(y) Mod 24 = 0 Then
      If Piece_No_Permission.Right(CTP_y_Pos, CTP_x_Pos) = True Then
         x = Int(CTP_x / 24) * 24: Vx = 0: Right_Stop_Status = True
      End If
   Else
      If Piece_No_Permission.Right(CTP_y_Pos, CTP_x_Pos) = True Or Piece_No_Permission.Right2(CTP_y_Pos, CTP_x_Pos) = True Then
         x = Int(CTP_x / 24) * 24: Vx = 0: Right_Stop_Status = True
      End If
   End If
End If
If Prv_Right_Stop_Status = True And Right_Stop_Status = False And Right_Move_Requested = True Then
   x = x + 1: Vx = 0.1: Exit Sub
End If
If Vertical_Stop_Status = False Then Vertical_Stop_Time = 0
If Prv_Vertical_Stop_Status = True And Vertical_Stop_Status = False Then
   y = y + 1: Vy = Vy_Level: Exit Sub
End If
If Right_Move_Requested = True And Left_Stop_Status = True And Right_Stop_Status = False Then Vx = 0.1: x = x + 1
If Left_Move_Requested = True And Right_Stop_Status = True And Left_Stop_Status = False Then Vx = -0.1: x = x - 1
End Sub

Sub Create_Permission_Database(CPD_Piece_no As Long)
Dim CPD_1 As Long
Dim CPD_2 As Long
Dim CPD_3 As Long
Dim CPD_4 As Long
Dim CPD_5 As Long
For CPD_1 = 1 To 18
   For CPD_2 = 0 To 11
      Piece_No_Permission.Vertical(CPD_1, CPD_2) = False
      Piece_No_Permission.Vertical2(CPD_1, CPD_2) = False
      Piece_No_Permission.Left(CPD_1, CPD_2) = False
      Piece_No_Permission.Left2(CPD_1, CPD_2) = False
      Piece_No_Permission.Right(CPD_1, CPD_2) = False
      Piece_No_Permission.Right2(CPD_1, CPD_2) = False
      If Position_Empty(CPD_1, CPD_2) = False Then
         For CPD_3 = 1 To Piece(CPD_Piece_no).Number_Of_Positions_To_Check_Down
            CPD_4 = CPD_1 - Piece(CPD_Piece_no).Check_These_Positions_Down_y(CPD_3)
            CPD_5 = CPD_2 - Piece(CPD_Piece_no).Check_These_Positions_Down_x(CPD_3)
            If CPD_2 <> 0 And CPD_2 <> 11 Then
               If CPD_4 > 0 And CPD_5 > 0 And CPD_5 < 11 Then
                  Piece_No_Permission.Vertical(CPD_4, CPD_5) = True
               End If
               If CPD_4 > 0 And CPD_5 - 1 > 0 And CPD_5 - 1 < 11 Then
                  Piece_No_Permission.Vertical2(CPD_4, CPD_5 - 1) = True
               End If
            End If
         Next CPD_3
         For CPD_3 = 1 To Piece(CPD_Piece_no).Number_Of_Positions_To_Check_Left
            CPD_4 = CPD_1 - Piece(CPD_Piece_no).Check_These_Positions_Left_y(CPD_3)
            CPD_5 = CPD_2 - Piece(CPD_Piece_no).Check_These_Positions_Left_x(CPD_3)
            If CPD_1 <> 18 Then
               If CPD_4 > 0 And CPD_5 - 1 > 0 And CPD_5 - 1 < 11 Then
                  Piece_No_Permission.Left(CPD_4, CPD_5 - 1) = True
               End If
               If CPD_4 - 1 > 0 And CPD_5 - 1 > 0 And CPD_5 - 1 < 11 Then
                  Piece_No_Permission.Left2(CPD_4 - 1, CPD_5 - 1) = True
               End If
            End If
         Next CPD_3
         For CPD_3 = 1 To Piece(CPD_Piece_no).Number_Of_Positions_To_Check_Right
            CPD_4 = CPD_1 - Piece(CPD_Piece_no).Check_These_Positions_Right_y(CPD_3)
            CPD_5 = CPD_2 - Piece(CPD_Piece_no).Check_These_Positions_Right_x(CPD_3)
            If CPD_1 <> 18 Then
               If CPD_4 > 0 And CPD_5 > 0 And CPD_5 < 11 Then
                  Piece_No_Permission.Right(CPD_4, CPD_5) = True
               End If
               If CPD_4 - 1 > 0 And CPD_5 > 0 And CPD_5 < 11 Then
                  Piece_No_Permission.Right2(CPD_4 - 1, CPD_5) = True
               End If
            End If
         Next CPD_3
      End If
   Next CPD_2
Next CPD_1
End Sub

Sub Stop_The_Piece(STP_Piece_No As Long, STP_x As Single, STP_y As Single)
Vertical_Stop_Time = 0
Dim STP_x_Pos
Dim STP_y_Pos
Dim STP_i As Long
Dim STP_Tmp1 As Long
Dim STP_Tmp2 As Long
STP_x_Pos = Int(STP_x / 24) + 1
STP_y_Pos = Int((STP_y) / 24) + 1
Check_The_Permissions Piece_No, (STP_x_Pos - 1) * 24, y, 0, 0
If Vertical_Stop_Status = False Then
   STP_x_Pos = STP_x_Pos + 1
Else
   If x - (STP_x_Pos - 1) * 24 > 12 Then
      Check_The_Permissions Piece_No, (STP_x_Pos) * 24, y, 0, 0
      If Vertical_Stop_Status = True Then
         STP_x_Pos = STP_x_Pos + 1
      End If
   End If
End If
x = (STP_x_Pos - 1) * 24
y = (STP_y_Pos - 1) * 24
Piece_Stopped = True
For STP_i = 1 To Piece(STP_Piece_No).Number_Of_Squares
   STP_Tmp1 = STP_y_Pos + Piece(STP_Piece_No).Piece_Is_This_y(STP_i)
   STP_Tmp2 = STP_x_Pos + Piece(STP_Piece_No).Piece_Is_This_x(STP_i)
   Position_Empty(STP_Tmp1, STP_Tmp2) = False
Next STP_i
End Sub

Public Function Turn_Permission(TP_Piece_No As Long, TP_x As Single, TP_Y As Single, TP_Current As Boolean) As Boolean
Dim TP_New_Piece_No As Long
Dim TP_New_x As Single
Dim TP_New_y As Single
Dim TP_x_Pos As Long
Dim TP_y_Pos As Long
Dim TP_i As Long
TP_New_Piece_No = Piece(TP_Piece_No).After_Turn_Piece_No
TP_New_x = TP_x + Piece(TP_Piece_No).After_Turn_dx
TP_New_y = TP_Y + Piece(TP_Piece_No).After_Turn_dy
If TP_Current = True Then
   TP_New_Piece_No = TP_Piece_No
   TP_New_x = Piece(TP_New_Piece_No).Creating_x
   TP_New_y = 1
End If
TP_x_Pos = Int(TP_New_x / 24) + 1
TP_y_Pos = Int(TP_New_y / 24) + 1
If TP_New_x < 0 Or TP_New_x + Piece(TP_New_Piece_No).Width > 240 Then Turn_Permission = False: Exit Function
If TP_New_y < 0 Or TP_New_y + Piece(TP_New_Piece_No).Height > 384 Then Turn_Permission = False: Exit Function
For TP_i = 1 To Piece(TP_New_Piece_No).Number_Of_Squares
   If Int(TP_New_x) Mod 24 = 0 Then
      If Int(TP_New_y) Mod 24 = 0 Then
         If Position_Empty(TP_y_Pos + Piece(TP_New_Piece_No).Piece_Is_This_y(TP_i), TP_x_Pos + Piece(TP_New_Piece_No).Piece_Is_This_x(TP_i)) = False Then Turn_Permission = False: Exit Function
      Else
         If Position_Empty(TP_y_Pos + Piece(TP_New_Piece_No).Piece_Is_This_y(TP_i), TP_x_Pos + Piece(TP_New_Piece_No).Piece_Is_This_x(TP_i)) = False Then Turn_Permission = False: Exit Function
         If Position_Empty(TP_y_Pos + Piece(TP_New_Piece_No).Piece_Is_This_y(TP_i) + 1, TP_x_Pos + Piece(TP_New_Piece_No).Piece_Is_This_x(TP_i)) = False Then Turn_Permission = False: Exit Function
      End If
   Else
      If Int(TP_New_y) Mod 24 = 0 Then
         If Position_Empty(TP_y_Pos + Piece(TP_New_Piece_No).Piece_Is_This_y(TP_i), TP_x_Pos + Piece(TP_New_Piece_No).Piece_Is_This_x(TP_i)) = False Then Turn_Permission = False: Exit Function
         If Position_Empty(TP_y_Pos + Piece(TP_New_Piece_No).Piece_Is_This_y(TP_i), TP_x_Pos + Piece(TP_New_Piece_No).Piece_Is_This_x(TP_i) + 1) = False Then Turn_Permission = False: Exit Function
      Else
         If Position_Empty(TP_y_Pos + Piece(TP_New_Piece_No).Piece_Is_This_y(TP_i), TP_x_Pos + Piece(TP_New_Piece_No).Piece_Is_This_x(TP_i)) = False Or Position_Empty(TP_y_Pos + Piece(TP_New_Piece_No).Piece_Is_This_y(TP_i) + 1, TP_x_Pos + Piece(TP_New_Piece_No).Piece_Is_This_x(TP_i)) = False Then Turn_Permission = False: Exit Function
         If Position_Empty(TP_y_Pos + Piece(TP_New_Piece_No).Piece_Is_This_y(TP_i) + 1, TP_x_Pos + Piece(TP_New_Piece_No).Piece_Is_This_x(TP_i)) = False Or Position_Empty(TP_y_Pos + Piece(TP_New_Piece_No).Piece_Is_This_y(TP_i) + 1, TP_x_Pos + Piece(TP_New_Piece_No).Piece_Is_This_x(TP_i)) = False Then Turn_Permission = False: Exit Function
         If Position_Empty(TP_y_Pos + Piece(TP_New_Piece_No).Piece_Is_This_y(TP_i), TP_x_Pos + Piece(TP_New_Piece_No).Piece_Is_This_x(TP_i)) = False Or Position_Empty(TP_y_Pos + Piece(TP_New_Piece_No).Piece_Is_This_y(TP_i) + 1, TP_x_Pos + Piece(TP_New_Piece_No).Piece_Is_This_x(TP_i) + 1) = False Then Turn_Permission = False: Exit Function
      End If
   End If
Next TP_i
Turn_Permission = True
End Function

Sub Check_Rows_After_Stop()
Dim CRAS_i As Long
Dim CRAS_j As Long
Dim CRAS_k As Long
Dim CRAS_Rows_To_Delete(1 To 4) As Long
Dim CRAS_Tmp As Long
CRAS_Tmp = 1
For CRAS_i = 17 To 1 Step -1
   For CRAS_j = 1 To 10
      If Position_Empty(CRAS_i, CRAS_j) = True Then GoTo At_Least_One_Position_Is_Empty
   Next CRAS_j
   CRAS_Rows_To_Delete(CRAS_Tmp) = CRAS_i
   CRAS_Tmp = CRAS_Tmp + 1
   If CRAS_Tmp = 5 Then Exit For
At_Least_One_Position_Is_Empty:
Next CRAS_i
CRAS_Rows_To_Delete(2) = CRAS_Rows_To_Delete(2) + 1
CRAS_Rows_To_Delete(3) = CRAS_Rows_To_Delete(3) + 2
CRAS_Rows_To_Delete(4) = CRAS_Rows_To_Delete(4) + 3
If CRAS_Tmp > 1 Then
   Score = Score + 100 * (2 ^ (CRAS_Tmp - 1)) * Level
   For CRAS_i = 1 To CRAS_Tmp - 1
      Form1.Picture1.Line (0, (CRAS_Rows_To_Delete(CRAS_i) - 1) * 24)-(244, (CRAS_Rows_To_Delete(CRAS_i) - 1) * 24 + 23), RGB(0, 0, 0), BF
      BitBlt Form1.Picture1.hDC, 0, 24, Form1.Picture1.Width, (CRAS_Rows_To_Delete(CRAS_i) - 1) * 24, Form1.Picture1.hDC, 0, 0, SRCCOPY
      Form1.Picture1.Line (0, 0)-(244, 25), RGB(0, 0, 0), BF
      Form1.Picture1.Line (0, 0)-(244, 0), RGB(64, 64, 64)
      Form1.Picture1.Line (0, 24)-(244, 24), RGB(64, 64, 64)
      For CRAS_j = 240 To 0 Step -24
         Form1.Picture1.Line (CRAS_j, 0)-(CRAS_j, 24), RGB(64, 64, 64)
      Next CRAS_j
      For CRAS_j = 240 To 0 Step -24
         For CRAS_k = 0 To 24 Step 24
            Form1.Picture1.PSet (CRAS_j, CRAS_k), RGB(96, 96, 96)
         Next CRAS_k
      Next CRAS_j
      Form1.Picture1.Refresh
      For CRAS_j = CRAS_Rows_To_Delete(CRAS_i) To 2 Step -1
         For CRAS_k = 1 To 10
            Position_Empty(CRAS_j, CRAS_k) = Position_Empty(CRAS_j - 1, CRAS_k)
         Next CRAS_k
      Next CRAS_j
      For CRAS_k = 1 To 10
         Position_Empty(1, CRAS_k) = True
      Next CRAS_k
   Next CRAS_i
End If
End Sub

