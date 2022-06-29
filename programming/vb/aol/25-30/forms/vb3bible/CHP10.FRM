VERSION 2.00
Begin Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shapes Project"
   ClientHeight    =   5865
   ClientLeft      =   705
   ClientTop       =   1005
   ClientWidth     =   5700
   Height          =   6300
   Icon            =   CHP10.FRX:0000
   Left            =   630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   5700
   Top             =   645
   Width           =   5850
   Begin CommandButton cmndClearScreen 
      Height          =   372
      Left            =   5040
      TabIndex        =   17
      Top             =   5244
      Width           =   372
   End
   Begin CommandButton cmndDrawTriangle 
      Caption         =   "&Triangle"
      Height          =   372
      Left            =   4080
      TabIndex        =   15
      Top             =   5244
      Width           =   852
   End
   Begin CommandButton cmndDrawBox 
      Caption         =   "&Box"
      Height          =   372
      Left            =   3120
      TabIndex        =   14
      Top             =   5244
      Width           =   852
   End
   Begin CommandButton cmndDrawSquare 
      Caption         =   "&Square"
      Height          =   372
      Left            =   2160
      TabIndex        =   12
      Top             =   5244
      Width           =   852
   End
   Begin CommandButton cmndDrawLine 
      Caption         =   "&Line"
      Height          =   372
      Left            =   1200
      TabIndex        =   13
      Top             =   5244
      Width           =   852
   End
   Begin CommandButton cmndDrawCircle 
      Caption         =   "&Circle"
      Height          =   372
      Left            =   240
      TabIndex        =   11
      Top             =   5244
      Width           =   852
   End
   Begin PictureBox pictSurface 
      Height          =   3660
      Left            =   120
      ScaleHeight     =   3630
      ScaleWidth      =   5355
      TabIndex        =   16
      Top             =   1200
      Width           =   5385
   End
   Begin ComboBox cmboPickPattern 
      Height          =   288
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   840
      Width           =   2652
   End
   Begin ComboBox cmboPickPen 
      Height          =   288
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   840
      Width           =   2652
   End
   Begin ComboBox cmboPickLine 
      Height          =   288
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   240
      Width           =   2652
   End
   Begin ComboBox cmboPickColor 
      Height          =   288
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   1812
   End
   Begin ComboBox cmboPickWidth 
      Height          =   288
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   732
   End
   Begin Line Line2 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   135
      X2              =   5517
      Y1              =   5130
      Y2              =   5130
   End
   Begin Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   135
      X2              =   135
      Y1              =   5145
      Y2              =   5720
   End
   Begin Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   615
      Left            =   135
      Top             =   5130
      Width           =   5415
   End
   Begin Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "&Pattern"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   252
      Left            =   2880
      TabIndex        =   10
      Top             =   600
      Width           =   732
   End
   Begin Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   3645
      Left            =   225
      Top             =   1320
      Width           =   5370
   End
   Begin Label Label5 
      Caption         =   "Label5"
      Height          =   12
      Left            =   3240
      TabIndex        =   9
      Top             =   840
      Width           =   732
   End
   Begin Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Lin&e"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   252
      Left            =   2880
      TabIndex        =   8
      Top             =   0
      Width           =   612
   End
   Begin Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "C&olor"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   252
      Left            =   960
      TabIndex        =   7
      Top             =   0
      Width           =   492
   End
   Begin Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Pe&n"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   252
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   612
   End
   Begin Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Width"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   252
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   492
   End
End

Sub cmboPickColor_Click ()
     pictSurface.FillColor = QBColor(cmboPickColor.ListIndex)
End Sub

Sub cmboPickLine_Click ()
     pictSurface.DrawStyle = cmboPickLine.ListIndex
End Sub

Sub cmboPickPattern_Click ()
     pictSurface.FillStyle = cmboPickPattern.ListIndex
End Sub

Sub cmboPickPen_Click ()
     pictSurface.DrawMode = cmboPickPen.ListIndex + 1
End Sub

Sub cmboPickWidth_Click ()
     pictSurface.DrawWidth = cmboPickWidth.ListIndex + 1
End Sub

Sub cmndClearScreen_Click ()
     pictSurface.Cls
End Sub

Sub cmndDrawBox_Click ()
     X1 = Int((75 - (1)) * Rnd + (1))        'Pick starting coordinates
     Y1 = Int((75 - (1)) * Rnd + (1))
     X2 = Int((50 - (1)) * Rnd + (1))        'Pick ending coordinates
     Y2 = Int((50 - (1)) * Rnd + (1))
     pictSurface.Line (X1, Y1)-Step(X2, Y2), , B    'Draw the box
End Sub

Sub cmndDrawCircle_Click ()
     X = Int((100 - 1) * Rnd + 1)            'Pick Center coordinates
     Y = Int((100 - 1) * Rnd + 1)
     R = Int((25 - 1) * Rnd + 1)             'Pick Radius
     pictSurface.Circle (X, Y), R            'Draw Circle
End Sub

Sub cmndDrawLine_Click ()
     X1 = Int((100 - 1) * Rnd + 1)           'Pick Starting Coordinates
     Y1 = Int((100 - 1) * Rnd + 1)
     X2 = Int((100 - 1) * Rnd + 1)           'Pick Ending Coordinates
     Y2 = Int((100 - 1) * Rnd + 1)
     pictSurface.Line (X1, Y1)-(X2, Y2)      'Draw Box
End Sub

Sub cmndDrawSquare_Click ()
     X1 = Int((50 - 1) * Rnd + 1)            'Pick Starting Coordinates
     Y1 = Int((50 - 1) * Rnd + 1)
     X2 = Int((50 - 1) * Rnd + 1)            'Pick Ending Coordinates
     Y2 = X2
     pictSurface.Line (X1, Y1)-Step(X2, Y2), , B    'Draw Square
End Sub

Sub cmndDrawTriangle_Click ()
     X1 = Int((100 - 1) * Rnd + 1)           'Pick one corner
     Y1 = Int((100 - 1) * Rnd + 1)
     X2 = Int((100 - 1) * Rnd + 1)           'Pick another corner
     Y2 = Int((100 - 1) * Rnd + 1)
     X3 = Int((100 - 1) * Rnd + 1)           'Pick third corner
     Y3 = Int((100 - 1) * Rnd + 1)
     pictSurface.Line (X1, Y1)-(X2, Y2)      'Draw Triangle
     pictSurface.Line (X2, Y2)-(X3, Y3)
     pictSurface.Line (X3, Y3)-(X1, Y1)
End Sub

Sub Form_Load ()
     For i = 1 To 10                           'Fill Width pick list with possible widths
          cmboPickWidth.AddItem Str$(i)
     Next i
     
     cmboPickColor.AddItem "Black"             'Fill Color pick list with possible colors
     cmboPickColor.AddItem "Blue"
     cmboPickColor.AddItem "Green"
     cmboPickColor.AddItem "Cyan"
     cmboPickColor.AddItem "Red"
     cmboPickColor.AddItem "Magenta"
     cmboPickColor.AddItem "Yellow"
     cmboPickColor.AddItem "White"
     cmboPickColor.AddItem "Gray"
     cmboPickColor.AddItem "Light Blue"
     cmboPickColor.AddItem "Light Green"
     cmboPickColor.AddItem "Light Cyan"
     cmboPickColor.AddItem "Light Red"
     cmboPickColor.AddItem "Light Magenta"
     cmboPickColor.AddItem "Light Yellow"
     cmboPickColor.AddItem "Bright White"
     
     cmboPickPattern.AddItem "Solid"            'Fill Pattern pick list with patterns
     cmboPickPattern.AddItem "Transparent"
     cmboPickPattern.AddItem "Horizontal Line"
     cmboPickPattern.AddItem "Vertical Line"
     cmboPickPattern.AddItem "Upward Diagonal"
     cmboPickPattern.AddItem "Downward Diagonal"
     cmboPickPattern.AddItem "Cross"
     cmboPickPattern.AddItem "Diagonal Cross"
     
     cmboPickPen.AddItem "Blackness"            'Fill Pen pick list with possible pens
     cmboPickPen.AddItem "Not Merge Pen"
     cmboPickPen.AddItem "Mask Not Pen"
     cmboPickPen.AddItem "Not Copy Pen"
     cmboPickPen.AddItem "Mask Pen Not"
     cmboPickPen.AddItem "Invert"
     cmboPickPen.AddItem "Xor Pen"
     cmboPickPen.AddItem "Not Mask Pen"
     cmboPickPen.AddItem "Mask Pen"
     cmboPickPen.AddItem "Not Xor Pen"
     cmboPickPen.AddItem "Nop"
     cmboPickPen.AddItem "Merge Not Pen"
     cmboPickPen.AddItem "Copy Pen"
     cmboPickPen.AddItem "Merge Pen Not"
     cmboPickPen.AddItem "Merge Pen"
     cmboPickPen.AddItem "Whiteness"
     
     cmboPickLine.AddItem "Solid"               'Fill Line pick list with possible pens
     cmboPickLine.AddItem "Dash"
     cmboPickLine.AddItem "Dot"
     cmboPickLine.AddItem "Dash-Dot"
     cmboPickLine.AddItem "Dash-Dot-Dot"
     cmboPickLine.AddItem "Invisible"
     cmboPickLine.AddItem "Inside Solid"
     
     cmboPickWidth.ListIndex = 0                '1    (These are all set to defaults)
     cmboPickColor.ListIndex = 0                'Black
     cmboPickPattern.ListIndex = 0              'Solid
     cmboPickPen.ListIndex = 13                 'Copy Pen
     cmboPickLine.ListIndex = 0                 'Solid

     pictSurface.ScaleHeight = 100              'Set up an easy custom scale
     pictSurface.ScaleWidth = 100
     pictSurface.AutoRedraw = True              'Make sure graphics redrawn immediately
End Sub

Sub pictSurface_DblClick ()
     pictSurface.Cls
End Sub

