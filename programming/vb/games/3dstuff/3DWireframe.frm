VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   ScaleHeight     =   9360
   ScaleWidth      =   9645
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   6015
      Left            =   8640
      Max             =   2
      Min             =   20
      TabIndex        =   11
      Top             =   120
      Value           =   10
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Corny Spaceship"
      Height          =   375
      Left            =   6720
      TabIndex        =   10
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cube"
      Height          =   375
      Left            =   6720
      TabIndex        =   9
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6120
      Top             =   7200
   End
   Begin VB.CheckBox Going 
      Caption         =   "Rotate"
      Height          =   255
      Left            =   4920
      TabIndex        =   8
      Top             =   6360
      Width           =   1455
   End
   Begin VB.TextBox txtz 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Text            =   "0"
      Top             =   7080
      Width           =   1095
   End
   Begin VB.TextBox txty 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Text            =   "0"
      Top             =   6720
      Width           =   1095
   End
   Begin VB.TextBox txtx 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Text            =   "0"
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Build Rotation Matrix"
      Height          =   615
      Left            =   2880
      TabIndex        =   1
      Top             =   6360
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   6015
      Left            =   0
      ScaleHeight     =   397
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   565
      TabIndex        =   0
      Top             =   120
      Width           =   8535
   End
   Begin VB.Label Label4 
      Caption         =   "10"
      Height          =   255
      Left            =   8640
      TabIndex        =   12
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "z:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "y:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "x:"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   6360
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TDLine 'this type holds a three dimensional line
x As Double
y As Double
z As Double
x2 As Double
y2 As Double
z2 As Double
End Type

Private Type TDPoint 'this type holds a three dimensional point
x As Double
y As Double
z As Double
End Type

'Private Declare Function GetPixel Lib "GDI32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
'Private Declare Function SetPixel Lib "GDI32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal color As Long) As Long
Dim RM(0 To 3, 0 To 3) As Double 'rotation matrix (we multiply by this to modify the point)
Dim PLine(0 To 29) As TDLine 'all the lines in a given object
Dim NumLines As Integer 'how many lines to draw
Dim Zoom As Integer 'zoom level

Private Sub Command1_Click()
Picture2.Cls 'clear the picture

Dim x As Double
Dim y As Double
Dim z As Double
Const pi = 3.14159265358979 'to convert to radians

x = txtx.Text * pi / 180 'convert each angle to radians
y = txty.Text * pi / 180
z = txtz.Text * pi / 180
MatrixBuild x, y, z 'build the matrix
Timer1_Timer 'run the timer (whether it's enabled or not) to draw the current frame

End Sub

Private Sub SquareLine()
PLine(0).x = -64 'fill in the array of lines for the cube
PLine(0).y = -64
PLine(0).z = -64
PLine(0).x2 = 64
PLine(0).y2 = -64
PLine(0).z2 = -64

PLine(1).x = 64
PLine(1).y = -64
PLine(1).z = -64
PLine(1).x2 = 64
PLine(1).y2 = 64
PLine(1).z2 = -64

PLine(2).x = 64
PLine(2).y = 64
PLine(2).z = -64
PLine(2).x2 = -64
PLine(2).y2 = 64
PLine(2).z2 = -64

PLine(3).x = -64
PLine(3).y = 64
PLine(3).z = -64
PLine(3).x2 = -64
PLine(3).y2 = -64
PLine(3).z2 = -64

PLine(4).x = -64
PLine(4).y = -64
PLine(4).z = 64
PLine(4).x2 = 64
PLine(4).y2 = -64
PLine(4).z2 = 64

PLine(5).x = 64
PLine(5).y = -64
PLine(5).z = 64
PLine(5).x2 = 64
PLine(5).y2 = 64
PLine(5).z2 = 64

PLine(6).x = 64
PLine(6).y = 64
PLine(6).z = 64
PLine(6).x2 = -64
PLine(6).y2 = 64
PLine(6).z2 = 64

PLine(7).x = -64
PLine(7).y = 64
PLine(7).z = 64
PLine(7).x2 = -64
PLine(7).y2 = -64
PLine(7).z2 = 64

PLine(8).x = -64
PLine(8).y = -64
PLine(8).z = -64
PLine(8).x2 = -64
PLine(8).y2 = -64
PLine(8).z2 = 64

PLine(9).x = 64
PLine(9).y = -64
PLine(9).z = -64
PLine(9).x2 = 64
PLine(9).y2 = -64
PLine(9).z2 = 64

PLine(10).x = -64
PLine(10).y = 64
PLine(10).z = -64
PLine(10).x2 = -64
PLine(10).y2 = 64
PLine(10).z2 = 64

PLine(11).x = 64
PLine(11).y = 64
PLine(11).z = -64
PLine(11).x2 = 64
PLine(11).y2 = 64
PLine(11).z2 = 64

PLine(12).x = 64
PLine(12).y = 64
PLine(12).z = 64
PLine(12).x2 = -64
PLine(12).y2 = -64
PLine(12).z2 = -64

PLine(13).x = 64
PLine(13).y = -64
PLine(13).z = 64
PLine(13).x2 = -64
PLine(13).y2 = 64
PLine(13).z2 = -64

NumLines = 14
End Sub

Private Sub ShipLine()
PLine(0).x = -40 'fill in the array of lines for the corny spaceship
PLine(0).y = -40
PLine(0).z = 0
PLine(0).x2 = 40
PLine(0).y2 = -40
PLine(0).z2 = 0

PLine(1).x = -40
PLine(1).y = -40
PLine(1).z = 0
PLine(1).x2 = 0
PLine(1).y2 = 0
PLine(1).z2 = 0

PLine(2).x = 40
PLine(2).y = -40
PLine(2).z = 0
PLine(2).x2 = 0
PLine(2).y2 = 0
PLine(2).z2 = 0

PLine(3).x = -40
PLine(3).y = -40
PLine(3).z = 0
PLine(3).x2 = 0
PLine(3).y2 = -40
PLine(3).z2 = 80

PLine(4).x = 0
PLine(4).y = 0
PLine(4).z = 0
PLine(4).x2 = 0
PLine(4).y2 = -40
PLine(4).z2 = 80

PLine(5).x = 40
PLine(5).y = -40
PLine(5).z = 0
PLine(5).x2 = 0
PLine(5).y2 = -40
PLine(5).z2 = 80

PLine(6).x = 0
PLine(6).y = 12
PLine(6).z = -12
PLine(6).x2 = 0
PLine(6).y2 = -40
PLine(6).z2 = 80

PLine(7).x = -52
PLine(7).y = -52
PLine(7).z = -12
PLine(7).x2 = 0
PLine(7).y2 = -40
PLine(7).z2 = 80

PLine(8).x = 52
PLine(8).y = -52
PLine(8).z = -12
PLine(8).x2 = 0
PLine(8).y2 = -40
PLine(8).z2 = 80

PLine(9).x = 0
PLine(9).y = 12
PLine(9).z = -12
PLine(9).x2 = 0
PLine(9).y2 = -40
PLine(9).z2 = -12

PLine(10).x = -52
PLine(10).y = -52
PLine(10).z = -12
PLine(10).x2 = 0
PLine(10).y2 = -40
PLine(10).z2 = -12

PLine(11).x = 52
PLine(11).y = -52
PLine(11).z = -12
PLine(11).x2 = 0
PLine(11).y2 = -40
PLine(11).z2 = -12

NumLines = 12
End Sub

Private Sub MatrixBuild(ByVal x As Double, ByVal y As Double, ByVal z As Double)
' this sub builds the rotation matrix with x, y and z as axis angles
Dim SinX, CosX, SinY, CosY, SinZ, CosZ, C1, C2

SinX = Sin(x)
CosX = Cos(x)
SinY = Sin(y)
CosY = Cos(y)
SinZ = Sin(z)
CosZ = Cos(z)

RM(0, 0) = (CosZ * CosY)
RM(0, 1) = (CosZ * -SinY * -SinX + SinZ * CosX)
RM(0, 2) = (CosZ * -SinY * CosX + SinZ * SinX)
RM(1, 0) = (-SinZ * CosY)
RM(1, 1) = (-SinZ * -SinY * -SinX + CosZ * CosX)
RM(1, 2) = (-SinZ * -SinY * CosX + CosZ * SinX)
RM(2, 0) = SinY
RM(2, 1) = CosY * -SinX
RM(2, 2) = CosY * CosX
End Sub

Private Function RotatePoint(ByVal x As Double, ByVal y As Double, ByVal z As Double) As TDPoint
' finds new point using Rotate Matrix with x, y and z as current point positions
Dim TempPoint As TDPoint
TempPoint.x = (x * RM(0, 0)) + (y * RM(0, 1)) + (z * RM(0, 2)) + RM(0, 3)
TempPoint.y = (x * RM(1, 0)) + (y * RM(1, 1)) + (z * RM(1, 2)) + RM(1, 3)
TempPoint.z = (x * RM(2, 0)) + (y * RM(2, 1)) + (z * RM(2, 2)) + RM(2, 3)

RotatePoint = TempPoint
End Function

Private Sub Command2_Click()
SquareLine 'build the square object
End Sub

Private Sub Command3_Click()
ShipLine 'build the ship object
End Sub


Private Sub Form_Load()
Zoom = 10 'default zoom
End Sub

Private Sub Going_Click()
If Going.Value = 0 Then Timer1.Enabled = False
If Going.Value = 1 Then Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Dim GetPoint As TDPoint
Dim EndPoint As TDPoint
Dim TZoom As Single
TZoom = Zoom / 10
Picture2.Cls
For i = 0 To NumLines - 1 'for every single line do this
GetPoint = RotatePoint(PLine(i).x, PLine(i).y, PLine(i).z) 'get the first point of line
PLine(i).x = GetPoint.x
PLine(i).y = GetPoint.y
PLine(i).z = GetPoint.z
GetPoint = RotatePoint(PLine(i).x2, PLine(i).y2, PLine(i).z2) 'get last point of line
PLine(i).x2 = GetPoint.x
PLine(i).y2 = GetPoint.y
PLine(i).z2 = GetPoint.z
'draw new line (we can ignore z because the view is orthographic, not perspective)
Picture2.Line (PLine(i).x * TZoom + 256, PLine(i).y * TZoom + 192)-(PLine(i).x2 * TZoom + 256, PLine(i).y2 * TZoom + 192), RGB(255, 255, 255)
Next i
End Sub

Private Sub VScroll1_Change()
Label4.Caption = VScroll1.Value
Zoom = VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
Zoom = VScroll1.Value
Label4.Caption = VScroll1.Value
End Sub
