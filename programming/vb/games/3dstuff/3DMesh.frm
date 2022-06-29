VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "3D Mesh Rotation from a bitmap"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6960
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".bmp"
      DialogTitle     =   "open a bitmap"
      Filter          =   "bitmaps|*.bmp"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Picture..."
      Height          =   255
      Left            =   6360
      TabIndex        =   9
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox txtz 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Text            =   "0"
      Top             =   7080
      Width           =   1095
   End
   Begin VB.TextBox txty 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Text            =   "0"
      Top             =   6720
      Width           =   1095
   End
   Begin VB.TextBox txtx 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Text            =   "0"
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Draw"
      Height          =   615
      Left            =   2640
      TabIndex        =   2
      Top             =   6360
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   6015
      Left            =   120
      ScaleHeight     =   397
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   405
      TabIndex        =   1
      Top             =   120
      Width           =   6135
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   6360
      Picture         =   "3DMesh.frx":0000
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   133
      TabIndex        =   0
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "z:"
      Height          =   255
      Left            =   -120
      TabIndex        =   8
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "y:"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "x:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   6360
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TDPoint
x As Double
y As Double
z As Double
End Type

Private Declare Function GetPixel Lib "GDI32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "GDI32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal color As Long) As Long
Dim RM(0 To 3, 0 To 3) As Double
Dim PPoint(0 To 127, 0 To 127) As TDPoint

Private Sub Command1_Click()
Picture2.Cls
Dim GetPoint As TDPoint
Dim color As Long
Dim CBump As Double
Dim XBump As Double
Dim zoom As Integer
Dim x As Double
Dim y As Double
Dim z As Double
Const pi = 3.14159265358979
'PointArray
zoom = 4
x = txtx.Text * pi / 180
y = txty.Text * pi / 180
z = txtz.Text * pi / 180
MatrixBuild x, y, z
For y = -32 To 31
For x = -32 To 31
color = GetPixel(Picture1.hDC, x + 32, y + 32) 'get color/height
XBump = CBump
CBump = -(Int(color Mod 256) / 5) 'convert to bump
PPoint(x + 32, y + 32) = RotatePoint(x, y, CBump) 'rotate point
If (x > -31) And (y > -31) Then
Picture2.Line (PPoint(x + 32, y + 32).x * zoom + 196, PPoint(x + 32, y + 32).y * zoom + 196)-(PPoint(x + 32, y + 31).x * zoom + 196, PPoint(x + 32, y + 31).y * zoom + 196), RGB(CBump * -5, CBump * -5, CBump * -5)
Picture2.Line (PPoint(x + 32, y + 32).x * zoom + 196, PPoint(x + 32, y + 32).y * zoom + 196)-(PPoint(x + 31, y + 32).x * zoom + 196, PPoint(x + 31, y + 32).y * zoom + 196), RGB(CBump * -5, CBump * -5, CBump * -5)
'Picture2.Line (PPoint(x + 32, y + 32).x * zoom + 196, PPoint(x + 32, y + 32).y * zoom + 196)-(PPoint(x + 32, y + 31).x * zoom + 196, PPoint(x + 32, y + 31).y * zoom + 196), RGB(0, XBump - CBump + 128, 0)
'Picture2.Line (PPoint(x + 32, y + 32).x * zoom + 196, PPoint(x + 32, y + 32).y * zoom + 196)-(PPoint(x + 31, y + 32).x * zoom + 196, PPoint(x + 31, y + 32).y * zoom + 196), RGB(0, XBump - CBump + 128, 0)

End If
Next x
Next y

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

'For C1 = 0 To 3 'identity matrix
'For C2 = 0 To 3
'If C1 = C2 Then
'    RM(C1, C2) = 0
'Else
'    RM(C1, C2) = 1
'End If
'Next C2
'Next C1

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

Private Function RotatePoint(ByVal x As Integer, ByVal y As Integer, ByVal z As Integer) As TDPoint
' finds new point using Rotate Matrix with x, y and z as current point positions
Dim TempPoint As TDPoint
TempPoint.x = (x * RM(0, 0)) + (y * RM(0, 1)) + (z * RM(0, 2)) + RM(0, 3)
TempPoint.y = (x * RM(1, 0)) + (y * RM(1, 1)) + (z * RM(1, 2)) + RM(1, 3)
TempPoint.z = (x * RM(2, 0)) + (y * RM(2, 1)) + (z * RM(2, 2)) + RM(2, 3)

RotatePoint = TempPoint
End Function

Private Sub Command2_Click()
CommonDialog1.ShowOpen
Picture1.Picture = LoadPicture(CommonDialog1.FileName)
End Sub

