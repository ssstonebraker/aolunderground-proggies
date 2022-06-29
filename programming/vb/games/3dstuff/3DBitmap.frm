VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".bmp"
      DialogTitle     =   "Open a bitmap"
      Filter          =   "Bitmap Files (*.bmp)|*.bmp"
   End
   Begin VB.CheckBox Perspect 
      Caption         =   "Perspective"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5760
      TabIndex        =   11
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CheckBox Despeckle 
      Caption         =   "Despeckle"
      Height          =   255
      Left            =   5760
      TabIndex        =   10
      Top             =   3840
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox Shade 
      Caption         =   "Z Shading"
      Height          =   255
      Left            =   5760
      TabIndex        =   9
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox txtz 
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Text            =   "0"
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox txty 
      Height          =   285
      Left            =   2400
      TabIndex        =   4
      Text            =   "0"
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txtx 
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Text            =   "0"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Draw"
      Height          =   615
      Left            =   3720
      TabIndex        =   2
      Top             =   3480
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   3255
      Left            =   3720
      ScaleHeight     =   213
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   229
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   120
      Picture         =   "3DBITMAP.frx":0000
      ScaleHeight     =   213
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   221
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "z:"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "y:"
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "x:"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   3480
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TDPoint 'this type holds a single point
x As Integer
y As Integer
z As Integer
End Type

'these are faster than pset and point
Private Declare Function GetPixel Lib "GDI32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "GDI32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal color As Long) As Long
Dim RM(0 To 3, 0 To 3) As Double 'rotation matrix
Dim PPoint(0 To 16383) As TDPoint 'all the points

Private Sub Command1_Click()
Picture2.Cls
Dim GetPoint As TDPoint
Dim color As Long
Dim cred, cgreen, cblue
Dim x As Double
Dim y As Double
Dim z As Double
Const pi = 3.14159265358979
'PointArray
x = txtx.Text * pi / 180 'convert to radians
y = txty.Text * pi / 180
z = txtz.Text * pi / 180
MatrixBuild x, y, z 'build the matrix
For y = -64 To 63 'for every single point
For x = -64 To 63
GetPoint = RotatePoint(x, y, 0) 'rotate point

color = GetPixel(Picture1.hDC, x + 64, y + 64) 'get the color of it
If Shade.Value = 1 Then 'shade depending on z distance
    cred = Int(color Mod 256)
    cblue = Int(color / 65536)
    cgreen = Int((color - (cblue * 65536) - cred) / 256)
    cred = cred + GetPoint.z * 3
    cgreen = cgreen + GetPoint.z * 3
    cblue = cblue + GetPoint.z * 3
    If cred > 255 Then cred = 255
    If cred < 0 Then cred = 0
    If cgreen > 255 Then cgreen = 255
    If cgreen < 0 Then cgreen = 0
    If cblue > 255 Then cblue = 255
    If cblue < 0 Then cblue = 0
    color = RGB(cred, cgreen, cblue)
End If



SetPixel Picture2.hDC, GetPoint.x + 64, GetPoint.y + 64, color 'set the new point
Next x
Next y

If Despeckle.Value = 1 Then
For y = 1 To 160 'despeckle
For x = 1 To 160
'for black each pixel, this sub finds the four pixels around it and takes an average
'of them for a new color
If (GetPixel(Picture2.hDC, x, y) = RGB(0, 0, 0)) Then
    fcolor = Picture2.Point(x - 1, y)
    cred = Int(fcolor Mod 256)
    cblue = Int(fcolor / 65536)
    cgreen = Int((fcolor - (cblue * 65536) - cred) / 256)
    
    fcolor = GetPixel(Picture2.hDC, x, y - 1)
    CRed2 = Int(fcolor Mod 256)
    CBlue2 = Int(fcolor / 65536)
    CGreen2 = Int((fcolor - (CBlue2 * 65536) - CRed2) / 256)
    
    fcolor = GetPixel(Picture2.hDC, x + 1, y)
    CRed3 = Int(fcolor Mod 256)
    CBlue3 = Int(fcolor / 65536)
    CGreen3 = Int((fcolor - (CBlue3 * 65536) - CRed3) / 256)
    
    fcolor = GetPixel(Picture2.hDC, x, y + 1)
    CRed4 = Int(fcolor Mod 256)
    CBlue4 = Int(fcolor / 65536)
    CGreen4 = Int((fcolor - (CBlue4 * 65536) - CRed4) / 256)
    
    SetPixel Picture2.hDC, x, y, RGB((cred + CRed2 + CRed3 + CRed4) / 4, _
        (cgreen + CGreen2 + CGreen3 + CGreen4) / 4, _
        (cblue + CBlue2 + CBlue3 + CBlue4) / 4)
End If

Next x
Next y
End If
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

Private Function RotatePoint(ByVal x As Integer, ByVal y As Integer, ByVal z As Integer) As TDPoint
' finds new point using Rotate Matrix with x, y and z as current point positions
Dim TempPoint As TDPoint
TempPoint.x = (x * RM(0, 0)) + (y * RM(0, 1)) + (z * RM(0, 2)) + RM(0, 3)
TempPoint.y = (x * RM(1, 0)) + (y * RM(1, 1)) + (z * RM(1, 2)) + RM(1, 3)
TempPoint.z = (x * RM(2, 0)) + (y * RM(2, 1)) + (z * RM(2, 2)) + RM(2, 3)

RotatePoint = TempPoint
End Function

Private Sub Picture1_Click()
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Then Exit Sub
Picture1.Picture = LoadPicture(CommonDialog1.FileName)

End Sub
