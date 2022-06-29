Attribute VB_Name = "progress"

Sub PercentBar(Shape As Control, Done As Integer, Total As Variant)

'Call PercentBar(Picture1, Label1.Caption, Label2.Caption)

On Error Resume Next
Shape.AutoRedraw = True
Shape.FillStyle = 0
Shape.DrawStyle = 0
Shape.FontName = "MS Sans Serif"
Shape.FontSize = 8.25
Shape.FontBold = False
X = Done / Total * Shape.Width
Shape.Line (0, 0)-(Shape.Width, Shape.Height), RGB(192, 192, 192), BF
Shape.Line (0, 0)-(X - 10, Shape.Height), RGB(64, 128, 128), BF
Shape.CurrentX = (Shape.Width / 2) - 100
Shape.CurrentY = (Shape.Height / 2) - 125
Shape.ForeColor = RGB(192, 192, 192)
'Shape.Print Percent(Done, Total, 100) & "%"
End Sub

Function Percent(Complete As Integer, Total As Variant, TotalOutput As Integer) As Integer
On Error Resume Next
Percent = Int(Complete / Total * TotalOutput)
End Function
