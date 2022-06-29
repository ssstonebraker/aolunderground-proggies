Attribute VB_Name = "Module1"
Public en As Boolean
Public thrd As clsThreads


Public Sub Create()
Set thrd = New clsThreads
Call thrd.Initialize(AddressOf Clock)
thrd.Enabled = True
End Sub
Public Sub Term()
x = thrd.Term()
End Sub
Public Sub Clock()
Do
Form1.Text1.Text = Now
If en = True Then Call thrd.Term
Loop
End Sub
