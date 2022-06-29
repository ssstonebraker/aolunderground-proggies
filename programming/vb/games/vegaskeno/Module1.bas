Attribute VB_Name = "Module1"
Sub LoadResStrings(frm As Form)
    On Error Resume Next
    Dim ctl As Control
    Dim obj As Object
    Dim fnt As Object
    Dim sCtlType As String
    Dim nVal As Integer
        'set the font
    Set fnt = frm.Font
    fnt.Name = LoadResString(20)
    fnt.Size = CInt(LoadResString(21))
    End Sub
    
