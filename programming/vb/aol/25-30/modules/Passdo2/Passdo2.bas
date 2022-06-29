Option Explicit

Sub main ()

    Dim nRetries As Integer
    Dim sPassword As String

    nRetries = 3

    Do
        sPassword = InputBox$("Enter the password and click OK")
        nRetries = nRetries - 1
    Loop Until sPassword = "noidea"


End Sub

