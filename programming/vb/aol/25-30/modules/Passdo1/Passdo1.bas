Option Explicit

Sub main ()

    Dim nRetries As Integer
    Dim sPassword As String

    nRetries = 3

    Do
        sPassword = InputBox$("Enter the password and click OK")
        nRetries = nRetries - 1
    Loop While sPassword <> "noidea" And nRetries > 0

    If sPassword <> "noidea" Then
        MsgBox "You got the password wrong - Access denied"
    Else
        MsgBox "Welcome to the system - password accepted"
    End If

End Sub

