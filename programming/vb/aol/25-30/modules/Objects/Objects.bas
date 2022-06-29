Option Explicit
Global NextFormNum As Integer       ' These variables must be global so
Global frmEnabledTimer As frmMain   ' all instances have access to them.

Sub DisableFrames ()
Dim i As Integer, j As Integer
Dim Frm As Form
    For i = 0 To Forms.Count - 1
        Set Frm = Forms(i)  ' Can now use frm as shorthand for Forms(i)

        If TypeOf Frm Is frmMain Then
        ' Protect ourselves in case there are other forms (types).

            If frmEnabledTimer Is Nothing Or frmEnabledTimer Is Frm Then
            ' Either there isn't a form with an enabled timer, or there is
            ' and it is this instance, so enable option buttons and frame.
                For j = 0 To 2
                    Frm!OptFlashTarget(j).Enabled = True
                Next
                Frm!fraX.Enabled = True
        
            Else
            ' There's a form with a timer enabled and it isn't this one
            ' so disable option buttons and frame.
                For j = 0 To 2
                    Frm!OptFlashTarget(j).Enabled = False
                Next
                Frm!fraX.Enabled = False
            End If
        End If
    Next
End Sub

Sub Flash (Frm As Form)
' Notice that argument is declared As Form.
' This is required because Screen.ActiveForm
' can be passed to this procedure.
Dim temp As Long, i As Integer

' Flash form by swaping forecolors and backcolors twice.
    For i = 1 To 2
        temp = Frm.BackColor
        Frm.BackColor = Frm!lstForms.ForeColor
        Frm!lstForms.BackColor = Frm!lstForms.ForeColor
        Frm!lstForms.ForeColor = temp
        Frm.Refresh
    Next
End Sub

