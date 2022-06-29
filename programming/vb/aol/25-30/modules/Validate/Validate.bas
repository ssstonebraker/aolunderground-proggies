Option Explicit
'-----------------------------------------------------------------------------------------
'   Name    :   Validate.Bas
'   Author  :   Peter Wright
'   Date    :   12 February 1993
'
'   Notice  :   This code is freely distributable.
'           :   Peter Wright, Psynet Ltd.
'           :   peter@gendev.demon.co.uk
'-----------------------------------------------------------------------------------------

Sub GotFocus (txtTextBox As TextBox)

'--------------------------------------------------------------------------------------
'   SubName :   GotFocus
'   Author  :   Peter Wright
'   Date    :   13 February 1993
'
'   Params  :   txtTextBox - The text box whose contents need to be selected
'
'   Notes   :   This code should be called from the textbox GotFocus event. It selects
'           :   any text in the text box enabling the user to overwrite if necessary
'
'   Sample  :                   Call GotFocus ( <name of textbox> )
'--------------------------------------------------------------------------------------
'                           C H A N G E    H I S T O R Y
'   [Date]      [Description]                                                   [Who]
'
'   20/6/94     Comments added to the code for Beginners Guide To VB            PJW
'
'--------------------------------------------------------------------------------------

    ' Use the SelStart and SelLength properties to select the values in the text boxes.
    txtTextBox.SelStart = 0
    txtTextBox.SelLength = Len(txtTextBox.Text)

End Sub

Sub KeyCheck (txtTextBox As Control, nKeyValue As Integer)

'--------------------------------------------------------------------------------------
'   SubName :   KeyCheck
'   Author  :   Peter Wright
'   Date    :   14 February 1993
'
'   Params  :   txtTextBox - The text box to check
'           :   nKeyValue  - The KeyAscii parameter passed to the Keypress event
'
'   Notes   :   This code handles keypresses into a text box, checking
'           :       1. That length is not exceeded
'           :       2. That the keys fit the mask
'
'           :   The tag property of the textbox should be set up with the first character
'           :   being A for alpha only, X for alphanumeric, D for date, I for Integer and
'           :   F for float (decimal) number. There should then be a space followed by
'           :   a number which is the maximum number of characters allowed in the text box, ie
'           :
'           :       A 50
'           :       D 10
'           :       I 10
'           :           etc, etc, etc
'
'   Sample  :                   Call Preload ( frmMainForm )
'--------------------------------------------------------------------------------------
'                           C H A N G E    H I S T O R Y
'   [Date]      [Description]                                                   [Who]
'
'   20/6/94     Comments added to the code for Beginners Guide To VB            PJW
'
'--------------------------------------------------------------------------------------

    ' Define a variable to hold the data type of the text box
    Dim sDataType As String

    ' Define a variable to hold the max length of the text box
    Dim nLength As Integer

    If Len(txtTextBox.Tag) = 0 Then Exit Sub

    ' Get the data type of the text box from the first character of the tag property
    sDataType = Left$(txtTextBox.Tag, 1)

    ' Get the maximum allowed length of the text box
    nLength = Val(Mid$(txtTextBox.Tag, 3, Len(txtTextBox.Tag) - 2))

    ' Ignore the backspace key
    If nKeyValue = 8 Then Exit Sub

    ' Check the keypressed for certain values
    Select Case sDataType

        Case "A" ' Accept Alphbetic keys only
            If Asc(UCase(Chr$(nKeyValue))) < 65 Or Asc(UCase(Chr$(nKeyValue))) > 90 Then nKeyValue = 0

        Case "D" ' Accept numbers, along with /
            If (nKeyValue < 48 Or nKeyValue > 57) And nKeyValue <> 47 Then nKeyValue = 0

        Case "I" ' Accept numbers, along with -
            If (nKeyValue < 48 Or nKeyValue > 57) And nKeyValue <> 45 Then nKeyValue = 0

        Case "F" ' Accept numbers, along with - and .
            If (nKeyValue < 48 Or nKeyValue > 57) And (nKeyValue > 46 Or nKeyValue < 45) Then nKeyValue = 0

    End Select

    ' If the new text exceeds the max length of the textbox then reject it.
    If txtTextBox.SelText = "" And Len(txtTextBox.Text) = nLength Then nKeyValue = 0
    
End Sub

