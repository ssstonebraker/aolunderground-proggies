' Force all runtime errors to be handled here.
Sub DisplayErrorMessageBox ()
    Select Case Err
        Case MCIERR_CANNOT_LOAD_DRIVER
            Msg$ = "Error load media device driver."
        Case MCIERR_DEVICE_OPEN
            Msg$ = "The device is not open or is not known."
        Case MCIERR_INVALID_DEVICE_ID
            Msg$ = "Invalid device id."
        Case MCIERR_INVALID_FILE
            Msg$ = "Invalid filename."
        Case MCIERR_UNSUPPORTED_FUNCTION
            Msg$ = "Action not available for this device."
        Case Else
            Msg$ = "Unknown error (" + Str$(Err) + ")."
    End Select

    MsgBox Msg$, 48, MCI_APP_TITLE
End Sub

' This subroutine allows any Windows events to be processed.
' This may be necessary to solve any synchronization
' problems with Windows events.
'
' This subroutine can also be used to force a delay in
' processing.
Sub WaitForEventsToFinish (NbrTimes As Integer)
    Dim i As Integer

    For i = 1 To NbrTimes
        dummy% = DoEvents()
    Next i
End Sub

