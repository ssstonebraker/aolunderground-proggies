'*** Global Constants ***
Global Const OFN_FILEMUSTEXIST = &H1000&
Global Const OFN_READONLY = &H4&

Global Const MCI_APP_TITLE = "MCI Control Application"

' These constants are defined in mmsystem.h.
Global Const MCIERR_INVALID_DEVICE_ID = 30257
Global Const MCIERR_DEVICE_OPEN = 30263
Global Const MCIERR_CANNOT_LOAD_DRIVER = 30266
Global Const MCIERR_UNSUPPORTED_FUNCTION = 30274
Global Const MCIERR_INVALID_FILE = 30304

Global Const MCI_MODE_NOT_OPEN = 524
Global Const MCI_MODE_PLAY = 526

Global Const MCI_FORMAT_MILLISECONDS = 0
Global Const MCI_FORMAT_TMSF = 10

Declare Function GetFocus Lib "User" () As Integer


'*** Global Variables ***
Global DialogCaption As String

