'Constants and error IDs for OLE 2.0 Client Control
'Created 1/19/93 - BOBM
'Updated 3/16/93 - BOBM

'Actions
Global Const OLE_CREATE_EMBED = 0
Global Const OLE_CREATE_NEW = 0           'from ole1 control
Global Const OLE_CREATE_LINK = 1
Global Const OLE_CREATE_FROM_FILE = 1     'from ole1 control
Global Const OLE_COPY = 4
Global Const OLE_PASTE = 5
Global Const OLE_UPDATE = 6
Global Const OLE_ACTIVATE = 7
Global Const OLE_CLOSE = 9
Global Const OLE_DELETE = 10
Global Const OLE_SAVE_TO_FILE = 11
Global Const OLE_READ_FROM_FILE = 12
Global Const OLE_INSERT_OBJ_DLG = 14
Global Const OLE_PASTE_SPECIAL_DLG = 15
Global Const OLE_FETCH_VERBS = 17
Global Const OLE_SAVE_TO_OLE1FILE = 18

'OLEType
Global Const OLE_LINKED = 0
Global Const OLE_EMBEDDED = 1
Global Const OLE_NONE = 3

'OLETypeAllowed
Global Const OLE_EITHER = 2

'UpdateOptions
Global Const OLE_AUTOMATIC = 0
Global Const OLE_FROZEN = 1
Global Const OLE_MANUAL = 2

'AutoActivate modes
Global Const OLE_ACTIVATE_MANUAL = 0
Global Const OLE_ACTIVATE_GETFOCUS = 1
Global Const OLE_ACTIVATE_DOUBLECLICK = 2

'SizeModes
Global Const OLE_SIZE_CLIP = 0
Global Const OLE_SIZE_STRETCH = 1
Global Const OLE_SIZE_AUTOSIZE = 2

'DisplayTypes
Global Const OLE_DISPLAY_CONTENT = 0
Global Const OLE_DISPLAY_ICON = 1

'Update Event Constants
Global Const OLE_CHANGED = 0
Global Const OLE_SAVED = 1
Global Const OLE_CLOSED = 2
Global Const OLE_RENAMED = 3

'Special Verb Values
Global Const VERB_PRIMARY = 0
Global Const VERB_SHOW = -1
Global Const VERB_OPEN = -2
Global Const VERB_HIDE = -3

'VerbFlag Bit Masks
Global Const VERBFLAG_GRAYED = &H1
Global Const VERBFLAG_DISABLED = &H2
Global Const VERBFLAG_CHECKED = &H8
Global Const VERBFLAG_SEPARATOR = &H800

'OLE Client Error IDs
'This first set are carried over from previous control
Global Const OLEERR_OutOfMem = 31001
Global Const OLEERR_CantOpenClipboard = 31003
Global Const OLEERR_NoObject = 31004
Global Const OLEERR_CantClose = 31006
Global Const OLEERR_CantPaste = 31007
Global Const OLEERR_InvProp = 31008
Global Const OLEERR_CantCopy = 31009
Global Const OLEERR_InvFormat = 31017
Global Const OLEERR_NoClass = 31018
Global Const OLEERR_NoSourceDoc = 31019
' InvAction is our first new error for the OLE2 control
Global Const OLEERR_InvAction = 31021
Global Const OLEERR_OleInitFailed = 31022
Global Const OLEERR_InvClass = 31023
Global Const OLEERR_CantLink = 31024
Global Const OLEERR_SourceTooLong = 31026
Global Const OLEERR_CantActivate = 31027
Global Const OLEERR_NotRunning = 31028
Global Const OLEERR_DialogBusy = 31029
Global Const OLEERR_InvalidSource = 31031
Global Const OLEERR_CantEmbed = 31032
Global Const OLEERR_CantFetchLinkSrc = 31033
Global Const OLEERR_InvalidVerb = 31034
Global Const OLEERR_NoCompatClipFmt = 31035
Global Const OLEERR_ErrorSavingFile = 31036
Global Const OLEERR_ErrorLoadingFile = 31037
Global Const OLEERR_BadVBVersion = 31038
Global Const OLEERR_CantAccessSource = 31039

' Arrange Method
' for MDI Forms
Global Const CASCADE = 0
Global Const TILE_HORIZONTAL = 1
Global Const TILE_VERTICAL = 2
Global Const ARRANGE_ICONS = 3


Global Const VerbMax = 10

