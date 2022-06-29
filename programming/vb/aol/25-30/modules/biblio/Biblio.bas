Attribute VB_Name = "Module1"
'Declaration for WinHelp
Declare Function OSWinHelp% Lib "User" Alias "WinHelp" (ByVal hwnd%, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

