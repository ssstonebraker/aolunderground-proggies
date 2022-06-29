Attribute VB_Name = "modReadMIMEMessage"
'========================================================================
' Copyright 1999 - Digital Press, John Rhoton
'
' This program has been written to illustrate the Internet Mail protocols.
' It is provided free of charge and unconditionally.  However, it is not
' intended for production use, and therefore without warranty or any
' implication of support.
'
' You can find an explanation of the concepts behind this code in
' the book:  Programmer's Guide to Internet Mail by John Rhoton,
' Digital Press 1999.  ISBN: 1-55558-212-5.
'
' For ordering information please see http://www.amazon.com or
' you can order directly with http://www.bh.com/digitalpress.
'
'========================================================================

Option Explicit

Public Sub Main()
  Dim strMIMEFile As String
  Dim intInputFileNumber As Integer
  Dim strCurrentLine As String
    
  strMIMEFile = Command()
  
  If strMIMEFile = "" Then strMIMEFile = Dir("test.mime")
  Do While Len(strMIMEFile) = 0 Or Dir(strMIMEFile) = ""
    strMIMEFile = InputBox("File name to open: ", "MIME Reader", Dir("*.mime"))
  Loop

  Load frmReadMIMEMessages

  intInputFileNumber = FreeFile
  Open strMIMEFile For Input As #intInputFileNumber
  Do While Not EOF(intInputFileNumber)
    Line Input #intInputFileNumber, strCurrentLine
    frmReadMIMEMessages.DisplayMessage.Lines.Add strCurrentLine
  Loop
  Close #intInputFileNumber

  frmReadMIMEMessages.Show
  
End Sub
