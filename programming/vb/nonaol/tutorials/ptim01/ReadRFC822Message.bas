Attribute VB_Name = "modReadRFC822Message"
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
  Dim strRFC822File As String
  Dim intInputFileNumber As Integer
  Dim strCurrentLine As String
    
  strRFC822File = Command()
  
  If strRFC822File = "" Then strRFC822File = Dir("*.RFC822")
  Do While Dir(strRFC822File) = ""
    strRFC822File = InputBox("File name to open: ", "RFC822 Reader", "jrr.RFC822")
  Loop

  Load frmReadRFC822Messages

  intInputFileNumber = FreeFile
  Open strRFC822File For Input As #intInputFileNumber
  Do While Not EOF(intInputFileNumber)
    Line Input #intInputFileNumber, strCurrentLine
    frmReadRFC822Messages.DisplayMessage.Lines.Add strCurrentLine
  Loop
  Close #intInputFileNumber

  frmReadRFC822Messages.Show
  
End Sub
