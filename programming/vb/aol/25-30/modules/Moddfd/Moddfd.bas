Attribute VB_Name = "modDFD"
Global gobjIDEAppInst As Object
#If Win16 Then
    Declare Function OSWritePrivateProfileString% Lib "KERNEL" Alias "WritePrivateProfileString" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
    Declare Function OSGetPrivateProfileString% Lib "KERNEL" Alias "GetPrivateProfileString" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnString$, ByVal NumBytes As Integer, ByVal FileName$)
#Else
    Declare Function OSWritePrivateProfileString% Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
    Declare Function OSGetPrivateProfileString% Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnString$, ByVal NumBytes As Integer, ByVal FileName$)
#End If

'--------------------------------------------------------------------------
'this is the startup point for the server
'it will add the entry to VB.INI if it doesn't already exist
'so that the add-in is on available next time VB is loaded
'--------------------------------------------------------------------------
Sub Main()
  Dim ReturnString As String
  '--- Check to see if we are in the VB.INI File.  If not, Add ourselves to the INI file
  #If Win16 Then
    Section$ = "Add-Ins16"
  #Else
    Section$ = "Add-Ins32"
  #End If
  ReturnString = String$(12, Chr$(0))
  ErrCode = OSGetPrivateProfileString(Section$, "DataFormDesigner.DFDClass", "NotFound", ReturnString, Len(ReturnString) + 1, "VB.INI")
  If Left(ReturnString, ErrCode) = "NotFound" Then
    ErrCode = OSWritePrivateProfileString%(Section$, "DataFormDesigner.DFDClass", "0", "VB.INI")
  End If
End Sub

'--------------------------------------------------------------------------
'this function strips the file name off of a path/filename
'for use with ISAM databases that need the directory only
'--------------------------------------------------------------------------
Function StripFileName(rsFileName As String) As String
  On Error Resume Next
  Dim i As Integer

  For i = Len(rsFileName) To 1 Step -1
    If Mid(rsFileName, i, 1) = "\" Then
      Exit For
    End If
  Next
  StripFileName = Mid(rsFileName, 1, i - 1)
End Function

'--------------------------------------------------------------------------
'this sub writes out the code that will be added to the VB project
'this is where you would add more code if you would like to
'add to the basic template provided here
'--------------------------------------------------------------------------
Sub WriteFrmCode(fh As Integer)
  On Error GoTo WCErr
  
  Dim i As Integer
  
  Print #fh, "Private Sub cmdAdd_Click()"
  Print #fh, "  Data1.Recordset.AddNew"
  Print #fh, "End Sub"
  Print #fh, ""
  Print #fh, "Private Sub cmdDelete_Click()"
  Print #fh, "  'this may produce an error if you delete the last"
  Print #fh, "  'record or the only record in the recordset"
  Print #fh, "  Data1.Recordset.Delete"
  Print #fh, "  Data1.Recordset.MoveNext"
  Print #fh, "End Sub"
  Print #fh, ""
  Print #fh, "Private Sub cmdRefresh_Click()"
  Print #fh, "  'this is really only needed for multi user apps"
  Print #fh, "  Data1.Refresh"
  Print #fh, "End Sub"
  Print #fh, ""
  Print #fh, "Private Sub cmdUpdate_Click()"
  Print #fh, "  Data1.UpdateRecord"
  Print #fh, "  Data1.Recordset.Bookmark = Data1.Recordset.LastModified"
  Print #fh, "End Sub"
  Print #fh, ""
  Print #fh, "Private Sub cmdClose_Click()"
  Print #fh, "  Unload Me"
  Print #fh, "End Sub"
  Print #fh, ""
  Print #fh, "Private Sub Data1_Error(DataErr As Integer, Response As Integer)"
  Print #fh, "  'This is where you would put error handling code"
  Print #fh, "  'If you want to ignore errors, comment out the next line"
  Print #fh, "  'If you want to trap them, add code here to handle them"
  Print #fh, "  MsgBox ""Data error event hit err:"" & Error$(DataErr)"
  Print #fh, "  Response = 0  'throw away the error"
  Print #fh, "End Sub"
  Print #fh, ""
  Print #fh, "Private Sub Data1_Reposition()"
  Print #fh, "  Screen.MousePointer = vbDefault"
  Print #fh, "  On Error Resume Next"
  Print #fh, "  'This will display the current record position"
  Print #fh, "  'for dynasets and snapshots"
  Print #fh, "  Data1.Caption = ""Record: "" & (Data1.Recordset.AbsolutePosition + 1)"
  Print #fh, "  'for the table object you must set the index property when"
  Print #fh, "  'the recordset gets created and use the following line"
  Print #fh, "  'Data1.Caption = ""Record: "" & (Data1.Recordset.RecordCount * (Data1.Recordset.PercentPosition * 0.01)) + 1"
  Print #fh, "End Sub"
  Print #fh, ""
  Print #fh, "Private Sub Data1_Validate(Action As Integer, Save As Integer)"
  Print #fh, "  'This is where you put validation code"
  Print #fh, "  'This event gets called when the following actions occur"
  Print #fh, "  Select Case Action"
  Print #fh, "    Case vbDataActionMoveFirst"
  Print #fh, "    Case vbDataActionMovePrevious"
  Print #fh, "    Case vbDataActionMoveNext"
  Print #fh, "    Case vbDataActionMoveLast"
  Print #fh, "    Case vbDataActionAddNew"
  Print #fh, "    Case vbDataActionUpdate"
  Print #fh, "    Case vbDataActionDelete"
  Print #fh, "    Case vbDataActionFind"
  Print #fh, "    Case vbDataActionBookMark"
  Print #fh, "    Case vbDataActionClose"
  Print #fh, "  End Select"
  Print #fh, "  Screen.MousePointer = vbHourglass"
  Print #fh, "End Sub"
  Print #fh, ""
  
  'write the code for the bound OLE client control(s)
  For i = 0 To frmDFD.lstOLECtls.ListCount - 1
    Print #fh, "Private Sub oleField" & frmDFD.lstOLECtls.List(i) & "_DblClick()"
    Print #fh, "  'this is the way to get data into an empty ole control"
    Print #fh, "  'and have it saved back to the table"
    Print #fh, "  oleField" & frmDFD.lstOLECtls.List(i) & ".InsertObjDlg"
    Print #fh, "End Sub"
    Print #fh, ""
  Next
  
  Exit Sub
  
WCErr:
  MsgBox Error$
  Exit Sub
  
End Sub
