Attribute VB_Name = "modIMEXCode"
Option Explicit

'global vars used in the Import Export Code
Global gnDataType As Integer
Global gImpDB As Database
Global gExpDB As Database
Global gExpTable As String

'data types
Global Const gnDT_NONE = -1
Global Const gnDT_JETMDB = 0
Global Const gnDT_DBASEIV = 1
Global Const gnDT_DBASEIII = 2
Global Const gnDT_FOXPRO26 = 3
Global Const gnDT_FOXPRO25 = 4
Global Const gnDT_FOXPRO20 = 5
Global Const gnDT_PARADOX4X = 6
Global Const gnDT_PARADOX3X = 7
Global Const gnDT_BTRIEVE = 8
Global Const gnDT_EXCEL50 = 9
Global Const gnDT_EXCEL40 = 10
Global Const gnDT_EXCEL30 = 11
Global Const gnDT_TEXTFILE = 12
Global Const gnDT_SQLDB = 13

Sub Export(rsFromTbl As String, rsToDB As String)

  On Error GoTo ExpErr

  Dim sConnect As String
  Dim sNewTblName As String
  Dim sDBName As String
  Dim nErrState As Integer
  Dim idxFrom As Index
  Dim idxTo As Index
  Dim sSQL As String              'local copy of sql string
  Dim sField As String
  Dim sFrom As String
  Dim sTmp As String
  Dim i As Integer

  If gnDataType = gnDT_SQLDB Then
    Set gExpDB = gwsMainWS.OpenDatabase(gsNULL_STR, 0, 0, "odbc;")
    If gExpDB Is Nothing Then Exit Sub
  End If

  MsgBar "Exporting '" & rsFromTbl & "'", True

  nErrState = 1
  Select Case gnDataType
    Case gnDT_JETMDB
      sConnect = "[;database=" & rsToDB & "]."
      Set gExpDB = gwsMainWS.OpenDatabase(rsToDB)
    Case gnDT_PARADOX3X
      sDBName = StripFileName(rsToDB)
      sConnect = "[Paradox 3.X;database=" & StripFileName(rsToDB) & "]."
      Set gExpDB = gwsMainWS.OpenDatabase(sDBName, 0, 0, gsPARADOX3X)
    Case gnDT_PARADOX4X
      sDBName = StripFileName(rsToDB)
      sConnect = "[Paradox 4.X;database=" & StripFileName(rsToDB) & "]."
      Set gExpDB = gwsMainWS.OpenDatabase(sDBName, 0, 0, gsPARADOX4X)
    Case gnDT_FOXPRO26
      sDBName = StripFileName(rsToDB)
      sConnect = "[FoxPro 2.6;database=" & StripFileName(rsToDB) & "]."
      Set gExpDB = gwsMainWS.OpenDatabase(sDBName, 0, 0, gsFOXPRO26)
    Case gnDT_FOXPRO25
      sDBName = StripFileName(rsToDB)
      sConnect = "[FoxPro 2.5;database=" & StripFileName(rsToDB) & "]."
      Set gExpDB = gwsMainWS.OpenDatabase(sDBName, 0, 0, gsFOXPRO25)
    Case gnDT_FOXPRO20
      sDBName = StripFileName(rsToDB)
      sConnect = "[FoxPro 2.0;database=" & StripFileName(rsToDB) & "]."
      Set gExpDB = gwsMainWS.OpenDatabase(sDBName, 0, 0, gsFOXPRO20)
    Case gnDT_DBASEIV
      sDBName = StripFileName(rsToDB)
      sConnect = "[dBase IV;database=" & StripFileName(rsToDB) & "]."
      Set gExpDB = gwsMainWS.OpenDatabase(sDBName, 0, 0, gsDBASEIV)
    Case gnDT_DBASEIII
      sDBName = StripFileName(rsToDB)
      sConnect = "[dBase III;database=" & StripFileName(rsToDB) & "]."
      Set gExpDB = gwsMainWS.OpenDatabase(sDBName, 0, 0, gsDBASEIII)
    Case gnDT_BTRIEVE
      sConnect = "[Btrieve;database=" & rsToDB & "]."
      Set gExpDB = gwsMainWS.OpenDatabase(rsToDB, 0, 0, gsBTRIEVE)
    Case gnDT_EXCEL50, gnDT_EXCEL40, gnDT_EXCEL30
      sConnect = "[Excel 5.0;database=" & rsToDB & "]."
      Set gExpDB = gwsMainWS.OpenDatabase(rsToDB, 0, 0, gsEXCEL50)
    Case gnDT_SQLDB
      sConnect = "[" & gExpDB.Connect & "]."
    Case gnDT_TEXTFILE
      sDBName = StripFileName(rsToDB)
      sConnect = "[Text;database=" & StripFileName(rsToDB) & "]."
      Set gExpDB = gwsMainWS.OpenDatabase(sDBName, 0, 0, gsTEXTFILES)
  End Select
  If gnDataType = gnDT_JETMDB Or gnDataType = gnDT_BTRIEVE Or _
     gnDataType = gnDT_SQLDB Or gnDataType = gnDT_EXCEL50 Or _
     gnDataType = gnDT_EXCEL40 Or gnDataType = gnDT_EXCEL30 Then
    With frmExpName
      .Label1.Caption = "Export " & rsFromTbl & " to:"
      .Label2.Caption = "in " & rsToDB
      .txtTable.Text = rsFromTbl
    End With
    frmExpName.Show vbModal
      
    If Len(gExpTable) = 0 Then
      MsgBar gsNULL_STR, False
      Exit Sub
    Else
      sNewTblName = gExpTable
    End If
  Else
    'get the table part of the file name
    'strip off the path
    For i = Len(rsToDB) To 1 Step -1
      If Mid(rsToDB, i, 1) = "\" Then
        Exit For
      End If
    Next
    sTmp = Mid(rsToDB, i + 1, Len(rsToDB))
    'strip off the extension
    For i = 1 To Len(sTmp)
      If Mid(sTmp, i, 1) = "." Then
        Exit For
      End If
    Next
    sNewTblName = Left(sTmp, i - 1)
  End If
  SetHourglass
  If Len(rsFromTbl) > 0 Then
    gdbCurrentDB.Execute "select * into " & sConnect & StripOwner(sNewTblName) & " from " & StripOwner(rsFromTbl)

    If gnDataType <> gnDT_TEXTFILE Then
      nErrState = 2
      MsgBar "Creating Indexes for '" & sNewTblName & "'", True
      gExpDB.Tabledefs.Refresh
      For Each idxFrom In gdbCurrentDB.Tabledefs(rsFromTbl).Indexes
        Set idxTo = gExpDB.Tabledefs(sNewTblName).CreateIndex(idxFrom.Name)
        With idxTo
          .Fields = idxFrom.Fields
          .Unique = idxFrom.Unique
          If gnDataType <> gnDT_SQLDB And gsDataType <> "ODBC" Then
            .Primary = idxFrom.Primary
          End If
        End With
        gExpDB.Tabledefs(sNewTblName).Indexes.Append idxTo
      Next
    End If
    MsgBar gsNULL_STR, False
    Screen.MousePointer = vbDefault
    MsgBox "Successfully Exported '" & rsFromTbl & "'.", 64
  Else
    sSQL = frmSQL.txtSQLStatement.Text
    sField = Mid(sSQL, 8, InStr(8, UCase(sSQL), "FROM") - 9)
    sFrom = " " & Mid(sSQL, InStr(UCase(sSQL), "FROM"), Len(sSQL))
    gdbCurrentDB.Execute "select " & sField & " into " & sConnect & sNewTblName & sFrom

    Screen.MousePointer = vbDefault
    MsgBar gsNULL_STR, False
    MsgBox "Successfully Exported SQL Statement.", 64
  End If

  Exit Sub

ExpErr:
  If Err = 3010 Then      'table exists
    If MsgBox("'" & sNewTblName & "' already exists - overwrite?", 32 + 1 + 256) = 1 Then
      gExpDB.Tabledefs.Delete sNewTblName
      Resume
    Else
      Screen.MousePointer = vbDefault
      MsgBar gsNULL_STR, False
      Exit Sub
    End If
  End If
 
  'nuke the new table if the indexes couldn't be created
  If nErrState = 2 Then
    gExpDB.Tabledefs.Delete sNewTblName
  End If
  ShowError
  Exit Sub

End Sub

Sub Import(rsImpTblName As String)
  On Error GoTo ImpErr

  Dim sOldTblName As String, sNewTblName As String, sConnect As String
  Dim idxFrom As Index
  Dim idxTo As Index
  Dim nErrState As Integer
  Dim i As Integer

  sOldTblName = MakeTableName(rsImpTblName, False)
  sNewTblName = MakeTableName(rsImpTblName, True)

  SetHourglass
  MsgBar "Importing '" & sNewTblName & "'", True

  nErrState = 1
  Select Case gnDataType
    Case gnDT_JETMDB
      sConnect = "[;database=" & gImpDB.Name & "]."
    Case gnDT_PARADOX3X
      sConnect = "[Paradox 3.X;database=" & StripFileName(rsImpTblName) & "]."
      Set gImpDB = gwsMainWS.OpenDatabase(StripFileName(rsImpTblName), 0, 0, gsPARADOX3X)
    Case gnDT_PARADOX4X
      sConnect = "[Paradox 4.X;database=" & StripFileName(rsImpTblName) & "]."
      Set gImpDB = gwsMainWS.OpenDatabase(StripFileName(rsImpTblName), 0, 0, gsPARADOX4X)
    Case gnDT_FOXPRO26
      sConnect = "[FoxPro 2.6;database=" & StripFileName(rsImpTblName) & "]."
      Set gImpDB = gwsMainWS.OpenDatabase(StripFileName(rsImpTblName), 0, 0, gsFOXPRO26)
    Case gnDT_FOXPRO25
      sConnect = "[FoxPro 2.5;database=" & StripFileName(rsImpTblName) & "]."
      Set gImpDB = gwsMainWS.OpenDatabase(StripFileName(rsImpTblName), 0, 0, gsFOXPRO25)
    Case gnDT_FOXPRO20
      sConnect = "[FoxPro 2.0;database=" & StripFileName(rsImpTblName) & "]."
      Set gImpDB = gwsMainWS.OpenDatabase(StripFileName(rsImpTblName), 0, 0, gsFOXPRO20)
    Case gnDT_DBASEIV
      sConnect = "[dBase IV;database=" & StripFileName(rsImpTblName) & "]."
      Set gImpDB = gwsMainWS.OpenDatabase(StripFileName(rsImpTblName), 0, 0, gsDBASEIV)
    Case gnDT_DBASEIII
      sConnect = "[dBase III;database=" & StripFileName(rsImpTblName) & "]."
      Set gImpDB = gwsMainWS.OpenDatabase(StripFileName(rsImpTblName), 0, 0, gsDBASEIII)
    Case gnDT_BTRIEVE
      sConnect = "[Btrieve;database=" & gImpDB.Name & "]."
    Case gnDT_EXCEL50, gnDT_EXCEL40, gnDT_EXCEL30
      sConnect = "[Excel 5.0;database=" & gImpDB.Name & "]."
    Case gnDT_SQLDB
      sConnect = "[" & gImpDB.Connect & "]."
    Case gnDT_TEXTFILE
      sConnect = "[Text;database=" & StripFileName(rsImpTblName) & "]."
      Set gImpDB = gwsMainWS.OpenDatabase(StripFileName(rsImpTblName), 0, 0, gsTEXTFILES)
  End Select
  gdbCurrentDB.Execute "select * into " & sNewTblName & " from " & sConnect & sOldTblName

  If gnDataType <> gnDT_TEXTFILE And gnDataType <> gnDT_EXCEL50 And _
     gnDataType <> gnDT_EXCEL40 And gnDataType <> gnDT_EXCEL30 Then
    nErrState = 2
    MsgBar gdbCurrentDB.RecordsAffected & " Rows Imported, Creating Indexes for '" & sNewTblName & "'", True
    gdbCurrentDB.Tabledefs.Refresh
    For Each idxFrom In gImpDB.Tabledefs(sOldTblName).Indexes
      Set idxTo = gdbCurrentDB.Tabledefs(sNewTblName).CreateIndex(idxFrom.Name)
      With idxTo
        .Fields = idxFrom.Fields
        .Unique = idxFrom.Unique
        If gnDataType <> gnDT_SQLDB And gsDataType <> gsSQLDB Then
          .Primary = idxFrom.Primary
        End If
      End With
      gdbCurrentDB.Tabledefs(sNewTblName).Indexes.Append idxTo
    Next
  End If
    
  frmImpExp.lstTables.AddItem sNewTblName
  frmTables.lstTables.AddItem sNewTblName
  Screen.MousePointer = vbDefault
  MsgBar gsNULL_STR, False
  MsgBox "Successfully Imported '" & sNewTblName & "'.", 64

  Exit Sub

NukeNewTbl:
  On Error Resume Next  'just in case it fails
  gdbCurrentDB.Tabledefs.Delete sNewTblName
  ShowError
  Exit Sub
 
ImpErr:
  'nuke the new table if the indexes couldn't be created
  If nErrState = 2 Then
    Resume NukeNewTbl
  End If
  ShowError
  Exit Sub

End Sub

Function MakeTableName(fname As String, newname As Integer) As String
  On Error Resume Next
  Dim i As Integer, t As Integer
  Dim tmp As String

  If gnDataType = gnDT_SQLDB And newname Then
    i = InStr(1, fname, ".")
    If i > 0 Then
      tmp = Mid(fname, 1, i - 1) & "_" & Mid(fname, i + 1, Len(fname))
    End If
  ElseIf InStr(fname, "\") > 0 Then
    'strip off path
    For i = Len(fname) To 1 Step -1
      If Mid(fname, i, 1) = "\" Then
        Exit For
      End If
    Next
    tmp = Mid(fname, i + 1, Len(fname))
    i = InStr(1, tmp, ".")
    If i > 0 Then
      tmp = Mid(tmp, 1, i - 1)
    End If
  Else
    tmp = fname
  End If

  If newname Then
    If DupeTableName(tmp) Then
      t = 1
      While DupeTableName(tmp + CStr(t))
        t = t + 1
      Wend
      tmp = tmp + CStr(t)
    End If
  End If

  MakeTableName = tmp

End Function
