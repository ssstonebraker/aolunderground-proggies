'------------------------------------------------------------
' VISDATA.BAS
' support functions for the Visual Data sample application
'
' General Information: This app is intended to demonstrate
'   and exercise all of the functionality available in the
'   VT (Virtual Table) Object layer in VB 3.0 Pro.
'
'   Any valid SQL statement may be sent via the Utility SQL
'   function excluding "select" statements which may be
'   executed from the Dynaset Create function. With these
'   two features, this simple app becomes a powerful data
'   definition and query tool accessing any ODBC driver
'   available at the time.
'
'   The app has the capability to perform all DDL (data
'   definition language) functions. These are accessed
'   from the "Tables" form. This form accesses the
'   "NewTable", "AddField" and "IndexAdd" forms to do
'   the actual table, field and index definition.
'   Tables and Indexes may be deleted when the corresponding
'   "Delete" button is enabled. It is not possible to
'   delete fields.
'
' Naming Conventions:
'   "f..."   = Form
'   "c..."   = Form control
'   "F..."   = Form level variable
'   "gst..." = Global String
'   "gf..."  = Global flag (true/false)
'   "gw..."  = Global 2 byte integer value
'
'------------------------------------------------------------

Option Explicit

'api declarations
Declare Function OSGetPrivateProfileString% Lib "Kernel" Alias "GetPrivateProfileString" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnString$, ByVal NumBytes As Integer, ByVal FileName$)
Declare Function OSWritePrivateProfileString% Lib "Kernel" Alias "WritePrivateProfileString" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
Declare Function OSGetWindowsDirectory% Lib "Kernel" Alias "GetWindowsDirectory" (ByVal a$, ByVal b%)

'global object variables
Global gCurrentDB As Database
Global gfDBOpenFlag As Integer
Global gCurrentDS As Dynaset
Global gCurrentTbl As Table
Global gCurrentQueryDef As querydef
Global gCurrentField As Field
Global gCurrentIndex As Index
Global gTableListSS As Snapshot

'global database variables
Global gstDataType As String
Global gstDBName As String
Global gstUserName As String
Global gstPassword As String
Global gstDataBase As String
Global gstDynaString As String
Global gstTblName As String
Global gfUpdatable As Integer
Global glQueryTimeout As Long
Global glLoginTimeout As Long
Global gstTableDynaFilter As String

'other global vars
Global gstZoomData As String
Global gwMaxGridRows As Long

'new field properties
Global gwFldType As Integer
Global gwFldSize As Integer

'global find values
Global gfFindFailed As Integer
Global gstFindExpr As String
Global gstFindOp As String
Global gstFindField As String
Global gfFindMatch As Integer
Global gfFromTableView As Integer

'global seek values
Global gstSeekOperator As String
Global gstSeekValue As String

'global flags
Global gfDBChanged As Integer
Global gfFromSQL As Integer
Global gfTransPending As Integer
Global gfAddTableFlag As Integer

'global constants
Global Const DEFAULTDRIVER = "SQL Server"
Global Const MODAL = 1
Global Const HOURGLASS = 11
Global Const DEFAULT_MOUSE = 0
Global Const YES = 6
Global Const MSGBOX_TYPE = 4 + 48 + 256
Global Const TRUE_ST = "True"
Global Const FALSE_ST = "False"
Global Const EOF_ERR = 626
Global Const FTBLS = 0
Global Const FFLDS = 1
Global Const FINDX = 2
Global Const MAX_GRID_ROWS = 31999
Global Const MAX_MEMO_SIZE = 20000
Global Const GETCHUNK_CUTOFF = 50

'field type constants
Global Const FT_TRUEFALSE = 1
Global Const FT_BYTE = 2
Global Const FT_INTEGER = 3
Global Const FT_LONG = 4
Global Const FT_CURRENCY = 5
Global Const FT_SINGLE = 6
Global Const FT_DOUBLE = 7
Global Const FT_DATETIME = 8
Global Const FT_STRING = 10
Global Const FT_BINARY = 11
Global Const FT_MEMO = 12

'table type constants
Global Const DB_TABLE = 1
Global Const DB_ATTACHEDTABLE = 6
Global Const DB_ATTACHEDODBC = 4
Global Const DB_QUERYDEF = 5
Global Const DB_SYSTEMOBJECT = &H80000002

'dynaset option parameter constants
Global Const VBDA_DENYWRITE = &H1
Global Const VBDA_DENYREAD = &H2
Global Const VBDA_READONLY = &H4
Global Const VBDA_APPENDONLY = &H8
Global Const VBDA_INCONSISTENT = &H10
Global Const VBDA_CONSISTENT = &H20
Global Const VBDA_SQLPASSTHROUGH = &H40

'db create/compact constants
Global Const DB_CREATE_GENERAL = ";langid=0x0809;cp=1252;country=0"
Global Const DB_VERSION10 = 1

' Microsoft Access QueryDef types
Global Const DB_QACTION = &HF0
Global Const DB_QCROSSTAB = &H10
Global Const DB_QDELETE = &H20
Global Const DB_QUPDATE = &H30
Global Const DB_QAPPEND = &H40
Global Const DB_QMAKETABLE = &H50

' Index Attributes
Global Const DB_UNIQUE = 1
Global Const DB_PRIMARY = 2
Global Const DB_PROHIBITNULL = 4
Global Const DB_IGNORENULL = 8
Global Const DB_DESCENDING = 1  'For each field in Index

Function ActionQueryType (qn As String) As String
  Dim i As Integer

  gTableListSS.MoveFirst
  While gTableListSS.EOF = False And gTableListSS!Name <> qn
    gTableListSS.MoveNext
  Wend
  If gTableListSS!Name = qn Then
    Select Case gTableListSS!Attributes
      Case DB_QCROSSTAB
        ActionQueryType = "Cross Tab"
      Case DB_QDELETE
        ActionQueryType = "Delete"
      Case DB_QUPDATE
        ActionQueryType = "Update"
      Case DB_QAPPEND
        ActionQueryType = "Append"
      Case DB_QMAKETABLE
        ActionQueryType = "Make Table"
    End Select
  Else
    ActionQueryType = ""
  End If

End Function

Function CheckTransPending (msg As String) As Integer

  If gfTransPending = True Then
    MsgBox msg + Chr(13) + Chr(10) + "Execute Commit or Rollback First.", 48
    CheckTransPending = True
  Else
    CheckTransPending = False
  End If

End Function

Sub CloseAllDynasets ()
  Dim i As Integer

  MsgBar "Closing Dynasets", True
  While i < forms.Count
    If forms(i).Tag = "Dynaset" Then
      Unload forms(i)
    Else
      i = i + 1
    End If
  Wend
  MsgBar "", False

End Sub

Function CopyData (from_db As Database, to_db As Database, from_nm As String, to_nm As String) As Integer
  On Error GoTo CopyErr

  Dim ds1 As Dynaset, ds2 As Dynaset
  Dim i As Integer

  Set ds1 = from_db.CreateDynaset(from_nm)
  Set ds2 = to_db.CreateDynaset(to_nm)

  While ds1.EOF = False
    ds2.AddNew
    For i = 0 To ds1.Fields.Count - 1
      ds2(i) = ds1(i)
    Next
    ds2.Update
    ds1.MoveNext
  Wend

  CopyData = True
  GoTo CopyEnd

CopyErr:
  ShowError
  CopyData = False
  Resume CopyEnd

CopyEnd:

End Function

Function CopyStruct (from_db As Database, to_db As Database, from_nm As String, to_nm As String, create_ind As Integer) As Integer
  On Error GoTo CSErr

  Dim i As Integer
  Dim tbl As New Tabledef    'table object
  Dim fld As Field           'field object
  Dim ind As Index           'index object

  'search to see if table exists
namesearch:
  For i = 0 To to_db.TableDefs.Count - 1
    If UCase(to_db.TableDefs(i).Name) = UCase(to_nm) Then
      If MsgBox(to_nm + " already exists, delete it?", 4) = YES Then
         to_db.TableDefs.Delete to_db.TableDefs(to_nm)
      Else
         to_nm = InputBox("Enter New Table Name:")
         If to_nm = "" Then
           Exit Function
         Else
           GoTo namesearch
         End If
      End If
      Exit For
    End If
  Next

  'strip off owner if needed
  If InStr(to_nm, ".") <> 0 Then
    to_nm = Mid(to_nm, InStr(to_nm, ".") + 1, Len(to_nm))
  End If
  tbl.Name = to_nm

  'create the fields
  For i = 0 To from_db.TableDefs(from_nm).Fields.Count - 1
    Set fld = New Field
    fld.Name = from_db.TableDefs(from_nm).Fields(i).Name
    fld.Type = from_db.TableDefs(from_nm).Fields(i).Type
    fld.Size = from_db.TableDefs(from_nm).Fields(i).Size
    fld.Attributes = from_db.TableDefs(from_nm).Fields(i).Attributes
    tbl.Fields.Append fld
  Next

  'create the indexes
  If create_ind <> False Then
    For i = 0 To from_db.TableDefs(from_nm).Indexes.Count - 1
      Set ind = New Index
      ind.Name = from_db.TableDefs(from_nm).Indexes(i).Name
      ind.Fields = from_db.TableDefs(from_nm).Indexes(i).Fields
      ind.Unique = from_db.TableDefs(from_nm).Indexes(i).Unique
      If gstDataType <> "ODBC" Then
        ind.Primary = from_db.TableDefs(from_nm).Indexes(i).Primary
      End If
      tbl.Indexes.Append ind
    Next
  End If

  'append the new table
  to_db.TableDefs.Append tbl

  CopyStruct = True
  GoTo CSEnd

CSErr:
  ShowError
  CopyStruct = False
  Resume CSEnd

CSEnd:

End Function

'sub used to create a sample table and fill it
'with NumbRecs number of rows
'can only be called from the debug window
'for example:
'CreateSampleTable "mytbl",100
Sub CreateSampleTable (TblName As String, NumbRecs As Long)
  Dim ds As Dynaset
  Dim ii As Long
  Dim t1 As New Tabledef
  Dim f1 As New Field
  Dim f2 As New Field
  Dim f3 As New Field
  Dim f4 As New Field
  Dim i1 As New Index
  Dim i2 As New Index

  'create the data holding table
  t1.Name = TblName
  
  f1.Name = "name"
  f1.Type = FT_STRING
  f1.Size = 25
  t1.Fields.Append f1

  f2.Name = "address"
  f2.Type = FT_STRING
  f2.Size = 25
  t1.Fields.Append f2

  f3.Name = "record"
  f3.Type = FT_STRING
  f3.Size = 10
  t1.Fields.Append f3

  f4.Name = "id"
  f4.Type = FT_LONG
  f4.Size = 4
  t1.Fields.Append f4

  gCurrentDB.TableDefs.Append t1

  'add the indexes
  i1.Name = TblName + "1"
  i1.Fields = "name"
  i1.Unique = False
  gCurrentDB.TableDefs(TblName).Indexes.Append i1

  i2.Name = TblName + "2"
  i2.Fields = "id"
  i2.Unique = True
  gCurrentDB.TableDefs(TblName).Indexes.Append i2

  'add records to the table in reverse order
  'so indexes have some work to do
  Set ds = gCurrentDB.CreateDynaset(TblName)
  For ii = NumbRecs To 1 Step -1
    ds.AddNew
    ds(0) = "name" + CStr(ii)
    ds(1) = "addr" + CStr(ii)
    ds(2) = "rec" + CStr(ii)
    ds(3) = ii
    ds.Update
  Next

End Sub

Function GetFieldType (ft As String) As Integer
  'return field length
  If ft = "String" Then
    GetFieldType = FT_STRING
  Else
    Select Case ft
      Case "Counter"
        GetFieldType = FT_LONG
      Case "True/False"
        GetFieldType = FT_TRUEFALSE
      Case "Byte"
        GetFieldType = FT_BYTE
      Case "Integer"
        GetFieldType = FT_INTEGER
      Case "Long"
        GetFieldType = FT_LONG
      Case "Currency"
        GetFieldType = FT_CURRENCY
      Case "Single"
        GetFieldType = FT_SINGLE
      Case "Double"
        GetFieldType = FT_DOUBLE
      Case "Date/Time"
        GetFieldType = FT_DATETIME
      Case "Binary"
        GetFieldType = FT_BINARY
      Case "Memo"
        GetFieldType = FT_MEMO
    End Select
  End If

End Function

Function GetFieldWidth (t As Integer)
  'determines the form control width
  'based on the field type
  Select Case t
    Case FT_TRUEFALSE
      GetFieldWidth = 850
    Case FT_BYTE
      GetFieldWidth = 650
    Case FT_INTEGER
      GetFieldWidth = 900
    Case FT_LONG
      GetFieldWidth = 1100
    Case FT_CURRENCY
      GetFieldWidth = 1800
    Case FT_SINGLE
      GetFieldWidth = 1800
    Case FT_DOUBLE
      GetFieldWidth = 2200
    Case FT_DATETIME
      GetFieldWidth = 2000
    Case FT_STRING
      GetFieldWidth = 3250
    Case FT_BINARY
      GetFieldWidth = 3250
    Case FT_MEMO
      GetFieldWidth = 3250
    Case Else
      GetFieldWidth = 3250
  End Select

End Function

Function GetINIString$ (ByVal szItem$, ByVal szDefault$)
  Dim tmp As String
  Dim x As Integer

  tmp = String$(2048, 32)
  x = OSGetPrivateProfileString("VISDATA", szItem$, szDefault$, tmp, Len(tmp), "VISDATA.INI")

  GetINIString = Mid$(tmp, 1, x)
End Function

Function GetNumbRecs (FDS As Dynaset) As Long
  Dim ds As Dynaset

  On Error GoTo GNRErr

  Set ds = FDS.Clone()
  If Not ds.EOF Then ds.MoveLast
  GetNumbRecs = ds.RecordCount
  ds.Close
  If FDS.Updatable = True Then
    gfUpdatable = True
  End If

  GoTo GNREnd

GNRErr:
  'just return because row count is non critical
  GetNumbRecs = -1
  Resume GNREnd

GNREnd:

End Function

Function GetNumbRecsSS (FDS As Snapshot) As Long
  Dim ds As Snapshot

  On Error GoTo GNRSSErr

  Set ds = FDS.Clone()
  If Not ds.EOF Then ds.MoveLast
  GetNumbRecsSS = ds.RecordCount
  ds.Close
  If FDS.Updatable = True Then
    gfUpdatable = True
  End If

  GoTo GNRSSEnd

GNRSSErr:
  'just return because row count is non critical
  GetNumbRecsSS = -1
  Resume GNRSSEnd

GNRSSEnd:

End Function

Function GetNumbRecsTbl (tbl As Table) As Long
  Dim tbl2 As Table

  On Error GoTo GNRTErr

  Set tbl2 = tbl.Clone()
  If Not tbl2.EOF Then tbl2.MoveLast
  GetNumbRecsTbl = tbl2.RecordCount
  tbl2.Close
  gfUpdatable = True

  GoTo GNRTEnd

GNRTErr:
  'just return because row count is non critical
  GetNumbRecsTbl = -1
  Resume GNRTEnd

GNRTEnd:

End Function

'----------------------------------------------------------------------------
'to use this function in any app,
'1. create a form with a grid
'2. create a dynaset
'3. call this function from the form with
'   grd    = your grid control name
'   dynst$ = your dynaset open string (table name or SQL select statement)
'   numb&  = the max number of rows to load (grid is limited to 2000)
'   start& = starting row (needed to display the record number in the
'            left column when loading blocks of records as the
'            DynaGrid form in this app does with the "More" button)
'----------------------------------------------------------------------------
Function LoadGrid (grd As Control, FDS As Snapshot, dynst$, numb&, start&) As Integer
   Dim ft As Integer               'field type
   Dim i As Integer, j As Integer  'for loop indexes
   Dim fn As String                'field name
   Dim rc As Integer               'record count
   Dim gs As String                'grid string

   On Error GoTo LGErr

   MsgBar "Loading Grid for Table View", True
   'setup the grid
   grd.Rows = 2       'reduce the grid
   grd.FixedRows = 0  'allow next step
   grd.Rows = 1       'clears the grid completely
   grd.Cols = FDS.Fields.Count + 1

   If start& = 0 Then        'only do it on first call
     On Error Resume Next
     'set the column widths
     For i = 0 To FDS.Fields.Count - 1
       ft = FDS(i).Type
       If ft = FT_STRING Then
         If FDS(i).Size > Len(FDS(i).Name) Then
           If FDS(i).Size <= 10 Then
             grd.ColWidth(i + 1) = FDS(i).Size * fTables.TextWidth("A")
           Else
             grd.ColWidth(i + 1) = 10 * fTables.TextWidth("A")
           End If
         Else
           If Len(FDS(i).Name) <= 10 Then
             grd.ColWidth(i + 1) = Len(FDS(i).Name) * fTables.TextWidth("A")
           Else
             grd.ColWidth(i + 1) = 10 * fTables.TextWidth("A")
           End If
         End If
       ElseIf ft = FT_MEMO Then
         grd.ColWidth(i + 1) = 1200
       Else
         grd.ColWidth(i + 1) = GetFieldWidth(ft)
       End If
     Next

     On Error GoTo LGErr
     'load the field names
     grd.Row = 0
     For i = 0 To FDS.Fields.Count - 1
       grd.Col = i + 1
       grd.Text = UCase(FDS(i).Name)
     Next
   End If

   rc = 1

   'fill method 1
   'add the rows with the additem method
   While FDS.EOF = False And rc <= numb
     gs = CStr(rc + start) + Chr$(9)
     For i = 0 To FDS.Fields.Count - 1
       If FDS(i).Type = FT_MEMO Then
         If FDS(i).FieldSize() < 255 Then
           gs = gs + StripNonAscii(vFieldVal(FDS(i))) + Chr$(9)
         Else
           'can only get the 1st 255 chars
           gs = gs + StripNonAscii(vFieldVal(FDS(i).GetChunk(0, 255))) + Chr$(9)
         End If
       ElseIf FDS(i).Type = FT_STRING Then
         gs = gs + StripNonAscii(vFieldVal(FDS(i))) + Chr$(9)
       Else
         gs = gs + vFieldVal(FDS(i)) + Chr$(9)
       End If
     Next
     gs = Mid(gs, 1, Len(gs) - 1)
     grd.AddItem gs
     FDS.MoveNext
     rc = rc + 1
   Wend

   'fill method 2
   'add the cells individually
'   While fds.EOF = False And rc <= numb
'     grd.Rows = rc + 1
'     grd.Row = rc
'     grd.Col = 0
'     grd.Text = CStr(rc + start)
'     For i = 0 To fds.Fields.Count - 1
'       grd.Col = i + 1
'       If fds(i).Type = FT_MEMO Then
'         'can only get the 1st 255 chars
'         grd.Text = StripNonAscii(vFieldVal((fds(i).GetChunk(0, 255))))
'       ElseIf fds(i).Type = FT_STRING Then
'         grd.Text = StripNonAscii(vFieldVal((fds(i))))
'       Else
'         grd.Text = CStr(vFieldVal(fds(i)))
'       End If
'     Next
'     fds.MoveNext
'     rc = rc + 1
'   Wend

   grd.FixedRows = 1   'freeze the field names
   grd.FixedCols = 1   'freeze the row numbers
   grd.Row = 1         'set current position
   grd.Col = 1

   LoadGrid = rc       'return number added
   GoTo LGEnd

LGErr:
   ShowError
   LoadGrid = False    'return 0
   Resume LGEnd

LGEnd:
   MsgBar "", False

End Function

Sub MsgBar (msg As String, pw As Integer)
  If msg = "" Then
    VDMDI.cMsg = "Ready"
  Else
    If pw = True Then
      VDMDI.cMsg = msg + ", please wait..."
    Else
      VDMDI.cMsg = msg
    End If
  End If
  VDMDI.cMsg.Refresh
End Sub

Sub Outlines (formname As Form)
    Dim drkgray As Long, fullwhite As Long
    Dim i As Integer
    Dim ctop As Integer, cleft As Integer, cright As Integer, cbottom As Integer

    ' Outline a form's controls for 3D look unless control's TAG
    ' property is set to "skip".

    Dim cname As Control
    drkgray = RGB(128, 128, 128)
    fullwhite = RGB(255, 255, 255)

    For i = 0 To (formname.Controls.Count - 1)
        Set cname = formname.Controls(i)
        If TypeOf cname Is Menu Then
            'Debug.Print "menu item"
        ElseIf (UCase(cname.Tag) = "OL") Then
                ctop = cname.Top - screen.TwipsPerPixelY
                cleft = cname.Left - screen.TwipsPerPixelX
                cright = cname.Left + cname.Width
                cbottom = cname.Top + cname.Height
                formname.Line (cleft, ctop)-(cright, ctop), drkgray
                formname.Line (cleft, ctop)-(cleft, cbottom), drkgray
                formname.Line (cleft, cbottom)-(cright, cbottom), fullwhite
                formname.Line (cright, ctop)-(cright, cbottom), fullwhite
        End If
    Next i
End Sub

Sub PicOutlines (pic As Control, ctl As Control)
    Dim drkgray As Long, fullwhite As Long
    Dim ctop As Integer, cleft As Integer, cright As Integer, cbottom As Integer

    ' Outline a form's controls for 3D look unless control's TAG
    ' property is set to "skip".

    Dim cname As Control
    drkgray = RGB(128, 128, 128)
    fullwhite = RGB(255, 255, 255)

    ctop = ctl.Top - screen.TwipsPerPixelY
    cleft = ctl.Left - screen.TwipsPerPixelX
    cright = ctl.Left + ctl.Width
    cbottom = ctl.Top + ctl.Height
    pic.Line (cleft, ctop)-(cright, ctop), drkgray
    pic.Line (cleft, ctop)-(cleft, cbottom), drkgray
    pic.Line (cleft, cbottom)-(cright, cbottom), fullwhite
    pic.Line (cright, ctop)-(cright, cbottom), fullwhite

End Sub

Sub RefreshTables (tbl_list As Control, IncludeQueries As Integer)
   Dim i As Integer, j As Integer, h As Integer
   Dim st As String
   Dim OkayToAdd As Integer

   On Error GoTo TRefErr

   MsgBar "Refreshing Table List", True
   SetHourglass VDMDI

   Set gTableListSS = gCurrentDB.ListTables()
   tbl_list.Clear

   If IncludeQueries And gstDataType = "MS Access" Then
     ' the ListTables method is used to display querydefs that might
     ' be present in an Access database, see below for optional code
     While gTableListSS.EOF = False
       st = gTableListSS!Name
       If VDMDI.PrefAllowSys.Checked = False Then
         If (gTableListSS!Attributes And DB_SYSTEMOBJECT) = 0 Then
           tbl_list.AddItem st
         End If
       Else
         tbl_list.AddItem st
       End If
       gTableListSS.MoveNext
     Wend
   Else
     ' this method uses the tabledefs collection but will not display
     ' querydefs in an Access database
     tbl_list.Clear
     For i = 0 To gCurrentDB.TableDefs.Count - 1
       st = gCurrentDB.TableDefs(i).Name
       If (gCurrentDB.TableDefs(i).Attributes And DB_SYSTEMOBJECT) = 0 Then
         tbl_list.AddItem st
       End If
     Next
   End If
  
   GoTo TRefEnd

TRefErr:
   ShowError
   gfDBOpenFlag = False
   Resume TRefEnd

TRefEnd:
   ResetMouse VDMDI
   MsgBar "", False

End Sub

Sub ResetMouse (f As Form)
  VDMDI.MousePointer = DEFAULT_MOUSE
  f.MousePointer = DEFAULT_MOUSE
End Sub

Function SetFldProperties (ft As String) As String
  'return field length
  If ft = "String" Then
    gwFldType = FT_STRING
  Else
    Select Case ft
      Case "Counter"
        SetFldProperties = "4"
        gwFldType = FT_LONG
        gwFldSize = 4
      Case "True/False"
        SetFldProperties = "1"
        gwFldType = FT_TRUEFALSE
        gwFldSize = 1
      Case "Byte"
        SetFldProperties = "1"
        gwFldType = FT_BYTE
        gwFldSize = 1
      Case "Integer"
        SetFldProperties = "2"
        gwFldType = FT_INTEGER
        gwFldSize = 2
      Case "Long"
        SetFldProperties = "4"
        gwFldType = FT_LONG
        gwFldSize = 4
      Case "Currency"
        SetFldProperties = "8"
        gwFldType = FT_CURRENCY
        gwFldSize = 8
      Case "Single"
        SetFldProperties = "4"
        gwFldType = FT_SINGLE
        gwFldSize = 4
      Case "Double"
        SetFldProperties = "8"
        gwFldType = FT_DOUBLE
        gwFldSize = 8
      Case "Date/Time"
        SetFldProperties = "8"
        gwFldType = FT_DATETIME
        gwFldSize = 8
      Case "Binary"
        SetFldProperties = "0"
        gwFldType = FT_BINARY
        gwFldSize = 0
      Case "Memo"
        SetFldProperties = "0"
        gwFldType = FT_MEMO
        gwFldSize = 0
    End Select
  End If
End Function

Sub SetHourglass (f As Form)
  DoEvents  'cause forms to repaint before going on
  VDMDI.MousePointer = HOURGLASS
  f.MousePointer = HOURGLASS
End Sub

Sub ShowError ()
  Dim s As String
  Dim crlf As String

  crlf = Chr(13) + Chr(10)
  s = "The following Error occurred:" + crlf + crlf
  'add the error string
  s = s + Error$ + crlf
  'add the error number
  s = s + "Number: " + CStr(Err)
  'beep and show the error
  Beep
  MsgBox (s)

End Sub

Function StripFileName (fname As String) As String
  On Error Resume Next
  Dim i As Integer

  For i = Len(fname) To 1 Step -1
    If Mid(fname, i, 1) = "\" Then
      Exit For
    End If
  Next

  StripFileName = Mid(fname, 1, i - 1)

End Function

Function StripNonAscii (vs As Variant) As String
  Dim i As Integer
  Dim ts As String

  For i = 1 To Len(vs)
    If Asc(Mid(vs, i, 1)) < 32 Or Asc(Mid(vs, i, 1)) > 126 Then
      ts = ts + " "
    Else
      ts = ts + Mid(vs, i, 1)
    End If
  Next

  StripNonAscii = ts

End Function

Function stTrueFalse (tf As Variant) As String
  If tf = True Then
    stTrueFalse = "True"
  Else
    stTrueFalse = "False"
  End If
End Function

Function TableType (tbl As String) As Integer
  Dim i As Integer

  gTableListSS.MoveFirst
  While gTableListSS.EOF = False And gTableListSS!Name <> tbl
    gTableListSS.MoveNext
  Wend
  If gTableListSS!Name = tbl Then
    TableType = gTableListSS!TableType
  Else
    TableType = 0
  End If

End Function

Function vFieldVal (fval As Variant) As Variant
  If IsNull(fval) Then
    vFieldVal = ""
  Else
    vFieldVal = CStr(fval)
  End If
End Function

