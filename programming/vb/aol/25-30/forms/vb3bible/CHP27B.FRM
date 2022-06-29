VERSION 2.00
Begin Form formDatasheet 
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Datasheet"
   ClientHeight    =   4476
   ClientLeft      =   540
   ClientTop       =   1560
   ClientWidth     =   5136
   Height          =   4920
   Left            =   480
   LinkTopic       =   "Form2"
   ScaleHeight     =   4476
   ScaleWidth      =   5136
   Top             =   1176
   Width           =   5256
   Begin Data dataGrid 
      Caption         =   "Data1"
      Connect         =   ""
      DatabaseName    =   ""
      Exclusive       =   0   'False
      Height          =   372
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordSource    =   ""
      Top             =   2880
      Visible         =   0   'False
      Width           =   1140
   End
   Begin CommandButton cmndSortByPath 
      Caption         =   "Sort by &Path"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   3960
      Width           =   1275
   End
   Begin CommandButton cmndSortByName 
      Caption         =   "Sort by &Name"
      Height          =   375
      Left            =   180
      TabIndex        =   3
      Top             =   3960
      Width           =   1275
   End
   Begin CommandButton cmndRefresh 
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   3060
      TabIndex        =   2
      Top             =   3960
      Width           =   975
   End
   Begin CommandButton cmndClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4140
      TabIndex        =   1
      Top             =   3960
      Width           =   855
   End
   Begin Grid gridDatasheet 
      BorderStyle     =   0  'None
      Cols            =   3
      FixedCols       =   0
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      GridLines       =   0   'False
      Height          =   3612
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4692
   End
   Begin Image imagTemp 
      DataField       =   "Icon"
      DataSource      =   "dataGrid"
      Height          =   612
      Left            =   120
      Top             =   3240
      Visible         =   0   'False
      Width           =   732
   End
End

Sub cmndClose_Click ()
     formDatasheet.Hide                     'hide the form
End Sub

Sub cmndRefresh_Click ()
     fillGrid ""                            'fill the grid, use existing sort
End Sub

Sub cmndSortByName_Click ()
     fillGrid "fileName"                    'fill the grid, sort by FileName
End Sub

Sub cmndSortByPath_Click ()
     fillGrid "pathName"                    'fill the grid, sort by "PathName"
End Sub

Sub fillGrid (sort)                         '"sort" is passed by the command buttons

     MousePointer = HOURGLASS               'we're going to be a long while!
     Static oldSort                         'remember old sort (for "Refresh" button)
     
     On Error Resume Next
     gridDatasheet.ColWidth(0) = 700
     gridDatasheet.ColWidth(1) = 1400
     gridDatasheet.ColWidth(2) = gridDatasheet.Width - 1500
     gridDatasheet.RowHeight(0) = 300
     gridDatasheet.Row = 0
     gridDatasheet.Col = 1
     gridDatasheet.Text = "Name"
     gridDatasheet.Col = 2
     gridDatasheet.Text = "Path"
     
     If sort = "" Then                      'refresh button doesn't pass sort
          sort = oldSort
     Else
          oldSort = sort                    'update the sort
     End If

     dataGrid.DatabaseName = formMain!dataMain.DatabaseName   'same as main form
     dataGrid.RecordSource = "SELECT * FROM Icon ORDER BY " & sort
     dataGrid.Options = DB_READONLY         'make sure nothing is changed here...
     dataGrid.Refresh                       'build the recordset
     dataGrid.Recordset.MoveLast            'make sure recordcount is current
     gridDatasheet.Rows = dataGrid.Recordset.RecordCount + 1  'make grid big enough
     dataGrid.Recordset.MoveFirst           'go to beginning of recordset
     Do While Not dataGrid.Recordset.EOF    'and while there are still records,
          i = i + 1                         'go to a new grid row
          gridDatasheet.RowHeight(i) = 500  'set the row height to size of the icon
          gridDatasheet.Row = i             'choose the row
          gridDatasheet.Col = 0             'first column
          gridDatasheet.Picture = imagTemp.Picture          'load the icon
          gridDatasheet.Col = 1             'next column
          gridDatasheet.Text = dataGrid.Recordset!filename  'load the filename
          gridDatasheet.Col = 2             'next column
          gridDatasheet.Text = dataGrid.Recordset!PathName  'load the pathname
          dataGrid.Recordset.MoveNext       'go to next record
     Loop
     MousePointer = DEFAULT                 'done!
          
End Sub

Sub Form_Load ()
     fillGrid "fileName"                     'fill the grid, sort by "FileName"
End Sub

