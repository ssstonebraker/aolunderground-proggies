Attribute VB_Name = "Module1"
Public wbWorld As Object
Public shtWorld As Object

Sub Setup()
    ChDir App.Path
    ChDrive App.Path
    ' Get the first sheet in WORLD.XLS.
    Set shtWorld = GetObject("world.xls")
    ' Get the workbook.
    Set wbWorld = shtWorld.Application.Workbooks("world.xls")
End Sub

' Set the objects to Nothing.
Sub CleanUp()
    ' This should force an unload of Microsoft Excel,
    ' providing no other applications or users have it loaded.
    Set shtWorld = Nothing
    Set wbWorld = Nothing
End Sub

' Fill the Continents combo box with the names
' of the sheets in the workbook.
Sub FillContinentsList()
    Dim shtContinent As Object
    
    ' Iterate through the collection of sheets and add
    ' the name of each sheet to the combo box.
    For Each shtContinent In wbWorld.Sheets
        Form1.listContinents.AddItem shtContinent.Name
    Next
    ' Select the first item and display it in the combo box.
    Form1.listContinents.Text = Form1.listContinents.List(0)

    Set shtContinent = Nothing
End Sub

' Fill the Continents combo box with the names
' of the features corresponding to a given continent.
Sub FillFeaturesList()
    Dim shtContinent As Object
    Dim rngFeatureList As Object
    Dim intFirstBlankCell As Integer
    Dim loop1 As Integer

    ' Hide the old ranking list.
    Form1.listTopRanking.Visible = False
    
    ' Get the sheet with the name of the continent selected in the Continents combo box.
    Set shtContinent = wbWorld.Sheets(Form1.listContinents.Text)
    ' Assign the first row of this sheet to an object.
    Set rngFeatureList = shtContinent.rows(1)
    
    ' See if it's an empty list.
    If (rngFeatureList.Cells(1, 1) = "") Then
        intFirstBlankCell = 0
    Else
        ' Search the row for the first blank cell.
        intFirstBlankCell = rngFeatureList.find("").column
    End If
    
    ' Empty the previous contents of the features combo box.
    Form1.listFeatures.Clear
            
    ' Add the items to the features combo box.
    For loop1 = 1 To intFirstBlankCell
            Form1.listFeatures.AddItem rngFeatureList.Cells(1, loop1)
    Next
    
    ' Select the first item and display it in the combo box.
    Form1.listFeatures.Text = Form1.listFeatures.List(0)

    ' Clean up.
    Set shtContinent = Nothing
    Set rngFeatureList = Nothing
End Sub

' Fill the list of ranking items.
Sub FillTopRankingList()
    Dim shtContinent As Object
    Dim intColumOfFeature As Integer
    Dim rngRankedList As Object
    Dim intFirstBlankCell As Integer
    Dim loop1 As Integer
    
    ' Get the sheet with the name of the continent selected in the Continents combo box.
    Set shtContinent = wbWorld.Sheets(Form1.listContinents.Text)
    
    ' Empty the previous contents of the ranking list box.
    Form1.listTopRanking.Clear
    
    ' If the feature selection is blank, do nothing.
    If (Form1.listFeatures <> "") Then
        
        ' Look up the column of the selected feature in the first row of the spreadsheet.
        intColumOfFeature = shtContinent.rows(1).find(Form1.listFeatures.Text).column
        
        ' Assign the column to an object.
         Set rngRankedList = shtContinent.Columns(intColumOfFeature)
        
        ' See if it's a blank list.
        If (rngRankedList.Cells(1, 1) = "") Then
            intFirstBlankCell = 0
        Else
            ' Search the row for the first blank cell.
            intFirstBlankCell = rngRankedList.find("").row
        End If
                
        ' Add the items to the features combo box.
        For loop1 = 2 To intFirstBlankCell
            Form1.listTopRanking.AddItem rngRankedList.Cells(loop1, 1)
        Next
    
        ' Show the new ranking list.
        Form1.listTopRanking.Visible = True
    
    End If
    
    ' Clean up.
    Set shtContinent = Nothing
    Set rngRankedList = Nothing
End Sub
