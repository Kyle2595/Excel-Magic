Attribute VB_Name = "Module1"
Dim approvedArray() As Variant
Dim pctCompl As Single

Sub CompileData()
'Main Sub
'This Sub contains mostly formating code, all Private Subs contain logic code

Debug.Print
Debug.Print
Debug.Print Now

Debug.Print "Saving approved states..."
PreserveApproved            'Calls the PreserveApproved() method defined below.  Used for saving approved state of parts so it doesn't get overwriten in master list

Debug.Print "Formatting first sheet..."
'Clear data to avoid adding to an existing list
With ActiveWorkbook.Worksheets(1)
    .Columns.Delete
    .Rows.Delete
End With

'Create Headers on master list
ActiveWorkbook.Worksheets(1).Range("A1") = "PART NUMBER"
ActiveWorkbook.Worksheets(1).Range("B1") = "DESCRIPTION"
ActiveWorkbook.Worksheets(1).Range("C1") = "TYPE"
ActiveWorkbook.Worksheets(1).Range("D1") = "MATERIAL"
ActiveWorkbook.Worksheets(1).Range("E1") = "WETTED PART"
ActiveWorkbook.Worksheets(1).Range("F1") = "QTY"
ActiveWorkbook.Worksheets(1).Range("G1") = "APPROVED"

'Format Headers on master list
With ActiveWorkbook.Worksheets(1).Rows("1")
.Font.Bold = True
.HorizontalAlignment = xlCenter
End With

'Format columns in on master list
ActiveWorkbook.Worksheets(1).Columns("A").ColumnWidth = 15
ActiveWorkbook.Worksheets(1).Columns("B").ColumnWidth = 80
ActiveWorkbook.Worksheets(1).Columns("C").ColumnWidth = 20
ActiveWorkbook.Worksheets(1).Columns("D").ColumnWidth = 15
ActiveWorkbook.Worksheets(1).Columns("E").ColumnWidth = 15
ActiveWorkbook.Worksheets(1).Columns("F").ColumnWidth = 10
ActiveWorkbook.Worksheets(1).Columns("G").ColumnWidth = 15

Debug.Print "Saving parts to master list..."
ReadData                    'Calls the ReadData() method defined below

lastrow = ActiveWorkbook.Worksheets(1).UsedRange.Rows.Count     'Finds the last row used on master list

'Format the Approved column to be a dropdown list
With ActiveWorkbook.Worksheets(1).Range("G2:G" & lastrow).Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
    Operator:=xlBetween, Formula1:="Yes, Yes - With Notes, No"
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = ""
    .ErrorTitle = ""
    .InputMessage = ""
    .ErrorMessage = ""
    .ShowInput = True
    .ShowError = True

End With

'Adds conditional if cell value = "Yes" then cell is green
With ActiveWorkbook.Worksheets(1).Range("G2:G" & lastrow).FormatConditions _
    .Add(xlCellValue, xlEqual, "Yes")
    .Interior.ColorIndex = 4
End With

'Adds conditional if cell value = "Yes - With Notes" then cell is yellow
With ActiveWorkbook.Worksheets(1).Range("G2:G" & lastrow).FormatConditions _
    .Add(xlCellValue, xlEqual, "Yes - With Notes")
    .Interior.ColorIndex = 6
End With

'Adds conditional if cell value = "No" then cell is Red
With ActiveWorkbook.Worksheets(1).Range("G2:G" & lastrow).FormatConditions _
    .Add(xlCellValue, xlEqual, "No")
    .Interior.ColorIndex = 3
End With

Debug.Print "Sorting parts and removing duplicates from the master list..."
MergeParts                  'Calls the MergeParts() method defined below

Debug.Print "Loading approved states..."
PopulateApprovedValues

Debug.Print "Done!"

End Sub

Private Sub ReadData()
'Loop through all worksheets (except Master List) and copy the data to the first Worksheet

Dim WS_Count As Integer     'Count of number of Worksheets
Dim i As Integer

'Set WS_Count equal to the number of worksheets in the active workbook.
WS_Count = ActiveWorkbook.Worksheets.Count

For i = 2 To WS_Count       'Loop through all worksheets starting at the second one

    Dim rangeArray As Variant
    Dim valueArray() As Variant
    rangeArray = ActiveWorkbook.Worksheets(i).UsedRange.Value
    'rangeArray now contains the values of ALL the cells in used range of the I'th Worksheet.
         
    'Loops through all cells in I'th Worksheet and save them to the master list (Nested loop O(n^2) time complexity)
    Dim lngCol As Long, lngRow As Long
    For lngRow = 2 To UBound(rangeArray, 1)
    NextRow = ActiveWorkbook.Worksheets(1).Range("A" & Rows.Count).End(xlUp).Row + 1
        For lngCol = 1 To 6
            'Debug.Print rangeArray(lngRow, lngCol)    'UNCOMMENT FOR TESTING
            ReDim valueArray(UBound(rangeArray, 2))
            valueArray(lngCol) = rangeArray(lngRow, lngCol)
            ActiveWorkbook.Worksheets(1).Cells(NextRow, lngCol).Value = valueArray(lngCol)  'Save value to Master List
        Next lngCol
    Next lngRow

Next i

End Sub

'This sub is never called, but I left it in just in case you want to use it in the future
'If you want to impliment it, you need to add a Worksheet for it to write to.  Otherwise it will overwrite the Worksheet in the second position
Private Sub CreatePivotTable()

'Remove Prevous Pivot Table so they don't overlap
On Error Resume Next
ActiveWorkbook.Worksheets(2).PivotTables(1).TableRange2.Clear

Dim pvtCache As PivotCache
Dim pvt As PivotTable

'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, sourceData:=ActiveWorkbook.Worksheets(1).UsedRange)

'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.CreatePivotTable(TableDestination:=Worksheets("QTY Lookup").Range("A3"), TableName:="Parts QTY")
    
 'Add items to the Pivot Table
    pvt.PivotFields("PART NUMBER").Orientation = xlRowField
    pvt.AddDataField pvt.PivotFields("QTY"), "Sum of QTY", xlSum
    
End Sub

Private Sub SortData()
'Sorts the Master List acording to part number

Dim lastrow As Long
lastrow = ActiveWorkbook.Worksheets(1).Cells(Rows.Count, 2).End(xlUp).Row

'key1 defines how this is sorted
ActiveWorkbook.Worksheets(1).Range("A2:G" & lastrow).Sort key1:=Range("A2:A" & lastrow), order1:=xlAscending, Header:=xlNo

End Sub

Private Sub MergeParts()
'Merges duplicates after sorting parts and keeps track of quantities

SortData        'Call the SortData() method

lastrow = ActiveWorkbook.Worksheets(1).UsedRange.Rows.Count
Set r = ActiveWorkbook.Worksheets(1).UsedRange.Resize(1)
With Application.WorksheetFunction

    For i = lastrow - 1 To 2 Step -1     'Loop through the rows
        Do While Cells(i, 1) = Cells(i + 1, 1)
            LastCol = r(r.Count).Column
            SumCol = LastCol + 1
            Cells(i, 6) = .Sum(Range(Cells(i, 6), Cells(i + 1, 6)))
            Rows(i + 1).Delete
        Loop
    Next i

End With

End Sub

Private Sub PreserveApproved()
'Saves the approved state of all the parts on the Master List so they don't get overwritten
'This only applies to the Master List, approved states on assembly worksheets are pulled from the Master List

lastrow = ActiveWorkbook.Worksheets(1).UsedRange.Rows.Count
ReDim approvedArray(lastrow, 2)     'approvedArray is defined at the very top of the program as a global variable

'Loop through the first Worksheet assigning the values of the Part Nuumber column and Approved column into the 2D array
For i = 2 To lastrow        'Array will start at (2,0) because of header offset
        approvedArray(i, 0) = ActiveWorkbook.Worksheets(1).Cells(i, 1)          'Part Number
        approvedArray(i, 1) = ActiveWorkbook.Worksheets(1).Cells(i, 7)          'Approved State
        approvedArray(i, 2) = ActiveWorkbook.Worksheets(1).Cells(i, 7).Address  'Excel address of approved state
    
Next i

End Sub

Private Sub AssignApproved(Wkst As Integer)
'Assigns the approved states of each part and propigates them throughout the assembly worksheets
'This is the logic method for PopulateApprovedValues() which calls this method for every sheet in the Excel document

lastrow = ActiveWorkbook.Worksheets(Wkst).UsedRange.Rows.Count

If Wkst > 1 Then
    'Adds conditional if cell value = "Yes" then cell is green
    With ActiveWorkbook.Worksheets(Wkst).Range("G2:G" & lastrow).FormatConditions _
        .Add(xlCellValue, xlEqual, "Yes")
        .Interior.ColorIndex = 4
    End With
    
    'Adds conditional if cell value = "Yes - With Notes" then cell is yellow
    With ActiveWorkbook.Worksheets(Wkst).Range("G2:G" & lastrow).FormatConditions _
        .Add(xlCellValue, xlEqual, "Yes - With Notes")
        .Interior.ColorIndex = 6
    End With

    'Adds conditional if cell value = "No" then cell is Red
    With ActiveWorkbook.Worksheets(Wkst).Range("G2:G" & lastrow).FormatConditions _
        .Add(xlCellValue, xlEqual, "No")
        .Interior.ColorIndex = 3
    End With
End If

'Nested loop menas O(n^3) time complexity which is kinda garbage (Worksheet -> Row -> array)
For i = 2 To lastrow        'Loop through the rows of the active Worksheet
    For j = LBound(approvedArray) To UBound(approvedArray)      'Loop through approvedArray
        If (approvedArray(j, 0) = ActiveWorkbook.Worksheets(Wkst).Cells(i, 1)) Then
        
            'This if statement saves the VALUE of the approved state on the Master List and the ADDRESS of the approved state on the assembly worksheets
            If Wkst > 1 Then
            On Error Resume Next
                ActiveWorkbook.Worksheets(Wkst).Cells(i, 7) = "=IF(TRIM('All Parts'!" & approvedArray(j, 2) & ")" & "<>" & Chr(34) & Chr(34) & ", 'All Parts'!" & approvedArray(j, 2) & ", " & Chr(34) & Chr(34) & ")"
            Else
                ActiveWorkbook.Worksheets(Wkst).Cells(i, 7) = approvedArray(j, 1)

            End If
        End If
    Next j      'End first loop
Next i          'End second loop

End Sub

Private Sub PopulateApprovedValues()
'Take Approved values from the master list and populate it to the individual style number Worksheets
'Also re-saves approved values to masterlist after formatting is complete

Dim WS_Count As Integer     'Count of number of Worksheets

'Set WS_Count equal to the number of worksheets in the active workbook.
WS_Count = ActiveWorkbook.Worksheets.Count

For i = 1 To WS_Count       'Begin the Worksheet loop.
  
  AssignApproved (i)
  
Next i

End Sub
