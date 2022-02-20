Attribute VB_Name = "Module1"
Function toJSON(rangeToParse As Range, parseAsArrays As Boolean) As String
    Dim rowCounter As Integer
    Dim columnCounter As Integer
    Dim parsedData As String: parsedData = "["
    Dim temp As String

    If parseAsArrays Then ' Check to see if we need to make our JSON an array; if not, we'll make it an object
        For rowCounter = 1 To rangeToParse.Rows.Count ' Loop through each row
            temp = "" ' Reset temp's value

            For columnCounter = 1 To rangeToParse.Columns.Count ' Loop through each column
                temp = temp & """" & rangeToParse.Cells(rowCounter, columnCounter) & """" & ","
            Next

            temp = "[" & Left(temp, Len(temp) - 1) & "]," ' Remove extra comma from after last object
            parsedData = parsedData & temp ' Add temp to the data we've already parsed
        Next
    Else
        For rowCounter = 2 To rangeToParse.Rows.Count ' Loop through each row starting with the second row so we don't include the header
            temp = "" ' Reset temp's value

            For columnCounter = 1 To rangeToParse.Columns.Count ' Loop through each column
                temp = temp & """" & rangeToParse.Cells(1, columnCounter) & """" & ":" & """" & rangeToParse.Cells(rowCounter, columnCounter) & """" & ","
            Next

            temp = "{" & Left(temp, Len(temp) - 1) & "}," ' Remove extra comma from after last object
            parsedData = parsedData & temp ' Add temp to the data we've already parsed
        Next
    End If

    parsedData = Left(parsedData, Len(parsedData) - 1) & "]" ' Remove extra comma and add the closing bracket for the JSON array
    toJSON = parsedData ' Return the JSON data
End Function

Function getValuesRange(sheet As String) As Range
    ' Row variables
    Dim usedRows As Integer: usedRows = 0
    Dim rowCounter As Integer: rowCounter = 1
    Dim rowsToCount As Integer: rowsToCount = 1000
    ' Column variables
    Dim usedColumns As Integer: usedColumns = 0
    Dim columnCounter As Integer: columnCounter = 1
    Dim columnsToCount As Integer: columnsToCount = 50

    Do While rowCounter <= rowsToCount ' Loop through each row
        Do While columnCounter <= columnsToCount ' Loop through each column
            If Worksheets(sheet).Cells(rowCounter, columnCounter) <> "" Then ' Check to see if the cell has a value
                usedRows = rowCounter ' Since the current row has a cell with a value in it, set usedRows to the current row

                If columnCounter > usedColumns Then
                    usedColumns = columnCounter ' If the current column is greater than usedColumns, set usedColumns to the current column
                End If

                If usedRows = rowsToCount Then
                    rowsToCount = rowsToCount + 100 ' If the value of usedRows reaches the rowsToCount limit, then extend the rowsToCount limit by 100
                End If

                If usedColumns = columnsToCount Then
                    columnsToCount = columnsToCount + 50 ' If the value of usedColumns reaches the columnsToCount limit, then extend the columnsToCount limit by 100
                End If
            End If
            columnCounter = columnCounter + 1 ' Increment columnCounter
        Loop

        rowCounter = rowCounter + 1 ' Increment rowCounter
        columnCounter = 1 ' Reset the columnCounter to 1 so we're always checking the first column every time we loop
    Loop

    Set getValuesRange = Worksheets(sheet).Range("a1", Worksheets(sheet).Cells(usedRows, usedColumns).Address) ' Return the range of cells that have values
End Function

Sub runHomeCreditModel()
    
    Dim api As API_Client
    Set api = New API_Client
    Dim data As String
    
    ' Initialize API Client
    api.Initialize "https://demo.modzy.engineering/", "2GMterSomnIcXOaPyHKu.HnlucJe7iwIcj5XfpenD"
    
    ' Extract CSV file before sending to job
    data = toJSON(getValuesRange("Preprocessed Data"), False)
    
    ' Submit job
    api.call_api_home_credit_model data, "ML Predictions"
    
    'Worksheets("ML Predictions").Range("c7") = body ' Set cell B1's value to our JSON data
End Sub
