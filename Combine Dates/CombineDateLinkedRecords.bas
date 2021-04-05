Attribute VB_Name = "Process"
Public Sub Process()
    
    Dim ThisBook As Workbook
    Dim ControlSheet As Worksheet
    Dim OutputSheet As Worksheet
    Dim OutputTable As ListObject

    Set ThisBook = Application.ActiveWorkbook
    Set ControlSheet = ThisBook.Worksheets("CombineAddInControl")
    Set OutputSheet = ThisBook.Worksheets("CombineAddInOutput")
    
    Dim InputTableName
    Dim IDColumnName
    Dim StartDateColumnName
    Dim EndDateColumnName
    Dim TypeColumnName
    Dim IDColumn As ListColumn
    Dim StartDateColumn As ListColumn
    Dim EndDateColumn As ListColumn
    Dim TypeColumn As ListColumn
    CheckType = False
    Dim SplitsTable
    Dim TransformsTable
    Dim Gap As Integer
    Dim IncludeWeekends As Boolean
    InputTableName = ControlSheet.Range("TableName").Value
    IDColumnName = ControlSheet.Range("ID").Value
    StartDateColumnName = ControlSheet.Range("Start").Value
    EndDateColumnName = ControlSheet.Range("End").Value
    TypeColumnName = ControlSheet.Range("Type").Value
    Gap = ControlSheet.Range("Gap").Value
    Set SplitsTable = ControlSheet.ListObjects("Splits")
    Set TransformsTable = ControlSheet.ListObjects("Transforms")
    Dim Splits As Dictionary
    Dim Transforms As Dictionary
    Set Splits = CreateObject("Scripting.Dictionary")
    Set Transforms = CreateObject("Scripting.Dictionary")
    
    Dim InputTable As ListObject
    
    If ControlSheet.Range("IncludeWeekends").Value = "Yes" Then
        IncludeWeekends = True
    Else
        IncludeWeekends = False
    End If
    
    Dim booksheet As Worksheet
    Dim sheettable As ListObject
    For Each booksheet In ThisBook.Worksheets
        If InputTable Is Nothing Then
            For Each sheettable In booksheet.ListObjects
                If sheettable.Name = InputTableName Then
                    Set InputTable = sheettable
                    Exit For
                End If
            Next
        Else
            Exit For
        End If
    Next
    
    Dim tablerow As ListRow
    
    For Each tablerow In SplitsTable.ListRows
        If tablerow.Range(1, 1).Value <> "" And Not Splits.Exists(tablerow.Range(1, 1).Value) Then
            Splits.Add tablerow.Range(1, 1).Value, 0
        End If
    Next
    
    For Each tablerow In TransformsTable.ListRows
        If tablerow.Range(1, 1).Value <> "" And Not Transforms.Exists(tablerow.Range(1, 1).Value) Then
            If tablerow.Range(1, 2).Value <> "" Then
                Transforms.Add tablerow.Range(1, 1).Value, tablerow.Range(1, 2).Value
            Else
                Transforms.Add tablerow.Range(1, 1).Value, "SUM"
            End If
        End If
    Next
    
    OutputSheet.Cells.Clear
    InputTable.Range.Copy Destination:=OutputSheet.Range("A1")
    Set OutputTable = OutputSheet.ListObjects(1)
    OutputTable.Name = "Output"
    
    Set IDColumn = OutputTable.ListColumns(IDColumnName)
    Set StartDateColumn = OutputTable.ListColumns(StartDateColumnName)
    Set EndDateColumn = OutputTable.ListColumns(EndDateColumnName)
    If TypeColumnName <> "" Then
        Set TypeColumn = OutputTable.ListColumns(TypeColumnName)
    End If
    
    For Each tablerow In OutputTable.ListRows
        StartDateColumn.DataBodyRange(tablerow.Index, 1).Value = DateValue(Replace(StartDateColumn.DataBodyRange(tablerow.Index, 1).Value, ".", "/"))
        EndDateColumn.DataBodyRange(tablerow.Index, 1).Value = DateValue(Replace(EndDateColumn.DataBodyRange(tablerow.Index, 1).Value, ".", "/"))
    Next
    
    With OutputTable.Sort
        .SortFields.Clear
        .SortFields.Add IDColumn.Range, xlSortOnValues, xlAscending
        If Not TypeColumn Is Nothing Then
            .SortFields.Add TypeColumn.Range, xlSortOnValues, xlAscending
        End If
        .SortFields.Add StartDateColumn.Range, xlSortOnValues, xlAscending
        .Header = xlYes
        .Apply
        .SortFields.Clear
    End With
    
    If IsEmpty(Gap) Then
        Gap = 1
    End If
    
    Dim nexttablerow As ListRow
    
    For Each tablerow In OutputTable.ListRows
        If tablerow.Index < OutputTable.ListRows.Count Then
            Set nexttablerow = OutputTable.ListRows(tablerow.Index + 1)
            If TypeColumn Is Nothing Then
                If tablerow.Range(1, IDColumn.Index).Value = nexttablerow.Range(1, IDColumn.Index).Value _
                And (tablerow.Range(1, EndDateColumn.Index).Value >= nexttablerow.Range(1, StartDateColumn.Index).Value - Gap _
                Or ((Weekday(DateValue(tablerow.Range(1, EndDateColumn.Index).Value)) = 6 And Weekday(DateValue(nexttablerow.Range(1, StartDateColumn.Index).Value)) = 2) And IncludeWeekends = True)) Then
                    Call Combine(OutputTable, StartDateColumn, Splits, Transforms, tablerow, nexttablerow)
                End If
            Else
                If tablerow.Range(1, IDColumn.Index).Value = nexttablerow.Range(1, IDColumn.Index).Value _
                And (tablerow.Range(1, EndDateColumn.Index).Value >= nexttablerow.Range(1, StartDateColumn.Index).Value - Gap _
                Or ((Weekday(DateValue(tablerow.Range(1, EndDateColumn.Index).Value)) = 6 And Weekday(DateValue(nexttablerow.Range(1, StartDateColumn.Index).Value)) = 2) And IncludeWeekends = True)) _
                And tablerow.Range(1, TypeColumn.Index).Value = nexttablerow.Range(1, TypeColumn.Index).Value Then
                    Call Combine(OutputTable, StartDateColumn, Splits, Transforms, tablerow, nexttablerow)
                End If
            End If
        End If
    Next
    
    With OutputTable.Sort
        .SortFields.Clear
        .SortFields.Add IDColumn.Range, xlSortOnValues, xlAscending
        If Not TypeColumn Is Nothing Then
            .SortFields.Add TypeColumn.Range, xlSortOnValues, xlAscending
        End If
        .SortFields.Add StartDateColumn.Range, xlSortOnValues, xlAscending
        .Header = xlYes
        .Apply
        .SortFields.Clear
    End With
    
    Dim rowcount
    rowcount = 1
    Do While rowcount <= OutputTable.ListRows.Count
        If IsEmpty(OutputTable.DataBodyRange(rowcount, IDColumn.Index).Value) Then
            OutputTable.ListRows(rowcount).Delete
        Else
            rowcount = rowcount + 1
        End If
    Loop
    
    OutputSheet.Activate
    OutputSheet.Range("A1").Activate
    
End Sub

Private Sub Combine(OutputTable As ListObject, StartDateColumn As ListColumn, Splits As Dictionary, Transforms As Dictionary, tablerow As ListRow, nexttablerow As ListRow)
    Dim split, transform
    For Each split In Splits
        If nexttablerow.Range(1, OutputTable.ListColumns(split).Index).Value <> tablerow.Range(1, OutputTable.ListColumns(split).Index).Value Then
            Exit Sub
        End If
    Next
    For Each transform In Transforms
        nexttablerow.Range(1, OutputTable.ListColumns(transform).Index).Value = nexttablerow.Range(1, OutputTable.ListColumns(transform).Index).Value + tablerow.Range(1, OutputTable.ListColumns(transform).Index).Value
    Next
    nexttablerow.Range(1, StartDateColumn.Index).Value = tablerow.Range(1, StartDateColumn.Index).Value
    tablerow.Range.Clear
End Sub

Public Sub ImportSheets()
    Dim UserBook
    Set UserBook = Application.ActiveWorkbook
    Dim AddInBook As Workbook
    Set AddInBook = Application.Workbooks("CombineDateLinkedRecords.xlsm")
    With AddInBook
        .Worksheets(Array("CombineAddInControl", "CombineAddInOutput")).Copy Before:=UserBook.Worksheets(1)
    End With
    AddInBook.Close SaveChanges:=False
End Sub
