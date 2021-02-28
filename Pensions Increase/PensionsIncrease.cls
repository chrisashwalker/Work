VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class PI

' Declare workbook and worksheets.

Private ThisBook As Workbook
Private ControlSheet As Worksheet
Private PercentagesSheet As Worksheet
Private WageTypeSheet As Worksheet
Private InputSheet As Worksheet
Private OutputSheet As Worksheet

' Declare tables.

Private Percentages As ListObject
Private WageTypes As ListObject
Private InputTable As ListObject
Private OutputTable As ListObject

' Declare dictionary objects.
' Wage type dictionary keys are wage type numbers; dictionary items are their PI versions.
' Rates dictionary keys are personnel numbers; dictionary items are their PI rates based on deemed to begin dates in infotype 41.
' IncreaseList dictionary keys are incremented with each new addition; dictionary items are custom Increase objects.

Private FinalSalaryWageTypes As Dictionary
Private CareerAverageWageTypes As Dictionary
Private GMPWageTypes As Dictionary
Private DependantOrIllHealthWageTypes As Dictionary

' Declare variables from the Control worksheet.

Private GMPDate As Date
Private PIDate As Date
Private GMPRate As Double
Private PercentagesFilled As Boolean
Private WageTypesFilled As Boolean
Private InputFilled As Boolean

' Declare additional variables.

Private PIEndDate As Date
Private GMPEndDate As Date
Private EndOfTime As Date
Private Order As String
Private FSRate As Double
Private CARate As Double
Private NullValue As Integer

Private FirstWageTypeColumn As Integer
Private LastWageTypeColumn As Integer
Private ColumnsFromWageTypeToAmount As Integer
Private ColumnsFromWageTypeToWageType As Integer
Private FinalSalaryDTBColumn As Integer
Private CareerAverageDTBColumn As Integer
Private AgeColumn As Integer


Private Sub SetVariables()

    Set ThisBook = Application.ActiveWorkbook
    Set ControlSheet = ThisBook.Worksheets("Control")
    Set PercentagesSheet = ThisBook.Worksheets("Percentages")
    Set WageTypesSheet = ThisBook.Worksheets("WageTypes")
    Set InputSheet = ThisBook.Worksheets("Input")
    Set OutputSheet = ThisBook.Worksheets("Output")
    
    Set Percentages = PercentagesSheet.ListObjects("Percentages")
    Set WageTypes = WageTypesSheet.ListObjects("WageTypes")
    Set InputTable = InputSheet.ListObjects("InputTable")
    Set FinalSalaryWageTypes = CreateObject("Scripting.Dictionary")
    Set CareerAverageWageTypes = CreateObject("Scripting.Dictionary")
    Set GMPWageTypes = CreateObject("Scripting.Dictionary")
    Set DependantOrIllHealthWageTypes = CreateObject("Scripting.Dictionary")
    Set WageIncreaseList = CreateObject("Scripting.Dictionary")
    
    PercentagesFilled = ControlSheet.Range("PercentagesFilled").Value
    WageTypesFilled = ControlSheet.Range("WageTypesFilled").Value
    InputFilled = ControlSheet.Range("InputFilled").Value
    GMPRate = ControlSheet.Range("GMPRate").Value
    EndOfTime = DateValue("Dec 31, 9999")

    NullValue = -99
    
    FirstWageTypeColumn = 4
    LastWageTypeColumn = 18
    ColumnsFromWageTypeToAmount = 1
    ColumnsFromWageTypeToWageType = 2
    FinalSalaryDTBColumn = 20
    CareerAverageDTBColumn = 21
    AgeColumn = 22
    
End Sub


Private Sub SetDates()    ' Set increase validity end dates.
    
    GMPDate = ControlSheet.Range("GMPDate").Value
    PIDate = ControlSheet.Range("PIDate").Value
    
    If IsEmpty(GMPDate) Then
        If IsEmpty(PIDate) Then
            Order = "NoIncreases"
        Else
            PIEndDate = EndOfTime
            Order = "PIOnly"
        End If
    ElseIf IsEmpty(PIDate) Then
        GMPEndDate = EndOfTime
        Order = "GMPOnly"
    ElseIf PIDate > GMPDate Then
        PIEndDate = EndOfTime
        GMPEndDate = PIDate - 1
        Order = "GMPFirst"
    ElseIf PIDate < GMPDate Then
        PIEndDate = GMPDate - 1
        GMPEndDate = EndOfTime
        Order = "PIFirst"
    Else
        PIEndDate = EndOfTime
        GMPEndDate = EndOfTime
        Order = "Both"
    End If
    
End Sub


Private Sub CheckInput()    ' Check if the program has sufficient data to run.

    Dim InsufficientData
    If (Not IsEmpty(GMPDate) And IsEmpty(GMPRate)) Or (Not IsEmpty(PIDate) And Not PercentagesFilled) Or Not WageTypesFilled Or Not InputFilled Then
        InsufficientData = MsgBox("The program is missing vital data and cannot yet run", vbOKOnly)
        If InsufficientData = vbOK Then
            Exit Sub
        End If
    End If
    
End Sub


Private Sub FindWageTypes()    ' Build the dictionary of relevant wage types and their PI versions from the WageTypes table.

    Dim WageType, WageTypeCategory, IncreaseWageType, DependantOrIllHealth
    For Each WageTypesRow In WageTypes.ListRows
    
        WageType = WageTypesRow.Range(1, 1).Value
        WageTypeCategory = WageTypesRow.Range(1, 2).Value
        IncreaseWageType = WageTypesRow.Range(1, 3).Value
        DependantOrIllHealth = WageTypesRow.Range(1, 4).Value
        
        If WageTypeCategory = "FS" And Not IsEmpty(IncreaseWageType) Then
            FinalSalaryWageTypes.Add WageType, IncreaseWageType
        ElseIf WageTypeCategory = "CA" And Not IsEmpty(IncreaseWageType) Then
            CareerAverageWageTypes.Add WageType, IncreaseWageType
        ElseIf WageTypeCategory = "GMP" And Not IsEmpty(IncreaseWageType) Then
            GMPWageTypes.Add WageType, IncreaseWageType
        End If
        
        If DependantOrIllHealth Then
            DependantOrIllHealthWageTypes.Add WageType, IncreaseWageType
        End If
    Next
    
End Sub


Private Sub ClearOutput()    ' Clear the output worksheet and build a new output table as a copy of the input table.

    OutputSheet.Cells.Clear
    InputSheet.Cells.Copy Destination:=OutputSheet.Range("A1")
    Set OutputTable = OutputSheet.ListObjects(1)
    OutputTable.Name = "Output"
    
End Sub


Private Sub SetRates(DataRow As ListRow)    ' Determine the Final Salary and Career Average rates of increase for the record based on its Deemed To Begin date

    Dim FinalSalaryDTB, CareerAverageDTB, PercentageValidFrom, PercentageValidTo, PercentageRate
    FinalSalaryDTB = DataRow.Range(1, FinalSalaryDTBColumn).Value
    CareerAverageDTB = DataRow.Range(1, CareerAverageDTBColumn).Value
    
    For Each PercentagesRow In Percentages.ListRows
    
        PercentageValidFrom = PercentagesRow.Range(1, 1).Value
        PercentageValidTo = PercentagesRow.Range(1, 2).Value
        PercentageRate = PercentagesRow.Range(1, 3).Value
    
        If Not IsEmpty(FinalSalaryDTB) And PercentageValidFrom <= FinalSalaryDTB And PercentageValidTo >= FinalSalaryDTB Then
            FSRate = PercentageRate
        End If
        If Not IsEmpty(CareerAverageDTB) And PercentageValidFrom <= CareerAverageDTB And PercentageValidTo >= CareerAverageDTB Then
            CARate = PercentageRate
        End If
        If FSRate <> NullValue And CARate <> NullValue Then
            Exit For
        End If
        
    Next
    
End Sub


Private Sub EvaluateWagePairs(DataRow As ListRow, WageTypeList As Dictionary, IncreaseList As Dictionary, Rate As Double)   ' Loop through the wage types and amounts stored against the record, calculating their increases and storing in a list

    Dim WageType
    Dim Amount
    Dim DependantOrIllHealth As Boolean
    Dim Age
    Dim GMPType As Boolean
    Dim Error As Range
    Set Error = DataRow.Range(1, 2)
    DependantOrIllHealth = False
    Age = DataRow.Range(1, AgeColumn).Value
    GMPType = False
    For ColumnCount = FirstWageTypeColumn To LastWageTypeColumn Step ColumnsFromWageTypeToWageType
    
        WageType = DataRow.Range(1, ColumnCount).Value
        Amount = DataRow.Range(1, ColumnCount + ColumnsFromWageTypeToAmount).Value
        
        If DependantOrIllHealthWageTypes.Exists(WageType) Then
            DependantOrIllHealth = True
        End If
        
        If WageTypeList.Exists(WageType) Then
        
            If GMPWageTypes.Exists(WageType) Then
                GMPType = True
            End If
        
            If IsEmpty(Rate) Or Rate = NullValue Then
                Error.Value = "RATE NOT FOUND"
                Exit For
            End If
            
            If IncreaseList.Exists(WageTypeList.Item(WageType)) Then
                IncreaseList.Item(WageTypeList.Item(WageType)) = IncreaseList.Item(WageTypeList.Item(WageType)) + (Amount * Rate)
            Else
                IncreaseList.Add WageTypeList.Item(WageType), Amount * Rate
            End If
            
        End If
    Next
    
    If Not GMPType And Age < 55 And Not DependantOrIllHealth Then
        Error.Value = "NOT ELIGIBLE"
        IncreaseList.RemoveAll
    End If
    
End Sub


Private Sub ApplyIncreases(DataRow As ListRow, IncreaseList As Dictionary)   ' Apply the list of increases to the output table

    Dim IncreaseWageType As Variant
    Dim WageTypeFound As Boolean
    Dim WageTypeRange As Range
    Dim AmountRange As Range
    Dim WageType
    Dim Amount
    Dim Error As Range
    Set Error = DataRow.Range(1, 2)
    
    For Each IncreaseWageType In IncreaseList.Keys
        WageTypeFound = False
        For ColumnCount = FirstWageTypeColumn To LastWageTypeColumn Step ColumnsFromWageTypeToWageType
            Set WageTypeRange = DataRow.Range(1, ColumnCount)
            Set AmountRange = DataRow.Range(1, ColumnCount + ColumnsFromWageTypeToAmount)
            WageType = WageTypeRange.Value
            Amount = AmountRange.Value
            
            If WageType = IncreaseWageType Then
                AmountRange.Value = Amount + IncreaseList.Item(IncreaseWageType)
                Exit For
            ElseIf IsEmpty(WageType) Then
                WageTypeRange.Value = IncreaseWageType
                AmountRange.Value = IncreaseList.Item(IncreaseWageType)
                Exit For
            ElseIf ColumnCount = LastWageTypeColumn Then
                Error.Value = "MAXIMUM WAGE PAIRS EXCEEDED"
            End If
        Next
    Next
    
End Sub


Private Sub TriggerIncreasesInOrder(ErrorRange As Range, DataRow As ListRow, NewDataRow As ListRow, NewDataRowStartDateRange As Range, NewDataRowEndDateRange As Range, NewDataRowError As Range, IncreaseList As Dictionary, FRate As Double, CRate As Double)    ' Sets off the first increase, GMP or PI, in accordance with the predetermined order; then creates a new row for the second increase, where appropriate
    If ErrorRange.Value <> "RATE NOT FOUND" Then
        Call ApplyIncreases(DataRow, IncreaseList)
        If ErrorRange.Value <> "MAXIMUM WAGE PAIRS EXCEEDED" Then
            Set NewDataRow = OutputTable.ListRows.Add
            DataRow.Range.Cells.Copy (OutputTable.DataBodyRange(NewDataRow.Index, 1))
            Set NewDataRowStartDateRange = NewDataRow.Range(1, 2)
            Set NewDataRowEndDateRange = NewDataRow.Range(1, 3)
            Set NewDataRowError = NewDataRow.Range(1, 2)
            IncreaseList.RemoveAll
            If Order = "GMPFirst" Then
                NewDataRowStartDateRange.Value = PIDate
                NewDataRowEndDateRange.Value = PIEndDate
                Call EvaluateWagePairs(NewDataRow, FinalSalaryWageTypes, IncreaseList, FRate)
                Call EvaluateWagePairs(NewDataRow, CareerAverageWageTypes, IncreaseList, CRate)
            ElseIf Order = "PIFirst" Then
                NewDataRowStartDateRange.Value = GMPDate
                NewDataRowEndDateRange.Value = GMPEndDate
                Call EvaluateWagePairs(NewDataRow, GMPWageTypes, IncreaseList, GMPRate)
            End If
            If Order <> "GMPOnly" And Order <> "PIOnly" And NewDataRowError.Value <> "RATE NOT FOUND" Then
                Call ApplyIncreases(NewDataRow, IncreaseList)
            End If
        End If
    End If
End Sub
                

Private Sub RecordLoop() ' Loops through each record and triggers the increase calculations

    Dim WageIncreaseList As Dictionary
    Set WageIncreaseList = CreateObject("Scripting.Dictionary")
    Dim ColumnCount As Integer
    Dim OutputRow As ListRow
    Dim NewRow As ListRow
    Dim OutputRowStartRange As Range
    Dim OutputRowEndRange As Range
    Dim NewRowStartRange As Range
    Dim NewRowEndRange As Range
    Dim Error As Range
    Dim NewRowError As Range

    If Order = "NoIncreases" Then
        OutputTable.DataBodyRange.Clear
        Exit Sub
    Else
        For Each OutputRow In OutputTable.ListRows
        
            FSRate = NullValue
            CARate = NullValue
            Call SetRates(OutputRow)
            WageIncreaseList.RemoveAll
            Set OutputRowStartRange = OutputRow.Range(1, 2)
            Set OutputRowEndRange = OutputRow.Range(1, 3)
            Set Error = OutputRow.Range(1, 2)
            
            If OutputRowEndRange.Value < GMPDate And OutputRowEndRange.Value < PIDate Then
                Error.Value = "RECORD ENDS BEFORE INCREASES"
            ElseIf OutputRowEndRange.Value <> EndOfTime Then
                Error.Value = "RECORD HAS A PROPER END DATE"
            ElseIf Order = "GMPFirst" And GMPDate < OutputRowEndRange.Value Then
                OutputRowStartRange.Value = GMPDate
                OutputRowEndRange.Value = GMPEndDate
                Call EvaluateWagePairs(OutputRow, GMPWageTypes, WageIncreaseList, GMPRate)
                Call TriggerIncreasesInOrder(Error, OutputRow, NewRow, NewRowStartRange, NewRowEndRange, NewRowError, WageIncreaseList, FSRate, CARate)
            ElseIf Order = "PIFirst" And PIDate < OutputRowEndRange.Value Then
                OutputRowStartRange.Value = PIDate
                OutputRowEndRange.Value = PIEndDate
                Call EvaluateWagePairs(OutputRow, FinalSalaryWageTypes, WageIncreaseList, FSRate)
                Call EvaluateWagePairs(OutputRow, CareerAverageWageTypes, WageIncreaseList, CARate)
                Call TriggerIncreasesInOrder(Error, OutputRow, NewRow, NewRowStartRange, NewRowEndRange, NewRowError, WageIncreaseList, FSRate, CARate)
            ElseIf Order = "Both" And GMPDate < OutputRowEndRange.Value Then
                OutputRowStartRange.Value = GMPDate
                OutputRowEndRange.Value = GMPEndDate
                Call EvaluateWagePairs(OutputRow, GMPWageTypes, WageIncreaseList, GMPRate)
                Call EvaluateWagePairs(OutputRow, FinalSalaryWageTypes, WageIncreaseList, FSRate)
                Call EvaluateWagePairs(OutputRow, CareerAverageWageTypes, WageIncreaseList, CARate)
                If ErrorRange.Value <> "RATE NOT FOUND" Then
                    Call ApplyIncreases(OutputRow, WageIncreaseList)
                End If
            Else
                Error.Value = "INCREASE OUTSIDE OF RECORD VALIDITY"
            End If
            
            If NewRowError.Value = "NOT ELIGIBLE" Then
                OutputRowEndRange.Value = EndOfTime
            End If
        Next
        
    End If
    
End Sub


Public Sub Process()    ' Launches the entire Guaranteed Minimum Pension and regular Pension Increase process

    SetVariables
    SetDates
    CheckInput
    FindWageTypes
    ClearOutput
    RecordLoop
    
End Sub






