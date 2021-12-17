VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("TableUtility.Base")

Option Explicit

Private Type TTableUtility
    GivenTable As ListObject
    IsAllColumnInOutput As Boolean
End Type

Private This As TTableUtility

Public Property Get GivenTable() As ListObject
    Set GivenTable = This.GivenTable
End Property

Public Property Set GivenTable(ByVal RHS As ListObject)
    Set This.GivenTable = RHS
End Property

Public Property Get IsAllColumnInOutput() As Boolean
    IsAllColumnInOutput = This.IsAllColumnInOutput
End Property

Public Property Let IsAllColumnInOutput(ByVal RHS As Boolean)
    This.IsAllColumnInOutput = RHS
End Property

Public Function FirstHeaderColumnNo() As Long
    ThrowErrorIfListObjectIsNotSet "FirstHeaderColumnNo"
    FirstHeaderColumnNo = GivenTable.ListColumns(1).Range.Column
End Function

Public Function LastHeaderColumnNo() As Long
    ThrowErrorIfListObjectIsNotSet "LastHeaderColumnNo"
    With GivenTable
        LastHeaderColumnNo = .ListColumns(.ListColumns.Count).Range.Column
    End With
End Function

Public Function GetDataBodyOfGivenColumns(ByVal ColumnsNameOrIndex As Variant) As Variant
    
    ThrowErrorIfListObjectIsNotSet "GetDataBodyOfGivenColumns"
    ThrowErrorIfNoListRow "GetDataBodyOfGivenColumns"
    Dim GivenColumnsIndex As Collection
    Set GivenColumnsIndex = GetGivenColumnsIndex(ColumnsNameOrIndex)
    Dim Output As Variant
    Dim Counter As Long
    With GivenTable
        ReDim Output(1 To .ListRows.Count, 1 To GivenColumnsIndex.Count)
        Dim CurrentColumnIndex As Variant
        For Each CurrentColumnIndex In GivenColumnsIndex
            Counter = Counter + 1
            Dim CurrentColumnValues As Variant
            CurrentColumnValues = .ListColumns(CurrentColumnIndex).DataBodyRange.Value
            Dim FirstColumnIndex  As Long
            FirstColumnIndex = LBound(CurrentColumnValues, 2)
            Dim CurrentRowIndex As Long
            For CurrentRowIndex = LBound(CurrentColumnValues, 1) To UBound(CurrentColumnValues, 1)
                Output(CurrentRowIndex, Counter) = CurrentColumnValues(CurrentRowIndex, FirstColumnIndex)
            Next CurrentRowIndex
        Next CurrentColumnIndex
    End With
    GetDataBodyOfGivenColumns = Output
    
End Function

Private Sub ThrowErrorIfNoListRow(CallingFunctionName As String)
    If This.GivenTable.ListRows.Count = 0 Then
        Err.Raise 13, CallingFunctionName, "Table has no list row"
    End If
End Sub

Private Function GetGivenColumnsIndex(ByVal ColumnsNameOrIndex As Variant) As Variant
    
    
    With GivenTable
        Dim GivenColumnsIndex As Collection
        Set GivenColumnsIndex = New Collection
        Dim CurrentColumnIndex As Long
        Dim CurrentColumnNameOrIndex As Variant
        For Each CurrentColumnNameOrIndex In ColumnsNameOrIndex
            CurrentColumnIndex = .ListColumns(CurrentColumnNameOrIndex).Index
            GivenColumnsIndex.Add CurrentColumnIndex, CStr(CurrentColumnNameOrIndex)
        Next CurrentColumnNameOrIndex
    End With
    
    Set GetGivenColumnsIndex = GivenColumnsIndex
    
End Function

Public Function GetFilteredItem(FilterInColumns As Variant, FilterByValue As Variant, _
                                Optional OutputColumns As Variant) As Variant
    
    Dim Testers As Variant
    Testers = CreateDefaultTester(FilterByValue)
    GetFilteredItem = GetFilteredItemUsingPredicate(FilterInColumns, Testers, OutputColumns)
    
End Function

Private Function CreateDefaultTester(FilterByValue As Variant) As Variant
    
    Dim Result() As Predicate
    If IsArray(FilterByValue) Then
        ReDim Result(LBound(FilterByValue) To UBound(FilterByValue))
        Dim CurrentRowIndex As Long
        For CurrentRowIndex = LBound(FilterByValue, 1) To UBound(FilterByValue, 1)
           Set Result(CurrentRowIndex) = DefaultPredicate.Create(FilterByValue(CurrentRowIndex))
        Next CurrentRowIndex
    Else
        ReDim Result(1 To 1)
        Set Result(1) = DefaultPredicate.Create(FilterByValue)
    End If
    
    CreateDefaultTester = Result
    
End Function



Public Function GetFilteredItemUsingPredicate(FilterInColumns As Variant, Testers As Variant, _
                                Optional OutputColumns As Variant) As Variant
    
    ThrowErrorIfListObjectIsNotSet "GetFilteredItem"
    ThrowErrorIfNoListRow "GetFilteredItem"
    
    Dim FilterInData As Variant
    FilterInData = GetDataBodyOfGivenColumns(FilterInColumns)
    Testers = ConvertTo1DArray(Testers)
        
    Dim ValidRowIndexes As Collection
    Set ValidRowIndexes = New Collection
    AddAllRowInCollectionAsValid ValidRowIndexes
    If ValidRowIndexes.Count = 0 Then Exit Function
        
    FilterOutInvalidDataUsingPredicate ValidRowIndexes, FilterInData, Testers
            
    Dim OutputColumnsIndexes As Collection
    If IsAllColumnInOutput Then
        Set OutputColumnsIndexes = GetGivenColumnsIndex(GivenTable.HeaderRowRange.Value)
    Else
        Set OutputColumnsIndexes = GetGivenColumnsIndex(OutputColumns)
    End If
    
    GetFilteredItemUsingPredicate = GetGivenRowOfData(GivenTable.DataBodyRange.Value, ValidRowIndexes, OutputColumnsIndexes)
    
End Function

Private Sub FilterOutInvalidDataUsingPredicate(ByRef ValidRowIndexes As Collection, FilterInData As Variant, Testers As Variant)
    Dim CurrentValidRow As Variant
    Dim Counter As Long
    Dim CurrentColumnIndex As Long
    Dim CurrentTester As Predicate
    For Each CurrentValidRow In ValidRowIndexes
        Counter = LBound(Testers, 1)
        For CurrentColumnIndex = LBound(FilterInData, 2) To UBound(FilterInData, 2)
            Set CurrentTester = Testers(Counter)
            If Not CurrentTester.IsValid(FilterInData(CurrentValidRow, CurrentColumnIndex)) Then
                ValidRowIndexes.Remove CStr(CurrentValidRow)
                Exit For
            End If
            Counter = Counter + 1
        Next CurrentColumnIndex
    Next CurrentValidRow
End Sub


Private Sub FilterOutInvalidData(ByRef ValidRowIndexes As Collection, FilterInData As Variant, FilterByValue As Variant)
    Dim CurrentValidRow As Variant
    Dim Counter As Long
    Dim CurrentColumnIndex As Long
    For Each CurrentValidRow In ValidRowIndexes
        Counter = LBound(FilterByValue, 1)
        For CurrentColumnIndex = LBound(FilterInData, 2) To UBound(FilterInData, 2)
            If StrComp(FilterInData(CurrentValidRow, CurrentColumnIndex), FilterByValue(Counter), vbTextCompare) <> 0 Then
                ValidRowIndexes.Remove CStr(CurrentValidRow)
                Exit For
            End If
            Counter = Counter + 1
        Next CurrentColumnIndex
    Next CurrentValidRow
End Sub

Private Sub AddAllRowInCollectionAsValid(ByRef ValidRowIndexes As Collection)
    
    'Add All Row as valid
    Dim CurrentRow As Variant
    If This.GivenTable.ListRows.Count = 0 Then Exit Sub
    Dim StartRow As Long
    StartRow = LBound(GivenTable.DataBodyRange.Value, 1)
    Dim EndRow As Long
    EndRow = UBound(GivenTable.DataBodyRange.Value, 1)
    For CurrentRow = StartRow To EndRow
        ValidRowIndexes.Add CurrentRow, CStr(CurrentRow)
    Next CurrentRow
        
    
End Sub

Private Function GetGivenRowOfData(Data As Variant, ValidRowIndexes As Collection, OutputColumnsIndexes As Collection) As Variant
        
    Dim Output As Variant
    If ValidRowIndexes.Count = 0 Then
        Exit Function
    End If
    ReDim Output(1 To ValidRowIndexes.Count, 1 To OutputColumnsIndexes.Count)
    Dim CurrentRow As Long
    Dim CurrentColumn As Long
    Dim ValidRowIndex As Variant
    For Each ValidRowIndex In ValidRowIndexes
        CurrentRow = CurrentRow + 1
        CurrentColumn = 0
        Dim OutputColumnsIndex As Variant
        For Each OutputColumnsIndex In OutputColumnsIndexes
            CurrentColumn = CurrentColumn + 1
            Output(CurrentRow, CurrentColumn) = Data(ValidRowIndex, OutputColumnsIndex)
        Next OutputColumnsIndex
    Next ValidRowIndex
    GetGivenRowOfData = Output
        
End Function

'Data Should be array of 1 Column or it will flatten them into 1D array.
Private Function ConvertTo1DArray(ByVal Data As Variant) As Variant
    
    Dim Output As Variant
    If IsArray(Data) Then
        Dim NumberOfRow As Long
        NumberOfRow = UBound(Data, 1) - LBound(Data, 1) + 1
        Dim NumberOfColumn As Long
        
        'Todo: Use better error handler
        On Error Resume Next
        NumberOfColumn = UBound(Data, 2) - LBound(Data, 2) + 1
        On Error GoTo 0
        If NumberOfColumn = 0 Then NumberOfColumn = 1
        
        ReDim Output(1 To NumberOfRow * NumberOfColumn)
        Dim CurrentData As Variant
        Dim Counter As Long
        For Each CurrentData In Data
            Counter = Counter + 1
            If Not IsObject(CurrentData) Then
                 Output(Counter) = CurrentData
            Else
                Set Output(Counter) = CurrentData
            End If
        Next CurrentData
    Else
        ReDim Output(1 To 1)
        Output(1) = Data
    End If
    ConvertTo1DArray = Output
    
End Function

Private Sub ThrowErrorIfListObjectIsNotSet(CallingFunctionName As String)
    If GivenTable Is Nothing Then
        Err.Raise 13, CallingFunctionName, "Set Table first"
    End If
End Sub

Public Sub DeleteAllRowExceptFirst()
    
    With GivenTable
        ThrowErrorIfNoListRow "DeleteAllRowExceptFirst"
        Dim Counter As Long
        .DataBodyRange.ClearContents
        For Counter = .ListRows.Count To 2 Step -1
            .ListRows(Counter).Delete
        Next Counter
    End With
    
End Sub

Public Function Create(GivenTable As ListObject) As TableUtility
    
    With Me
        Set .GivenTable = GivenTable
    End With
    Set Create = Me
    
End Function

Private Function FindMinAndMax(ByVal GivenCollection As Collection) As Interval
    
    Dim Minimum As Long
    Dim Maximum As Long
    
    Dim CurrentItem As Variant
    For Each CurrentItem In GivenCollection
        If CurrentItem < Minimum Then
            Minimum = CurrentItem
        End If
        If CurrentItem > Maximum Then
            Maximum = CurrentItem
        End If
    Next CurrentItem
    Set FindMinAndMax = Interval.Create(Minimum, Maximum)
    
End Function


