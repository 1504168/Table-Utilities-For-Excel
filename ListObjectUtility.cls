VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListObjectUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private Type TListObjectUtility
    GivenListobject As ListObject
    IsAllColumnInOutput As Boolean
End Type

Private This As TListObjectUtility

Public Property Get IsAllColumnInOutput() As Boolean
    IsAllColumnInOutput = This.IsAllColumnInOutput
End Property

Public Property Let IsAllColumnInOutput(ByVal RHS As Boolean)
    This.IsAllColumnInOutput = RHS
End Property

Public Property Set GivenListobject(ByVal RHS As ListObject)
    Set This.GivenListobject = RHS
End Property

Public Function FirstHeaderColumnNo() As Long
    ThrowErrorIfListObjectIsNotSet "FirstHeaderColumnNo"
    FirstHeaderColumnNo = This.GivenListobject.ListColumns(1).Range.Column
End Function

Public Function LastHeaderColumnNo() As Long
    ThrowErrorIfListObjectIsNotSet "LastHeaderColumnNo"
    With This.GivenListobject
        LastHeaderColumnNo = This.GivenListobject.ListColumns(.ListColumns.Count).Range.Column
    End With
End Function

Public Function getDataBodyOfGivenColumns(ByVal ColumnNames As Variant) As Variant
    
    ThrowErrorIfListObjectIsNotSet "getDataBodyOfGivenColumns"
    
    Dim givenColumnsIndex As Collection
    Set givenColumnsIndex = getGivenColumnsIndex(ColumnNames)
    Dim CurrentRow As Long
    With This.GivenListobject
        Dim Output As Variant
        ReDim Output(1 To .ListRows.Count, 1 To givenColumnsIndex.Count)
        For CurrentRow = LBound(.DataBodyRange.Value, 1) To UBound(.DataBodyRange.Value, 1)
            Dim Counter As Long
            Counter = 0
            Dim CurrentColumnIndex As Variant
            For Each CurrentColumnIndex In givenColumnsIndex
                Counter = Counter + 1
                Output(CurrentRow, Counter) = .DataBodyRange(CurrentRow, CurrentColumnIndex).Value
            Next CurrentColumnIndex
        Next CurrentRow
    End With
    getDataBodyOfGivenColumns = Output
    
End Function

Private Function getGivenColumnsIndex(ByVal ColumnNames As Variant)
    
    Dim givenColumnsIndex As Collection
    With This.GivenListobject
        Set givenColumnsIndex = New Collection
        Dim CurrentColumnIndex As Long
        If IsArray(ColumnNames) Then
            Dim CurrentColumnName As Variant
            For Each CurrentColumnName In ColumnNames
                CurrentColumnIndex = .ListColumns(CurrentColumnName).Index
                givenColumnsIndex.Add CurrentColumnIndex, CurrentColumnName
            Next CurrentColumnName
        Else
            CurrentColumnIndex = .ListColumns(ColumnNames).Index
            givenColumnsIndex.Add CurrentColumnIndex, ColumnNames
        End If
    End With
    
    Set getGivenColumnsIndex = givenColumnsIndex
    
End Function

Public Function getFilteredItem(FilterInColumns As Variant, FilterByValue As Variant, _
                                Optional OutputColumns As Variant) As Variant
    
    ThrowErrorIfListObjectIsNotSet "getFilteredItem"
    
    Dim givenColumnsIndex As Collection
    Set givenColumnsIndex = getGivenColumnsIndex(FilterInColumns)
    FilterByValue = ConvertTo1DArray(FilterByValue)

    With This.GivenListobject
        
        Dim ValidRowIndexes As Collection
        Set ValidRowIndexes = New Collection
        AddAllRowInCollectionAsValid ValidRowIndexes
        
        Dim CurrentValidRow As Variant
        For Each CurrentValidRow In ValidRowIndexes
            Dim Counter As Long
            Counter = LBound(FilterByValue, 1)
            Dim CurrentColumnIndex As Variant
            For Each CurrentColumnIndex In givenColumnsIndex
                
                Dim IsValidRow As Boolean
                IsValidRow = True
                
                If .DataBodyRange(CurrentValidRow, CurrentColumnIndex).Value <> FilterByValue(Counter) Then
                    IsValidRow = False
                End If
                Counter = Counter + 1
                
                If Not IsValidRow Then
                    ValidRowIndexes.Remove CStr(CurrentValidRow)
                    Exit For
                End If
                
            Next CurrentColumnIndex
            
        Next CurrentValidRow

    End With
    
    Dim OutputColumnsIndexes As Collection
    If This.IsAllColumnInOutput Then
        Set OutputColumnsIndexes = getGivenColumnsIndex(This.GivenListobject.HeaderRowRange.Value)
    Else
        Set OutputColumnsIndexes = getGivenColumnsIndex(OutputColumns)
    End If
    
    getFilteredItem = GetGivenRowOfData(This.GivenListobject.DataBodyRange.Value, ValidRowIndexes, OutputColumnsIndexes)
    
End Function

Private Sub AddAllRowInCollectionAsValid(ByRef ValidRowIndexes As Collection)
    
    With This.GivenListobject
        'Add All Row as valid
        Dim CurrentRow As Variant
        For CurrentRow = LBound(.DataBodyRange.Value, 1) To UBound(.DataBodyRange.Value, 1)
            ValidRowIndexes.Add CurrentRow, CStr(CurrentRow)
        Next CurrentRow
        
    End With
    
End Sub

Private Function GetGivenRowOfData(Data As Variant, ValidRowIndexes As Collection, OutputColumnsIndexes As Collection) As Variant
        
    Dim Output As Variant
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

'GivenData Should be array of 1 Column or it will flatten them into 1D array.
Private Function ConvertTo1DArray(ByVal GivenData As Variant)
    
    Dim Output As Variant
    If IsArray(GivenData) Then
        Dim NumberOfRow As Long
        NumberOfRow = UBound(GivenData, 1) - LBound(GivenData, 1) + 1
        Dim NumberOfColumn As Long
        
        'Todo: Use better error handler
        On Error Resume Next
        NumberOfColumn = UBound(GivenData, 2) - LBound(GivenData, 2) + 1
        On Error GoTo 0
        If NumberOfColumn = 0 Then NumberOfColumn = 1
        
        ReDim Output(1 To NumberOfRow * NumberOfColumn)
        Dim CurrentData As Variant
        Dim Counter As Long
        For Each CurrentData In GivenData
            Counter = Counter + 1
            Output(Counter) = CurrentData
        Next CurrentData
    Else
        ReDim Output(1 To 1)
        Output(1) = GivenData
    End If
    ConvertTo1DArray = Output
    
End Function

Private Sub ThrowErrorIfListObjectIsNotSet(CallingFunctionName As String)
    If This.GivenListobject Is Nothing Then
        Err.Raise 13, CallingFunctionName, "Set Table first"
    End If
End Sub


