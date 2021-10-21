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
Option Explicit

Private Type TTableUtility
    GivenListobject As ListObject
    IsAllColumnInOutput As Boolean
End Type

Private This As TTableUtility

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

Public Function GetDataBodyOfGivenColumns(ByVal ColumnNamesOrIndexes As Variant) As Variant
    
    ThrowErrorIfListObjectIsNotSet "getDataBodyOfGivenColumns"
    Dim GivenColumnsIndex As Collection
    Set GivenColumnsIndex = GetGivenColumnsIndex(ColumnNamesOrIndexes)
    Dim Output As Variant
    Dim Counter As Long
    With This.GivenListobject
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

Private Function GetGivenColumnsIndex(ByVal ColumnNamesOrIndexes As Variant)
    
    Dim GivenColumnsIndex As Collection
    With This.GivenListobject
        Set GivenColumnsIndex = New Collection
        Dim CurrentColumnIndex As Long
        Dim CurrentColumnNameOrIndex As Variant
        For Each CurrentColumnNameOrIndex In ColumnNamesOrIndexes
            CurrentColumnIndex = .ListColumns(CurrentColumnNameOrIndex).Index
            GivenColumnsIndex.Add CurrentColumnIndex, CStr(CurrentColumnNameOrIndex)
        Next CurrentColumnNameOrIndex
    End With
    
    Set GetGivenColumnsIndex = GivenColumnsIndex
    
End Function


'TODO: Need to modify this one.
Public Function GetFilteredItem(FilterInColumns As Variant, FilterByValue As Variant, _
                                Optional OutputColumns As Variant) As Variant
    
    ThrowErrorIfListObjectIsNotSet "GetFilteredItem"
    
    Dim GivenColumnsIndex As Collection
    Set GivenColumnsIndex = GetGivenColumnsIndex(FilterInColumns)
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
            For Each CurrentColumnIndex In GivenColumnsIndex
                
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
        Set OutputColumnsIndexes = GetGivenColumnsIndex(This.GivenListobject.HeaderRowRange.Value)
    Else
        Set OutputColumnsIndexes = GetGivenColumnsIndex(OutputColumns)
    End If
    
    GetFilteredItem = GetGivenRowOfData(This.GivenListobject.DataBodyRange.Value, ValidRowIndexes, OutputColumnsIndexes)
    
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

Public Sub DeleteAllRowExceptFirst()
    With This.GivenListobject
        Dim Counter As Long
        For Counter = .ListRows.Count To 2 Step -1
            .ListRows(Counter).Delete
        Next Counter
    End With
End Sub

Public Function Create(GivenTable As ListObject) As TableUtility
    
    With Me
        Set .GivenListobject = GivenTable
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


