Attribute VB_Name = "TestForTableUtility"
Option Explicit
Public Const FLAVOR_TABLE As String = "FlavorTable"

Private Enum Fill
    Down = 1
    Up = 2
End Enum

Private Sub FillDown(GivenSheet As Worksheet, StartRowNumber As Long, EndRowNumber As Long, _
                     ColumnNumber As Long)
    With GivenSheet
        Dim RangeAddressText As String
        RangeAddressText = .Range(.Cells(StartRowNumber, ColumnNumber), .Cells(EndRowNumber, ColumnNumber)).Address
        Dim AllData As Variant
        .Range(RangeAddressText).Value = AllData
    End With
End Sub

Sub TestGetOnlyGivenColumnWhenColumnHeadingIsProvided()
    Dim Table       As ListObject
    Set Table = Sheet1.ListObjects(FLAVOR_TABLE)
    Dim Util        As TableUtility
    Set Util = TableUtility.Create(Table)
    Dim OnlyConcernedArea As Variant
    OnlyConcernedArea = Util.GetDataBodyOfGivenColumns(Array("Main Option", "Price per person"))
End Sub

Sub TestGetOnlyGivenColumnWhenColumnIndexIsProvided()
    Dim Table       As ListObject
    Set Table = Sheet1.ListObjects(FLAVOR_TABLE)
    Dim Util        As TableUtility
    Set Util = TableUtility.Create(Table)
    Dim OnlyConcernedArea As Variant
    OnlyConcernedArea = Util.GetDataBodyOfGivenColumns(Array(1, 3))
End Sub

Sub TestGetFilteredItem()
    Dim Table       As ListObject
    Set Table = Sheet1.ListObjects(FLAVOR_TABLE)
    Dim Util        As TableUtility
    Set Util = TableUtility.Create(Table)
    Dim OnlyConcernedArea As Variant
    OnlyConcernedArea = Util.GetFilteredItem(Array("Main Option"), Array("Cup cakes"), Array("Flavours", "Price per person"))
End Sub


