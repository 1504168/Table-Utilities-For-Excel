Attribute VB_Name = "TestTableUtility"
Option Explicit
Public Const SALES_DATA_TABLE As String = "SalesDataTable"

Private Sub DoTest()
    
    'Test on This table
    Dim Table As ListObject
    Set Table = FilterTestDataSheet.ListObjects(SALES_DATA_TABLE)
    
    'Create Utility Object
    Dim Util As TableUtility
    Set Util = TableUtility.Create(Table)
    
    'Creating Date Tester
    Dim DateTester As Predicate
    Set DateTester = TesterCreator.Create(Operator.IN_BETWEEN, #1/1/2020#, #12/31/2020#)
    
    'Creating Text Contains tester
    Dim RegionTester As Predicate
    Set RegionTester = TesterCreator.Create(Operator.CONTAINS, "EA")
    
    'Creating Greater Than Tester
    Dim UnitCostTester As Predicate
    Set UnitCostTester = TesterCreator.Create(Operator.GREATER_THAN, 100)
    
    
    Dim FilteredData As Variant
    Util.IsAllColumnInOutput = True
    
    'Filter Data
    FilteredData = Util.GetFilteredItemUsingPredicate(Array("Order Date", "Region", "Unit Cost"), Array(DateTester, RegionTester, UnitCostTester))
    
    'Manully inspected number of row is 15
    Dim NumberOfRow As Long
    NumberOfRow = UBound(FilteredData, 1) - LBound(FilteredData, 1) + 1
    Debug.Assert NumberOfRow = 15
    MsgBox "Test Pass"
    
End Sub

