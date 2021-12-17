Attribute VB_Name = "TableUtilityTest"
Option Explicit
Option Private Module
Public Const SALES_DATA_TABLE As String = "SalesDataTable"

'& is denoting that a number is long. Assert will return inconclusive if type are not same.

'@TestModule
'@Folder("Tests")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider
Private Util As TableUtility

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
    
    Dim Table As ListObject
    Set Table = FilterTestDataSheet.ListObjects(SALES_DATA_TABLE)
    'Create Utility Object
    Set Util = TableUtility.Create(Table)
    
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    Set Util = Nothing
    
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    Util.IsAllColumnInOutput = False
End Sub


'@TestMethod("Test GetFilteredItemUsingPredicate")
Private Sub TestUsingPredicate()
    
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
    
    Assert.AreEqual 15&, NumberOfRow
    
End Sub

'@TestMethod("Test GetFilteredItem")
Private Sub TestWithoutPredicate()

    Dim FilteredData As Variant
    Util.IsAllColumnInOutput = True
    
    'Filter Data
    FilteredData = Util.GetFilteredItem(Array("Item Type", "Order Priority"), Array("Clothes", "L"))
    
    'Manully inspected number of row is 23
    Dim NumberOfRow As Long
    NumberOfRow = UBound(FilteredData, 1) - LBound(FilteredData, 1) + 1
    Debug.Print NumberOfRow
    
    
    Assert.AreEqual 23&, NumberOfRow
    
End Sub

