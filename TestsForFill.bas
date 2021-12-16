Attribute VB_Name = "TestsForFill"
Option Explicit
Option Private Module
Public Const FILLER_TEST_TABLE As String = "FillerTestTable"
Public Const FOR_FILL_UP_TABLE As String = "ForFillUpTable"

'@TestModule
'@Folder("Tests.Filler")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("FillDown Test")
Private Sub TestFillDownMethod()
    On Error GoTo TestFail
    Dim TestOnArray As Variant
    TestOnArray = Sheet2.ListObjects(FILLER_TEST_TABLE).DataBodyRange.Value
    Dim Actual As Variant
    Actual = Filler.FillDown(TestOnArray)
    Dim Expected As Variant
    Expected = Sheet1.ListObjects(FLAVOR_TABLE).DataBodyRange.Value
    Assert.SequenceEquals Expected, Actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("FillUp Test")
Private Sub TestFillUpMethod()
    On Error GoTo TestFail
    Dim TestOnArray As Variant
    TestOnArray = Sheet2.ListObjects(FOR_FILL_UP_TABLE).DataBodyRange.Value
    Dim Actual As Variant
    Actual = Filler.FillUp(TestOnArray)
    Dim Expected As Variant
    Expected = Sheet1.ListObjects(FLAVOR_TABLE).DataBodyRange.Value
    Debug.Print IsSequenceEqual(Actual, Expected, True)
    Assert.SequenceEquals Expected, Actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

Private Function IsSequenceEqual(FirstArray As Variant, SecondArray As Variant _
                                                       , Optional PrintNotEqualIndexes As Boolean = False) As Boolean
    Dim CurrentRowIndex As Long
    Dim IsEqual As Boolean
    IsEqual = True
    For CurrentRowIndex = LBound(FirstArray, 1) To UBound(FirstArray, 1)
        Dim CurrentColumnIndex As Long
        For CurrentColumnIndex = LBound(FirstArray, 2) To UBound(FirstArray, 2)
            If FirstArray(CurrentRowIndex, CurrentColumnIndex) <> SecondArray(CurrentRowIndex, CurrentColumnIndex) Then
                IsEqual = False
                If PrintNotEqualIndexes Then
                    Debug.Print "Not Equal At : (" & CurrentRowIndex & "," & CurrentColumnIndex & ")"
                Else
                    IsSequenceEqual = IsEqual
                    Exit Function
                End If
            End If
        Next CurrentColumnIndex
    Next CurrentRowIndex
    IsSequenceEqual = IsEqual

End Function


