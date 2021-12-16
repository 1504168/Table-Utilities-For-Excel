VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TesterCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("Predicate.Creator")

Option Explicit

Public Enum Operator
    
    Equal = 1
    DOES_NOT_EQUAL = 2
    GREATER_THAN = 3
    GREATER_THAN_OR_EQUAL_TO = 4
    LESS_THAN = 5
    LESS_THAN_OR_EQUAL_TO = 6
    CONTAINS = 7
    DOES_NOT_CONTAINS = 8
    STARTS_WITH = 9
    ENDS_WITH = 10
    IN_BETWEEN = 11
    NOT_IN_BETWEEN = 12
    
End Enum

Public Function Create(TypeOfTester As Operator, ParamArray Args() As Variant) As Predicate
    
    Dim Tester As Predicate
    Select Case TypeOfTester
        Case Operator.Equal
            Set Tester = EqualPredicate.Create(Args(0))
            
        Case Operator.DOES_NOT_EQUAL
            Set Tester = NotEqualPredicate.Create(Args(0))
            
        Case Operator.GREATER_THAN
            Set Tester = GreaterThanPredicate.Create(Args(0))
            
        Case Operator.GREATER_THAN_OR_EQUAL_TO
            Set Tester = GreaterThanOrEqualPredicate.Create(Args(0))
            
        Case Operator.LESS_THAN
            Set Tester = LessThanPredicate.Create(Args(0))
            
        Case Operator.LESS_THAN_OR_EQUAL_TO
            Set Tester = LessThanOrEqualPredicate.Create(Args(0))
            
        Case Operator.CONTAINS
            Set Tester = ContainsPredicate.Create(Args(0))
                
        Case Operator.DOES_NOT_CONTAINS
            Set Tester = NotContainsPredicate.Create(Args(0))
            
        Case Operator.STARTS_WITH
           Set Tester = StartWithPredicate.Create(CStr(Args(0)))
           
        Case Operator.ENDS_WITH
            Set Tester = EndWithPredicate.Create(CStr(Args(0)))
            
        Case Operator.IN_BETWEEN
            Set Tester = BetweenDatePredicate.Create(CDate(Args(0)), CDate(Args(1)))
            
        Case Operator.NOT_IN_BETWEEN
            Set Tester = NotBetweenDatePredicate.Create(CDate(Args(0)), CDate(Args(1)))
        Case Else
            Set Tester = DefaultPredicate.Create(Args(0))
    End Select
    Set Create = Tester
    
End Function
