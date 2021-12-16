VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BetweenDatePredicate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("Predicate.Implementation")
Option Explicit

Implements Predicate

Private Type TBetweenDatePredicate
    StartDate As Date
    EndDate As Date
End Type

Private This As TBetweenDatePredicate

Public Property Get StartDate() As Date
    StartDate = This.StartDate
End Property

Public Property Let StartDate(ByVal RHS As Date)
    This.StartDate = RHS
End Property

Public Property Get EndDate() As Date
    EndDate = This.EndDate
End Property

Public Property Let EndDate(ByVal RHS As Date)
    This.EndDate = RHS
End Property

Public Function Create(GivenStartDate As Date, GivenEndDate As Date) As BetweenDatePredicate
    Dim CurrentBetweenDatePredicate As BetweenDatePredicate
    Set CurrentBetweenDatePredicate = New BetweenDatePredicate
    With CurrentBetweenDatePredicate
        .StartDate = GivenStartDate
        .EndDate = GivenEndDate
    End With
    Set Create = CurrentBetweenDatePredicate
End Function

Private Function Predicate_IsValid(RunTestOn As Variant) As Boolean
    Predicate_IsValid = (RunTestOn >= StartDate And RunTestOn <= EndDate)
End Function
