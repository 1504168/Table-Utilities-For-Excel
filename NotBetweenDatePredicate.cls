VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NotBetweenDatePredicate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("Predicate.Implementation")

Option Explicit

Implements Predicate
Private Type TNotBetweenDatePredicate
    StartDate As Date
    EndDate As Date
End Type

Private This As TNotBetweenDatePredicate

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

Public Function Create(GivenStartDate As Date, GivenEndDate As Date) As NotBetweenDatePredicate
    Dim CurrentNotBetweenDatePredicate As NotBetweenDatePredicate
    Set CurrentNotBetweenDatePredicate = New NotBetweenDatePredicate
    With CurrentNotBetweenDatePredicate
        .StartDate = GivenStartDate
        .EndDate = GivenEndDate
    End With
    Set Create = CurrentNotBetweenDatePredicate
End Function

Private Function Predicate_IsValid(RunTestOn As Variant) As Boolean
    Predicate_IsValid = Not (RunTestOn >= StartDate And RunTestOn <= EndDate)
End Function

