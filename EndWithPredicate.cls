VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EndWithPredicate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("Predicate.Implementation")

Option Explicit

Implements Predicate
Private Type TEndWithPredicate
    FindText As String
End Type

Private This As TEndWithPredicate

Public Property Get FindText() As String
    FindText = This.FindText
End Property

Public Property Let FindText(ByVal RHS As String)
    This.FindText = RHS
End Property

Public Function Create(GivenFindText As String) As EndWithPredicate
    Dim CurrentEndWithPredicate As EndWithPredicate
    Set CurrentEndWithPredicate = New EndWithPredicate
    With CurrentEndWithPredicate
        .FindText = GivenFindText
    End With
    Set Create = CurrentEndWithPredicate
End Function

Private Function Predicate_IsValid(RunTestOn As Variant) As Boolean
    Predicate_IsValid = (Right(UCase(RunTestOn), Len(This.FindText)) = This.FindText)
End Function
