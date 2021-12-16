VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StartWithPredicate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("Predicate.Implementation")

Option Explicit
Implements Predicate

Private Type TStartWithPredicate
    FindText As String
End Type

Private This As TStartWithPredicate

Public Property Get FindText() As String
    FindText = This.FindText
End Property

Public Property Let FindText(ByVal RHS As String)
    This.FindText = UCase(RHS)
End Property

Public Function Create(GivenFindText As String) As Predicate
    Dim CurrentStartWithPredicate As StartWithPredicate
    Set CurrentStartWithPredicate = New StartWithPredicate
    With CurrentStartWithPredicate
        .FindText = GivenFindText
    End With
    Set Create = CurrentStartWithPredicate
End Function

Private Function Predicate_IsValid(RunTestOn As Variant) As Boolean
    Predicate_IsValid = (Left(UCase(RunTestOn), Len(This.FindText)) = This.FindText)
End Function

