VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BeginWithPredicate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit
Implements Predicate

Private Type TBeginWithPredicate
    FindText As String
End Type

Private This As TBeginWithPredicate

Public Property Get FindText() As String
    FindText = This.FindText
End Property

Public Property Let FindText(ByVal RHS As String)
    This.FindText = UCase(RHS)
End Property

Public Function Create(GivenFindText As String) As Predicate
    Dim CurrentBeginWithPredicate As BeginWithPredicate
    Set CurrentBeginWithPredicate = New BeginWithPredicate
    With CurrentBeginWithPredicate
        .FindText = GivenFindText
    End With
    Set Create = CurrentBeginWithPredicate
End Function

Private Function Predicate_IsValid(ParamArray args() As Variant) As Boolean
    Debug.Print Left(UCase(args(0)), Len(This.FindText))
    Predicate_IsValid = (Left(UCase(args(0)), Len(This.FindText)) = This.FindText)
End Function

