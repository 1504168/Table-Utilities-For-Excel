VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NotContainsPredicate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("Predicate.Implementation")

Option Explicit

Private Type TNotContainsPredicate
    FindText As Variant
End Type

Private This As TNotContainsPredicate

Public Property Get FindText() As Variant
    If IsObject(This.FindText) Then
        Set FindText = This.FindText
    Else
        FindText = This.FindText
    End If
End Property

Public Property Let FindText(ByVal RHS As Variant)
    This.FindText = RHS
End Property

Public Property Set FindText(ByVal RHS As Variant)
    Set This.FindText = RHS
End Property

Public Function Create(GivenFindText As Variant) As NotContainsPredicate
    Dim CurrentNotContainsPredicate As NotContainsPredicate
    Set CurrentNotContainsPredicate = New NotContainsPredicate
    With CurrentNotContainsPredicate
        If IsObject(GivenFindText) Then
            Set .FindText = GivenFindText
        Else
            .FindText = GivenFindText
        End If
    End With
    Set Create = CurrentNotContainsPredicate
End Function

Private Function Predicate_IsValid(RunTestOn As Variant) As Boolean
    Predicate_IsValid = Not (InStr(1, RunTestOn, This.FindText, vbTextCompare) <> 0)
End Function
