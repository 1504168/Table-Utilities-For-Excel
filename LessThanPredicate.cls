VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LessThanPredicate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("Predicate.Implementation")

Option Explicit

Implements Predicate
Private Type TLessThanPredicate
    Data As Variant
End Type

Private This As TLessThanPredicate

Public Property Get Data() As Variant
    If IsObject(This.Data) Then
        Set Data = This.Data
    Else
        Data = This.Data
    End If
End Property

Public Property Let Data(ByVal RHS As Variant)
    This.Data = RHS
End Property

Public Property Set Data(ByVal RHS As Variant)
    Set This.Data = RHS
End Property

Public Function Create(GivenData As Variant) As LessThanPredicate
    Dim CurrentLessThanPredicate As LessThanPredicate
    Set CurrentLessThanPredicate = New LessThanPredicate
    With CurrentLessThanPredicate
        If IsObject(GivenData) Then
            Set .Data = GivenData
        Else
            .Data = GivenData
        End If
    End With
    Set Create = CurrentLessThanPredicate
End Function

Private Function Predicate_IsValid(RunTestOn As Variant) As Boolean
    Predicate_IsValid = (RunTestOn < This.Data)
End Function
