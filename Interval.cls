VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Interval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("TableUtility.Helper")

Option Explicit

Private Type TInterval
    Minimum As Long
    Maximum As Long
End Type

Private This As TInterval

Public Property Get Minimum() As Long
    Minimum = This.Minimum
End Property

Public Property Let Minimum(ByVal RHS As Long)
    This.Minimum = RHS
End Property

Public Property Get Maximum() As Long
    Maximum = This.Maximum
End Property

Public Property Let Maximum(ByVal RHS As Long)
    This.Maximum = RHS
End Property

Public Function Create(GivenMinimum As Long, GivenMaximum As Long) As Interval
    Dim CurrentInterval As Interval
    Set CurrentInterval = New Interval
    With CurrentInterval
        .Minimum = GivenMinimum
        .Maximum = GivenMaximum
    End With
    Set Create = CurrentInterval
End Function

