VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Filler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit
Private Enum TypeOfFill
    Down = 1
    Up = 2
End Enum

Public Function FillDown(GivenArray As Variant) As Variant
    FillDown = FillArray(GivenArray, Down)
End Function

Public Function FillUp(GivenArray As Variant) As Variant
    FillUp = FillArray(GivenArray, Up)
End Function

Private Function FillArray(ByVal GivenArray As Variant, FillType As TypeOfFill) As Variant
    
    Dim LoopStartIndex As Long
    Dim LoopEndIndex As Long
    Dim LoopStep As Long
    
    If FillType = TypeOfFill.Down Then
        LoopStartIndex = LBound(GivenArray, 1)
        LoopEndIndex = UBound(GivenArray, 1)
        LoopStep = 1
    Else
        LoopStartIndex = UBound(GivenArray, 1)
        LoopEndIndex = LBound(GivenArray, 1)
        LoopStep = -1
    End If
    
    Dim LastNonEmptyValue As Variant
    Dim CurrentColumnIndex As Long
    For CurrentColumnIndex = LBound(GivenArray, 2) To UBound(GivenArray, 2)
        Dim CurrentRowIndex As Long
        LastNonEmptyValue = vbNullString
        For CurrentRowIndex = LoopStartIndex To LoopEndIndex Step LoopStep
            If Trim(GivenArray(CurrentRowIndex, CurrentColumnIndex)) <> vbNullString Then
                LastNonEmptyValue = GivenArray(CurrentRowIndex, CurrentColumnIndex)
            End If
            GivenArray(CurrentRowIndex, CurrentColumnIndex) = LastNonEmptyValue
        Next CurrentRowIndex
    Next CurrentColumnIndex
    FillArray = GivenArray
    
End Function

