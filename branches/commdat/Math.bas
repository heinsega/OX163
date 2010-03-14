Attribute VB_Name = "Math"
Option Explicit

Public Function Max(ParamArray numbers() As Variant) As Variant
    If UBound(numbers) >= 0 Then
        Dim index As Long, maxNumber As Variant
        maxNumber = numbers(0)
        For index = 0 To UBound(numbers)
            If numbers(index) > maxNumber Then maxNumber = numbers(index)
        Next index
        Max = maxNumber
    Else
        Max = Empty
    End If
End Function

Public Function Min(ParamArray numbers() As Variant) As Variant
    If UBound(numbers) >= 0 Then
        Dim index As Long, minNumber As Variant
        minNumber = numbers(0)
        For index = 0 To UBound(numbers)
            If numbers(index) < minNumber Then minNumber = numbers(index)
        Next index
        Min = minNumber
    Else
        Min = Empty
    End If
End Function

