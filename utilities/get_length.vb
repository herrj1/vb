Public Function GetLength(arr As Variant) As Integer
   If IsEmpty(arr) Then
      GetLength = 0
   Else
      GetLength = UBound(arr) - LBound(arr)
   End If
End Function