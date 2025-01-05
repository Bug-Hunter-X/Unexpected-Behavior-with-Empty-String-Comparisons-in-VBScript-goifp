Function MyFunction(param1)
  ' Explicitly check for Empty and Null values
  If IsEmpty(param1) Or IsNull(param1) Then
    Exit Function
  ElseIf Trim(param1) = "" Then 'Added Trim for whitespace check
    Exit Function
  End If
  'Rest of the function code
End Function