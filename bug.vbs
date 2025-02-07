Function MyFunction(param)
  If IsEmpty(param) Then
    ' Handle empty parameter
    MsgBox "Parameter is empty", vbExclamation
    Exit Function
  End If

  ' ... rest of the function
End Function