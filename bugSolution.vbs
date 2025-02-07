Function MyFunction(param)
  If IsEmpty(param) Then
    ' Handle empty parameter gracefully
    param = ""
  End If

  ' Now the function can proceed without errors, even if the parameter is empty
  MsgBox "Parameter value is: " & param
End Function