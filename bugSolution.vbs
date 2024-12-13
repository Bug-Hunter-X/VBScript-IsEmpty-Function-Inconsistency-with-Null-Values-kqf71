Function MyFunction(param1)
  If param1 = vbNullString Or IsNull(param1) Then
    ' Handle the null or empty case appropriately
    param1 = ""  ' Or some other default value
  End If
  ' ... rest of the function ...
End Function