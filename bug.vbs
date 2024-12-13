Function MyFunction(param1)
  If IsEmpty(param1) Then
    Err.Raise 13, , "Type mismatch: Expected String, got Null"
  End If
  ' ... rest of the function ...
End Function