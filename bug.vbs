Function GetObject(progID)
  On Error Resume Next
  Set obj = GetObject(progID)
  If Err.Number <> 0 Then
    Set obj = CreateObject(progID)
  End If
  Set GetObject = obj
End Function

'This function may throw an error if the object creation fails and error handling is not robust.
Set myObj = GetObject("MyObject.MyClass")
If IsObject(myObj) Then
  'Use myObj
Else
  'Handle the error
End If