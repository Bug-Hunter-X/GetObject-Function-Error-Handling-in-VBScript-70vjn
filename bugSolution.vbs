Function GetObjectSafe(progID)
  Dim obj, errNum
  On Error Resume Next
  Set obj = GetObject(progID)
  errNum = Err.Number
  On Error GoTo 0
  If errNum <> 0 Then
    ' Check for specific errors if needed
    If Err.Number = 429 Then 'ActiveX component can't create object
       MsgBox "Error creating object: " & Err.Description, vbCritical
    Else
       MsgBox "Error getting object: " & Err.Description, vbCritical
    End If
    Set obj = Nothing
  End If
  Set GetObjectSafe = obj
End Function

'Example usage
Set myObj = GetObjectSafe("MyObject.MyClass")
If IsObject(myObj) Then
  'Use myObj
Else
  'Handle the case where the object was not created
End If