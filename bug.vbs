Function that doesn't properly handle empty strings can cause unexpected behavior or errors. For example, consider a function designed to extract the first three characters from a string. If it receives an empty string, it might throw an error or return an unexpected result.  Here's an example illustrating the error:

```vbscript
Function GetFirstThreeChars(strInput)
  If Len(strInput) >= 3 Then
    GetFirstThreeChars = Left(strInput, 3)
  Else
    GetFirstThreeChars = ""
  End If
End Function

Dim myString
myString = ""
MsgBox GetFirstThreeChars(myString) 'Returns empty string, as expected

myString = "abcde"
MsgBox GetFirstThreeChars(myString) 'Returns "abc", as expected

myString = Null
MsgBox GetFirstThreeChars(myString) 'Throws error: Type mismatch
```
The error occurs because the `Len` function cannot handle `Null` values.  A robust solution would check for `Null` values before using `Len`.