Attribute VB_Name = "WCString"
Option Explicit

' String Helper Functions

Public Function IsSubString(text As String, search As String)
    IsSubString = InStr(text, search) <> 0
End Function

Public Function IsStartsWith(text As String, search As String)
    IsStartsWith = Left(text, Len(search)) = search
End Function

Public Function IsEndsWith(text As String, search As String)
    IsEndsWith = Right(text, Len(search)) = search
End Function

Private Sub UnitTest()
    ' IsSubString
    Debug.Assert (IsSubString("beetest", "eet") = True)
    Debug.Assert (IsSubString("beetest", "tata") = False)

    ' IsStartsWith
    Debug.Assert (IsSubString("beetest", "bee") = True)
    Debug.Assert (IsSubString("beetest", "tata") = False)

    ' IsEndsWith
    Debug.Assert (IsEndsWith("beetest", "test") = True)
    Debug.Assert (IsEndsWith("beetest", "tata") = False)
End Sub
