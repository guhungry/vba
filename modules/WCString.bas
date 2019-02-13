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
