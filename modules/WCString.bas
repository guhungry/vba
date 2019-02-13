Attribute VB_Name = "WCString"
Option Explicit

' String Helper Functions

Public Function IsSubString(text As String, search As String)
    IsSubString = InStr(text, search) <> 0
End Function
