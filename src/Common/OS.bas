Attribute VB_Name = "OS"

Public Function isMac()
    isMac = isOS("macintosh")
End Function

Public Function isWindows()
    isWindows = isOS("windows")
End Function

Private Function isOS(Name As String)
    isOS = WCString.IsStartsWith(os(), Name)
End Function

Private Function os()
    os = LCase(Application.OperatingSystem)
End Function
