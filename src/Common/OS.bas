Attribute VB_Name = "OS"

Public Function isMac()
    isMac = isMac("macintosh")
End Function

Public Function isWindows()
    isWindows = isOS("windows")
End Function

Private Function isOS(name As String)
    isOS = WCString.IsStartsWith(os(), name)
End Function

Private Function os()
    os = LCase(Application.OperatingSystem)
End Function
