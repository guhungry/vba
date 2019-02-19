Attribute VB_Name = "WCDate"
Option Explicit

' Extract date from text. Support date in 'dd MMM yyyy' and dd/mm/yyyy
Public Function ExtractDate(text As String)
    ExtractDate = DateValue(FindDate(text))
End Function

Private Function FindDate(text As String)
    Dim value As String: value = WCRegEx.Match(text, "\d\d? \w+ \d\d\d\d")
    
    If value = "" Then
        value = WCRegEx.Match(text, "\d\d/\d\d/\d\d\d\d")
        value = ToShortMonth(value)
    End If
    FindDate = value
End Function

Private Function ToShortMonth(text As String)
    text = Replace(text, "/01/", " Jan ")
    text = Replace(text, "/02/", " Feb ")
    text = Replace(text, "/03/", " Mar ")
    text = Replace(text, "/04/", " Apr ")
    text = Replace(text, "/05/", " May ")
    text = Replace(text, "/06/", " Jun ")
    text = Replace(text, "/07/", " Jul ")
    text = Replace(text, "/08/", " Aug ")
    text = Replace(text, "/09/", " Sep ")
    text = Replace(text, "/10/", " Oct ")
    text = Replace(text, "/11/", " Nov ")
    ToShortMonth = Replace(text, "/12/", " Dec ")
End Function

Private Sub UnitTest()
    ' dd MMM yyyy
    TestExtractDate "14 May 1984", "14/05/84"
    TestExtractDate "prefix 14 May 1984 suffix", "14/05/84"
    
    ' dd/MM/yyyy
    TestExtractDate "14/05/1984", "14/05/84"
    TestExtractDate " prefix 14/05/1994  noise", "14/05/94"
End Sub

Private Sub TestExtractDate(text As String, expected As String)
    If ExtractDate(text) <> DateValue(expected) Then
        MsgBox "[text]=" & text & ",[value]=" & ExtractDate(text) & ",[expected]=" & DateValue(expected)
    End If
End Sub
