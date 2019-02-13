Attribute VB_Name = "WCRegEx"
Option Explicit

' Basic Regular Expression Matcher
' Support for \d \w + *
' Example WCRegEx.Match("* Last Update 12 February 2019.", "\d+ \w+ \d\d\d\d") will match 12 February 2019
Public Function Match(text As String, pattern As String)
    Dim startText As Integer: startText = 1
    Dim indexText As Integer: indexText = startText
    Dim currentText As String: currentText = ""
    Dim lengthText As Integer: lengthText = Len(text)
    Dim result As String: result = ""
    
    Dim indexPattern As Integer: indexPattern = 1
    Dim currentPattern As String: currentPattern = ""
    Dim nextPattern As String: nextPattern = ""
    Dim lastPattern As String: lastPattern = ""
    Dim lengthPattern As Integer: lengthPattern = Len(pattern)
    
    Do While indexPattern <= lengthPattern And indexText <= lengthText
        currentText = Mid(text, indexText, 1)
        currentPattern = FindNextPattern(pattern, indexPattern)
        nextPattern = FindNextPattern(pattern, indexPattern + Len(currentPattern))
        Dim activePattern As String: activePattern = FindActivePattern(currentPattern, lastPattern)
        
        Dim isMatch As Boolean: isMatch = InCharSet(activePattern, currentText)

        ' Text Index
        If isMatch Then
            indexText = indexText + 1
        End If

        ' Pattern Index
        If isMatch Then
            If nextPattern = "?" Then
                lastPattern = ""
                indexPattern = indexPattern + Len(currentPattern) + 1
            ElseIf Not IsSubString("+*", currentPattern) Then
                lastPattern = currentPattern
                indexPattern = indexPattern + Len(currentPattern)
            ElseIf indexText > lengthText Then
                indexPattern = indexPattern + Len(currentPattern)
            End If
        ElseIf IsSubString("+*", currentPattern) Then
            lastPattern = ""
            indexPattern = indexPattern + Len(currentPattern)
        ElseIf nextPattern <> "" And IsSubString("*?", nextPattern) Then
            lastPattern = ""
            indexPattern = indexPattern + Len(currentPattern) + 1
        Else
            startText = startText + 1
            indexText = startText
            
            indexPattern = 1
            currentPattern = ""
            lastPattern = ""
        End If
    Loop
    
    If indexPattern > lengthPattern Then
        Match = Mid(text, startText, indexText - startText)
    Else
        Match = ""
    End If
End Function

Private Function FindActivePattern(current As String, last As String)
    If (current = "+") Then
        FindActivePattern = last
    Else
        FindActivePattern = current
    End If
End Function

Private Function FindNextPattern(pattern As String, index As Integer)
    If (Len(pattern) < index) Then
        FindNextPattern = ""
    ElseIf (Mid(pattern, index, 1) = "\") Then
        FindNextPattern = Mid(pattern, index, 2)
    Else
        FindNextPattern = Mid(pattern, index, 1)
    End If
End Function

Private Function InCharSet(pattern As String, char As String)
    InCharSet = WCString.IsSubString(CharSet(pattern), char)
End Function

Private Function CharSet(pattern As String)
    If pattern = "\d" Then
        CharSet = "0123456789"
    ElseIf pattern = "\w" Then
        CharSet = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    ElseIf pattern = "\*" Then
        CharSet = "*"
    ElseIf pattern = "\+" Then
        CharSet = "+"
    ElseIf pattern = "\\" Then
        CharSet = "\"
    ElseIf pattern = "\?" Then
        CharSet = "?"
    Else
        CharSet = pattern
    End If
End Function

Private Sub UnitTest()
    TestMatch "0123", "12", "12"
    TestMatch "0123", "1+", "1"
    TestMatch "0113", "1+", "11"
    TestMatch "0113", "1+", "11"
    TestMatch "0123", "12*", "12"
    TestMatch "013", "12*", "1"
    TestMatch "012a*b3", "a\*b", "a*b"
    TestMatch "012a+b3", "a\+b", "a+b"
    TestMatch "012a\b3", "a\\b", "a\b"
    TestMatch "012a?b3", "a\?b", "a?b"
    TestMatch "0121ab3", "1a?b", "1ab"
    TestMatch "012b3", "2a?b", "2b"
End Sub

Private Sub TestMatch(text As String, pattern As String, expected As String)
    Dim result As String: result = Match(text, pattern)
    If result <> expected Then
        MsgBox "Failed Test: [text]=" & text & ",[pattern]=" & pattern & ",[result]=" & result & ",[expected]=" & expected
    End If
End Sub

