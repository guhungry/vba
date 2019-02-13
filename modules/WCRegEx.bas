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
        Exit Function
    End If
    
    Dim first As String: first = Mid(pattern, index, 1)
    If first = "\" Then
        FindNextPattern = Mid(pattern, index, 2)
    ElseIf first = "[" Then
        Dim result As String: result = first
        Dim current As String: current = first
        
        ' Find charset in []
        Do While Len(pattern) >= index + Len(result) And current <> "]" And current <> ""
            current = FindNextPattern(pattern, index + Len(result))
            result = result & current
        Loop
        
        FindNextPattern = result
    Else
        FindNextPattern = first
    End If
End Function

Private Function InCharSet(pattern As String, char As String)
    Dim charList As String: charList = CharSet(pattern)
    
    If WCString.IsStartsWith(pattern, "[") And WCString.IsEndsWith(pattern, "]") Then
        If WCString.IsStartsWith(charList, "^") Then
            InCharSet = Not WCString.IsSubString(Mid(charList, 2), char)
        Else
            InCharSet = WCString.IsSubString(charList, char)
        End If
    Else
        InCharSet = WCString.IsSubString(charList, char)
    End If
End Function

Private Function CharSet(pattern As String)
    If pattern = "\d" Then
        CharSet = "0123456789"
    ElseIf pattern = "\w" Then
        CharSet = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    ElseIf pattern = "\s" Then
        CharSet = " " & Chr(9)
    ElseIf pattern = "\*" Then
        CharSet = "*"
    ElseIf pattern = "\+" Then
        CharSet = "+"
    ElseIf pattern = "\\" Then
        CharSet = "\"
    ElseIf pattern = "\?" Then
        CharSet = "?"
    ElseIf pattern = "\[" Then
        CharSet = "["
    ElseIf pattern = "\]" Then
        CharSet = "]"
    ElseIf WCString.IsStartsWith(pattern, "[") Then
        Dim tempPattern  As String: tempPattern = Mid(pattern, 2, Len(pattern) - 2)
        Dim current  As String: current = ""
        Dim index As Integer: index = 1
        Dim length As Integer: length = Len(tempPattern)
        Dim result As String: result = ""
        
        ' Expand charset in []
        Do While length >= index
            current = FindNextPattern(tempPattern, index)
            index = index + Len(current)
            result = result & CharSet(current)
        Loop
        CharSet = result
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

