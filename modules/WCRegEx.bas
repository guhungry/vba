Attribute VB_Name = "WCRegEx"
Option Explicit

' Basic Regular Expression Matcher
' Support for \d \w + *
' Example WCRegEx.Match("* Last Update 12 February 2019.", "\d+ \w+ \d\d\d\d") will match 12 February 2019
Public Function Match(text As String, pattern As String)
    Dim result As Variant: result = DoMatch(text, pattern)

    If result(0) Then
        Match = Mid(text, result(1), result(2))
    Else
        Match = ""
    End If
End Function

Public Function IsMatch(text As String, pattern As String)
    Dim result As Variant: result = DoMatch(text, pattern)

    IsMatch = result(0)
End Function

Private Function DoMatch(text As String, pattern As String)
    Dim finalPattern As Variant: finalPattern = AnalysePattern(pattern)

    DoMatch = DoSubMatch(text, (finalPattern(0)), (finalPattern(1)), (finalPattern(2)))
End Function

Private Function AnalysePattern(pattern As String)
    Dim result(0 To 3) As Variant
    result(0) = pattern
    result(1) = WCString.IsStartsWith(pattern, "^")
    result(2) = WCString.IsEndsWith(pattern, "$")

    If result(1) Then
        result(0) = Mid(result(0), 2)
    End If

    If result(2) Then
        result(0) = Left(result(0), Len(result(0)) - 1)
    End If

    AnalysePattern = result
End Function

Private Function DoSubMatch(text As String, pattern As String, checkStart As Boolean, checkEnd As Boolean)
    Dim startText As Integer: startText = 1
    Dim indexText As Integer: indexText = startText
    Dim currentText As String: currentText = ""
    Dim lengthText As Integer: lengthText = Len(text)

    Dim indexPattern As Integer: indexPattern = 1
    Dim currentPattern As String: currentPattern = ""
    Dim nextPattern As String: nextPattern = ""
    Dim lastPattern As String: lastPattern = ""
    Dim lengthPattern As Integer: lengthPattern = Len(pattern)

    Do While indexPattern <= lengthPattern And indexText <= lengthText
        currentText = Mid(text, indexText, 1)
        currentPattern = FindNextPattern(pattern, indexPattern)
        nextPattern = FindNextPattern(pattern, indexPattern + Len(currentPattern))
        Dim activePattern As String: activePattern = FindActivePattern(currentPattern, lastPattern, nextPattern, indexPattern)

        Dim IsMatch As Boolean: IsMatch = InCharSet(activePattern, currentText)

        ' Text Index
        If IsMatch Then
            indexText = indexText + 1
        End If

        ' Pattern Index
        If IsMatch Then
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
            If checkStart Then
                startText = lengthText
            End If

            startText = startText + 1
            indexText = startText

            indexPattern = 1
            currentPattern = ""
            lastPattern = ""
        End If
    Loop

    Dim result(0 To 3) As Variant
    result(0) = indexPattern > lengthPattern

    If result(0) And checkEnd And lengthText <> indexText - 1 Then
        result(0) = False
    End If

    If result(0) Then
        result(1) = startText
        result(2) = indexText - startText
    Else
        result(1) = 0
        result(2) = 0
    End If

    DoSubMatch = result
End Function

Private Function FindActivePattern(current As String, last As String, nextPattern As String, index As Integer)
    If WCString.IsSubString("+*", current) Then
        FindActivePattern = last
    ElseIf current = "^" And index = 1 Then
        FindActivePattern = nextPattern
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
    ElseIf Left(pattern, 1) = "\" And Len(pattern) = 2 Then
        CharSet = Mid(pattern, 2)
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

    ' +
    TestMatch "0123", "1+", "1"
    TestMatch "0113", "1+", "11"
    TestMatch "0113", "1+", "11"

    ' *
    TestMatch "0123", "12*3", "123"
    TestMatch "012223", "12*3", "12223"
    TestMatch "013", "12*3", "13"

    ' ?
    TestMatch "0121ab3", "1a?b", "1ab"
    TestMatch "012b3", "2a?b", "2b"

    ' Escape with \
    TestMatch "012a*b3", "a\*b", "a*b"
    TestMatch "012a+b3", "a\+b", "a+b"
    TestMatch "012a\b3", "a\\b", "a\b"
    TestMatch "012a?b3", "a\?b", "a?b"
    TestMatch "012a^b3", "a\^b", "a^b"
    TestMatch "^012a^b3", "\^0", "^0"

    ' \d
    TestMatch " 1", "\d", "1"
    TestMatch " a", "\d", ""

    ' \w
    TestMatch " a", "\w", "a"
    TestMatch " 1", "\w", ""

    ' \s
    TestMatch " ", "\s", " "

    ' []
    TestMatch "a", "[bc]", ""
    TestMatch "b", "[bc]", "b"
    TestMatch "1", "[\d]", "1"
    TestMatch "a", "[\d]", ""

    ' [^]
    TestMatch "a", "[^bc]", "a"
    TestMatch "b", "[^bc]", ""
    TestMatch "a", "[^\d]", "a"
    TestMatch "1", "[^\d]", ""

    ' ^
    TestMatch "beetest", "^bee", "bee"
    TestMatch "beetest", "^ee", ""

    ' $
    TestMatch "beetest", "est$", "est"
    TestMatch "beetest", "ee$", ""
End Sub

Private Sub TestMatch(text As String, pattern As String, expected As String)
    Dim result As String: result = Match(text, pattern)
    If result <> expected Then
        MsgBox "Failed Test: [text]=" & text & ",[pattern]=" & pattern & ",[result]=" & result & ",[expected]=" & expected
    End If
End Sub
