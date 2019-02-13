Attribute VB_Name = "WCRegEx"
Option Explicit

' Basic Regular Expression Matcher
' Support for \d \w + *
' Example WCRegEx.Match("* Last Update 12 February 2019.", "\d+ \w+ \d\d\d\d") will match 12 February 2019

Public Function Match(text As String, pattern As String)
    Dim startText As Integer: startText = 1
    Dim indexText As Integer: indexText = startText
    Dim lengthText As Integer: lengthText = Len(text)
    Dim result As String: result = ""
    
    Dim indexPattern As Integer: indexPattern = 1
    Dim currentPattern As String: currentPattern = ""
    Dim lastPattern As String: lastPattern = ""
    Dim lengthPattern As Integer: lengthPattern = Len(pattern)
    
    Do While indexPattern <= lengthPattern And indexText <= lengthText
        currentPattern = NextPattern(pattern, indexPattern)
        Dim activePattern As String: activePattern = FindActivePattern(currentPattern, lastPattern)
        
        Dim isMatch As Boolean: isMatch = InCharSet(activePattern, Mid(text, indexText, 1))

        ' Text Index
        If isMatch Then
            indexText = indexText + 1
        End If

        ' Pattern Index
        If isMatch Then
            If Not IsSubString("+*", currentPattern) Then
                lastPattern = currentPattern
                indexPattern = indexPattern + Len(currentPattern)
            ElseIf indexText > lengthText Then
                indexPattern = indexPattern + Len(currentPattern)
            End If
        ElseIf IsSubString("+*", currentPattern) Then
            lastPattern = ""
            indexPattern = indexPattern + Len(currentPattern)
        ElseIf NextPattern(pattern, indexPattern + Len(currentPattern)) = "*" Then
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

Private Function NextPattern(pattern As String, index As Integer)
    If (Len(pattern) < index) Then
        NextPattern = ""
    ElseIf (Mid(pattern, index, 1) = "\") Then
        NextPattern = Mid(pattern, index, 2)
    Else
        NextPattern = Mid(pattern, index, 1)
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
    Else
        CharSet = pattern
    End If
End Function




