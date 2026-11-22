Attribute VB_Name = "modStringUtils"
Option Explicit

'==============================================================
' modStringUtils
'
' VB5 does not provide the Split, Join or Replace functions -
' they were only added in VB6 - so this module implements the
' small subset this library needs on top of functions that are
' guaranteed to exist in VB5 (InStr, Mid$, Left$, Len).
'==============================================================

''' Splits Text on every occurrence of Delimiter, filling Parts with the
''' resulting zero-based array of substrings (mirrors VB6's Split).
'''
''' This is a Sub with a ByRef array parameter, rather than a Function
''' returning an array, because VB5 doesn't support assigning a whole
''' array from a function's return value ("Can't assign to array") -
''' that capability was only added in VB6 alongside Split/Join/Replace
''' themselves. Filling the caller's array by reference works in VB5.
Public Sub SplitString(ByVal Text As String, ByVal Delimiter As String, ByRef Parts() As String)
    Dim count As Long
    Dim startPos As Long
    Dim foundPos As Long

    ReDim Parts(0 To 0)
    count = 0
    startPos = 1

    If Len(Delimiter) = 0 Then
        Parts(0) = Text
        Exit Sub
    End If

    Do
        foundPos = InStr(startPos, Text, Delimiter)
        If foundPos = 0 Then
            ReDim Preserve Parts(0 To count)
            Parts(count) = Mid$(Text, startPos)
            Exit Do
        Else
            ReDim Preserve Parts(0 To count)
            Parts(count) = Mid$(Text, startPos, foundPos - startPos)
            count = count + 1
            startPos = foundPos + Len(Delimiter)
        End If
    Loop
End Sub

''' Joins an array of strings using Delimiter (mirrors VB6's Join function).
Public Function JoinStrings(ByRef Parts() As String, ByVal Delimiter As String) As String
    Dim i As Long
    Dim result As String

    result = ""
    For i = LBound(Parts) To UBound(Parts)
        If i > LBound(Parts) Then result = result & Delimiter
        result = result & Parts(i)
    Next i

    JoinStrings = result
End Function

''' Replaces every occurrence of Find within Text with Replacement
''' (mirrors VB6's Replace function for the common case).
Public Function ReplaceString(ByVal Text As String, ByVal Find As String, ByVal Replacement As String) As String
    Dim result As String
    Dim startPos As Long
    Dim foundPos As Long

    If Len(Find) = 0 Then
        ReplaceString = Text
        Exit Function
    End If

    result = ""
    startPos = 1

    Do
        foundPos = InStr(startPos, Text, Find)
        If foundPos = 0 Then
            result = result & Mid$(Text, startPos)
            Exit Do
        Else
            result = result & Mid$(Text, startPos, foundPos - startPos) & Replacement
            startPos = foundPos + Len(Find)
        End If
    Loop

    ReplaceString = result
End Function
