VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CHttpHeaders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==============================================================
' CHttpHeaders
'
' Ordered, case-insensitive collection of HTTP header name/value
' pairs. Backed by a plain Collection (available since VB4) so
' there is no dependency on the Scripting.Dictionary object,
' which is not guaranteed to be installed on a base NT4 machine.
'
' Usage:
'     Dim headers As New CHttpHeaders
'     headers.SetHeader "Content-Type", "text/plain"
'     If headers.Exists("content-type") Then
'         Debug.Print headers.GetHeader("Content-Type")
'     End If
'==============================================================

Private m_Items As Collection

Private Sub Class_Initialize()
    Set m_Items = New Collection
End Sub

Public Property Get Count() As Long
    Count = m_Items.Count
End Property

''' Sets Name/Value, replacing any existing header with the same name
''' (comparison is case-insensitive, per the HTTP specification).
Public Sub SetHeader(ByVal Name As String, ByVal Value As String)
    Dim pair As New CHttpHeaderPair

    If Exists(Name) Then m_Items.Remove UCase$(Name)

    pair.Name = Name
    pair.Value = Value
    m_Items.Add pair, UCase$(Name)
End Sub

''' Returns True when a header with the given name (case-insensitive)
''' has been set.
Public Function Exists(ByVal Name As String) As Boolean
    Dim pair As CHttpHeaderPair

    On Error Resume Next
    Set pair = m_Items(UCase$(Name))
    On Error GoTo 0

    Exists = Not (pair Is Nothing)
End Function

''' Returns the value for Name, or "" if the header has not been set.
Public Function GetHeader(ByVal Name As String) As String
    Dim pair As CHttpHeaderPair

    On Error Resume Next
    Set pair = m_Items(UCase$(Name))
    On Error GoTo 0

    If pair Is Nothing Then
        GetHeader = ""
    Else
        GetHeader = pair.Value
    End If
End Function

Public Sub Remove(ByVal Name As String)
    If Exists(Name) Then m_Items.Remove UCase$(Name)
End Sub

Public Sub Clear()
    Set m_Items = New Collection
End Sub

''' Returns the Nth header (1-based) as a CHttpHeaderPair, for enumeration.
Public Function Item(ByVal Index As Long) As CHttpHeaderPair
    Set Item = m_Items(Index)
End Function

''' Renders all headers as "Name: Value" lines terminated with CRLF,
''' ready to be appended to an HTTP request/response message.
Public Function ToRawHeaders() As String
    Dim i As Long
    Dim pair As CHttpHeaderPair
    Dim result As String

    result = ""
    For i = 1 To m_Items.Count
        Set pair = m_Items(i)
        result = result & pair.Name & ": " & pair.Value & HTTP_CRLF
    Next i

    ToRawHeaders = result
End Function
