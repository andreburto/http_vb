VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CHttpRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==============================================================
' CHttpRequest
'
' Describes an outgoing HTTP request. Build one, optionally set
' Headers/Body, then hand it to CHttpClient.Execute.
'
' Usage:
'     Dim req As New CHttpRequest
'     req.Method = HTTP_METHOD_POST
'     req.Url = "http://example.com/api/items"
'     req.Headers.SetHeader "Content-Type", "application/json"
'     req.Body = "{""name"":""demo""}"
'
'     Dim client As New CHttpClient
'     Dim resp As CHttpResponse
'     Set resp = client.Execute(req)
'==============================================================

Public Method As String
Public Url As String
Public Body As String
Private m_Headers As CHttpHeaders

Private Sub Class_Initialize()
    Method = HTTP_METHOD_GET
    Url = ""
    Body = ""
    Set m_Headers = New CHttpHeaders
End Sub

Public Property Get Headers() As CHttpHeaders
    Set Headers = m_Headers
End Property

''' Builds the raw request-line + headers + body for ParsedUrl, filling
''' in Host/User-Agent/Connection/Content-Length automatically when the
''' caller has not already supplied them.
Public Function BuildRawRequest(ByRef ParsedUrl As CUrl) As String
    Dim hostHeader As String
    Dim raw As String

    If ParsedUrl.Port = HTTP_DEFAULT_PORT Then
        hostHeader = ParsedUrl.Host
    Else
        hostHeader = ParsedUrl.Host & ":" & ParsedUrl.Port
    End If

    If Not m_Headers.Exists("Host") Then m_Headers.SetHeader "Host", hostHeader
    If Not m_Headers.Exists("User-Agent") Then m_Headers.SetHeader "User-Agent", HTTP_USER_AGENT
    If Not m_Headers.Exists("Connection") Then m_Headers.SetHeader "Connection", "close"
    If Len(Body) > 0 And Not m_Headers.Exists("Content-Length") Then
        m_Headers.SetHeader "Content-Length", CStr(Len(Body))
    End If

    raw = UCase$(Method) & " " & ParsedUrl.Path & " " & HTTP_VERSION & HTTP_CRLF
    raw = raw & m_Headers.ToRawHeaders & HTTP_CRLF
    raw = raw & Body

    BuildRawRequest = raw
End Function
