VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CHttpClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==============================================================
' CHttpClient
'
' High level entry point for the library: opens a socket to the
' target host, sends an HTTP request and returns the parsed
' response. One instance can be reused for multiple requests;
' each call opens and closes its own connection (no keep-alive,
' since the library always sends "Connection: close").
'
' Usage:
'     Dim client As New CHttpClient
'     Dim resp As CHttpResponse
'     Set resp = client.HttpGet("http://example.com/")
'     Debug.Print resp.StatusCode & " " & resp.StatusText
'     Debug.Print resp.Body
'==============================================================

Public ConnectTimeoutMs As Long
Public ReceiveTimeoutMs As Long

Private Sub Class_Initialize()
    ConnectTimeoutMs = HTTP_CONNECT_TIMEOUT_MS
    ReceiveTimeoutMs = HTTP_RECEIVE_TIMEOUT_MS
End Sub

''' Convenience wrapper for a simple GET request.
Public Function HttpGet(ByVal Url As String) As CHttpResponse
    Dim request As New CHttpRequest
    request.Method = HTTP_METHOD_GET
    request.Url = Url
    Set HttpGet = Execute(request)
End Function

''' Convenience wrapper for a simple POST request.
Public Function HttpPost(ByVal Url As String, ByVal Body As String, Optional ByVal ContentType As String = "application/x-www-form-urlencoded") As CHttpResponse
    Dim request As New CHttpRequest
    request.Method = HTTP_METHOD_POST
    request.Url = Url
    request.Body = Body
    request.Headers.SetHeader "Content-Type", ContentType
    Set HttpPost = Execute(request)
End Function

''' Sends Request and returns the parsed response. Raises an error if the
''' connection cannot be established or the response cannot be read.
Public Function Execute(ByRef Request As CHttpRequest) As CHttpResponse
    Dim parsedUrl As New CUrl
    Dim sock As New CSocket
    Dim rawRequest As String
    Dim rawResponse As String
    Dim chunk As String
    Dim response As New CHttpResponse

    parsedUrl.Parse Request.Url
    rawRequest = Request.BuildRawRequest(parsedUrl)

    sock.Connect parsedUrl.Host, parsedUrl.Port, ConnectTimeoutMs
    sock.SendText rawRequest

    rawResponse = ""
    Do
        chunk = sock.ReceiveText(ReceiveTimeoutMs)
        rawResponse = rawResponse & chunk
    Loop Until Len(chunk) = 0 Or Not sock.Connected

    sock.CloseSocket

    response.Parse rawResponse
    Set Execute = response
End Function
