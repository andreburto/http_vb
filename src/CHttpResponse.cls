VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CHttpResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==============================================================
' CHttpResponse
'
' Holds the parsed result of an HTTP request: status line,
' headers and body. Instances are created and populated by
' CHttpClient - consumers only need to read the public members.
'
' Usage:
'     Dim resp As CHttpResponse
'     Set resp = client.HttpGet("http://example.com/")
'     Debug.Print resp.StatusCode & " " & resp.StatusText
'     Debug.Print resp.Headers.GetHeader("Content-Type")
'     Debug.Print resp.Body
'==============================================================

Public StatusCode As Integer
Public StatusText As String
Public HttpVersion As String
Public Body As String
Private m_Headers As CHttpHeaders

Private Sub Class_Initialize()
    StatusCode = 0
    StatusText = ""
    HttpVersion = ""
    Body = ""
    Set m_Headers = New CHttpHeaders
End Sub

Public Property Get Headers() As CHttpHeaders
    Set Headers = m_Headers
End Property

''' Parses a complete raw HTTP response (status line, headers, body) as
''' received from the socket layer, decoding chunked transfer-encoding
''' and trimming to Content-Length when present.
Public Sub Parse(ByVal RawResponse As String)
    Dim headerEnd As Long
    Dim rawHeaderBlock As String
    Dim headerLines() As String
    Dim statusParts() As String
    Dim i As Long
    Dim colonPos As Long
    Dim name As String
    Dim value As String
    Dim contentLength As Long

    headerEnd = InStr(1, RawResponse, HTTP_CRLF & HTTP_CRLF)
    If headerEnd = 0 Then
        Err.Raise vbObjectError + 1200, "CHttpResponse", "Response is missing the blank line that separates headers from the body"
    End If

    rawHeaderBlock = Left$(RawResponse, headerEnd - 1)
    Body = Mid$(RawResponse, headerEnd + 4)

    modStringUtils.SplitString rawHeaderBlock, HTTP_CRLF, headerLines

    ' First line is the status line, e.g. "HTTP/1.1 200 OK"
    modStringUtils.SplitString headerLines(0), " ", statusParts
    If UBound(statusParts) < 1 Then
        Err.Raise vbObjectError + 1201, "CHttpResponse", "Malformed status line: '" & headerLines(0) & "'"
    End If

    HttpVersion = statusParts(0)
    StatusCode = CInt(statusParts(1))

    StatusText = ""
    For i = 2 To UBound(statusParts)
        If i > 2 Then StatusText = StatusText & " "
        StatusText = StatusText & statusParts(i)
    Next i

    For i = 1 To UBound(headerLines)
        If Len(headerLines(i)) > 0 Then
            colonPos = InStr(1, headerLines(i), ":")
            If colonPos > 0 Then
                name = Trim$(Left$(headerLines(i), colonPos - 1))
                value = Trim$(Mid$(headerLines(i), colonPos + 1))
                m_Headers.SetHeader name, value
            End If
        End If
    Next i

    If LCase$(m_Headers.GetHeader("Transfer-Encoding")) = "chunked" Then
        Body = DecodeChunkedBody(Body)
    ElseIf m_Headers.Exists("Content-Length") Then
        contentLength = CLng(m_Headers.GetHeader("Content-Length"))
        If contentLength >= 0 And contentLength <= Len(Body) Then
            Body = Left$(Body, contentLength)
        End If
    End If
End Sub

''' Decodes an HTTP/1.1 chunked-transfer-encoded body.
Private Function DecodeChunkedBody(ByVal Encoded As String) As String
    Dim pos As Long
    Dim result As String
    Dim lineEnd As Long
    Dim sizeLine As String
    Dim chunkSize As Long
    Dim semiPos As Long

    pos = 1
    result = ""

    Do
        lineEnd = InStr(pos, Encoded, HTTP_CRLF)
        If lineEnd = 0 Then Exit Do

        sizeLine = Mid$(Encoded, pos, lineEnd - pos)
        semiPos = InStr(1, sizeLine, ";")
        If semiPos > 0 Then sizeLine = Left$(sizeLine, semiPos - 1)
        sizeLine = Trim$(sizeLine)

        If Len(sizeLine) = 0 Then Exit Do
        chunkSize = CLng("&H" & sizeLine)

        If chunkSize = 0 Then Exit Do

        pos = lineEnd + 2
        result = result & Mid$(Encoded, pos, chunkSize)
        pos = pos + chunkSize + 2 ' skip the CRLF that follows each chunk's data
    Loop

    DecodeChunkedBody = result
End Function
