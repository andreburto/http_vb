VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CUrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==============================================================
' CUrl
'
' Minimal "http://host[:port]/path?query" parser. Only the plain
' http scheme is supported - there is no SSL/TLS stack available
' on a stock Windows NT4 + Internet Explorer 4 machine, so https
' URLs are rejected rather than silently mishandled.
'
' Usage:
'     Dim u As New CUrl
'     u.Parse "http://example.com:8080/path?x=1"
'     Debug.Print u.Host; u.Port; u.Path
'==============================================================

Public Scheme As String
Public Host As String
Public Port As Integer
Public Path As String

Private Sub Class_Initialize()
    Scheme = "http"
    Host = ""
    Port = HTTP_DEFAULT_PORT
    Path = "/"
End Sub

''' Parses Url and populates Scheme, Host, Port and Path.
Public Sub Parse(ByVal Url As String)
    Dim rest As String
    Dim schemeSep As Long
    Dim pathSep As Long
    Dim portSep As Long
    Dim hostPort As String

    schemeSep = InStr(1, Url, "://")
    If schemeSep = 0 Then
        Err.Raise vbObjectError + 1100, "CUrl", "'" & Url & "' is not an absolute URL (missing scheme)"
    End If

    Scheme = LCase$(Left$(Url, schemeSep - 1))
    If Scheme <> "http" Then
        Err.Raise vbObjectError + 1101, "CUrl", "Unsupported scheme '" & Scheme & "' - only plain http is supported on this platform"
    End If

    rest = Mid$(Url, schemeSep + 3)

    pathSep = InStr(1, rest, "/")
    If pathSep = 0 Then
        hostPort = rest
        Path = "/"
    Else
        hostPort = Left$(rest, pathSep - 1)
        Path = Mid$(rest, pathSep)
    End If

    If Len(hostPort) = 0 Then
        Err.Raise vbObjectError + 1102, "CUrl", "'" & Url & "' does not specify a host"
    End If

    portSep = InStr(1, hostPort, ":")
    If portSep = 0 Then
        Host = hostPort
        Port = HTTP_DEFAULT_PORT
    Else
        Host = Left$(hostPort, portSep - 1)
        Port = CInt(Mid$(hostPort, portSep + 1))
    End If
End Sub
