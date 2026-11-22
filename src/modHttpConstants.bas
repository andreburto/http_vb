Attribute VB_Name = "modHttpConstants"
Option Explicit

'==============================================================
' modHttpConstants
'
' Shared constants for the HTTP layer.
'==============================================================

Public Const HTTP_CRLF As String = vbCrLf
Public Const HTTP_VERSION As String = "HTTP/1.1"
Public Const HTTP_DEFAULT_PORT As Integer = 80
Public Const HTTP_USER_AGENT As String = "HTTP-VB5/1.0 (Windows NT4)"

Public Const HTTP_CONNECT_TIMEOUT_MS As Long = 15000
Public Const HTTP_RECEIVE_TIMEOUT_MS As Long = 30000

Public Const HTTP_METHOD_GET As String = "GET"
Public Const HTTP_METHOD_POST As String = "POST"
Public Const HTTP_METHOD_HEAD As String = "HEAD"
Public Const HTTP_METHOD_PUT As String = "PUT"
Public Const HTTP_METHOD_DELETE As String = "DELETE"
