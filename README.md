# HTTP VB5

A library for issuing and parsing HTTP/1.1 requests from Visual Basic 5,
targeting a stock **Windows NT4** machine with **Internet Explorer 4**.
It only uses components that ship with that platform: the Winsock 1.1
API (`WSOCK32.DLL`), `KERNEL32.DLL`, and the standard VB5 runtime. There
is no dependency on Winsock 2 (`WS2_32.DLL`, added in NT4 SP4), the
Microsoft Winsock Control (`MSWINSCK.OCX`), or `Scripting.Dictionary`.

## Layout

All source lives in [src/](/d:/projects/http_vb/src):

| Layer | File | Purpose |
| --- | --- | --- |
| Project | [HttpVB5.vbp](/d:/projects/http_vb/src/HttpVB5.vbp) | VB5 ActiveX DLL project tying the modules/classes together |
| Sockets | [modWinsock.bas](/d:/projects/http_vb/src/modWinsock.bas) | Raw Winsock 1.1 API declarations, types and constants |
| Sockets | [CSocket.cls](/d:/projects/http_vb/src/CSocket.cls) | Blocking TCP socket: `Connect`, `SendText`, `ReceiveText`, `CloseSocket` |
| Utility | [modStringUtils.bas](/d:/projects/http_vb/src/modStringUtils.bas) | `Split`/`Join`/`Replace` replacements (VB5 predates these) |
| Utility | [modHttpConstants.bas](/d:/projects/http_vb/src/modHttpConstants.bas) | Shared HTTP constants (methods, timeouts, version, user agent) |
| HTTP | [CUrl.cls](/d:/projects/http_vb/src/CUrl.cls) | Parses `http://host[:port]/path?query` |
| HTTP | [CHttpHeaderPair.cls](/d:/projects/http_vb/src/CHttpHeaderPair.cls) | Internal name/value pair used by `CHttpHeaders` |
| HTTP | [CHttpHeaders.cls](/d:/projects/http_vb/src/CHttpHeaders.cls) | Case-insensitive, ordered header collection |
| HTTP | [CHttpRequest.cls](/d:/projects/http_vb/src/CHttpRequest.cls) | Describes an outgoing request (method, URL, headers, body) |
| HTTP | [CHttpResponse.cls](/d:/projects/http_vb/src/CHttpResponse.cls) | Parses the status line, headers and body (incl. chunked transfer-encoding) |
| HTTP | [CHttpClient.cls](/d:/projects/http_vb/src/CHttpClient.cls) | High level entry point: `HttpGet`, `HttpPost`, `Execute` |

The library is layered bottom-up, as called for by the project
instructions: `CSocket` is a self-contained TCP wrapper that knows
nothing about HTTP, and the `CHttp*` classes build request/response
handling on top of it.

## Usage

```vb
Dim client As New CHttpClient
Dim resp As CHttpResponse

' Simple GET
Set resp = client.HttpGet("http://example.com/")
Debug.Print resp.StatusCode & " " & resp.StatusText
Debug.Print resp.Headers.GetHeader("Content-Type")
Debug.Print resp.Body

' POST with a custom request
Dim req As New CHttpRequest
req.Method = HTTP_METHOD_POST
req.Url = "http://example.com/api/items"
req.Headers.SetHeader "Content-Type", "application/json"
req.Body = "{""name"":""demo""}"
Set resp = client.Execute(req)
```

## Known limitations

These are deliberate trade-offs to stay within what an NT4 + VB5 + IE4
machine can do, and to keep the initial version focused:

- **HTTP only, no HTTPS** - there is no SSL/TLS stack available without
  additional components, so `CUrl` rejects any scheme other than `http`.
- **No persistent connections** - every request sends `Connection: close`
  and opens a fresh socket; this keeps response reading simple (read
  until the peer closes) at the cost of one TCP handshake per request.
- **Text-oriented bodies** - `CSocket` exchanges `String` data (converted
  via `StrConv`), which is convenient for headers/HTML/JSON/XML but is
  not a fully binary-safe transport for arbitrary downloads.
- **Redirects are not followed automatically** - callers inspect
  `StatusCode` and `Headers.GetHeader("Location")` themselves.

## Requirements

- Windows NT4
- Internet Explorer 4
- VB5 development environment
