VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==============================================================
' CSocket
'
' Minimal blocking TCP socket wrapper built directly on the
' Winsock 1.1 API (WSOCK32.DLL, see modWinsock) so the library
' only relies on components already present on a stock Windows
' NT4 install. This is the foundation the HTTP classes are built
' on top of; it knows nothing about HTTP itself.
'
' Note: ReceiveText/SendText exchange plain Strings (converted to
' and from a Byte buffer internally via StrConv). This keeps the
' API simple for text based protocols such as HTTP headers, HTML,
' JSON and XML, but is not a fully binary-safe transport.
'
' Usage:
'     Dim sock As New CSocket
'     sock.Connect "www.example.com", 80, 10000
'     sock.SendText "GET / HTTP/1.0" & vbCrLf & vbCrLf
'     Do
'         Dim chunk As String
'         chunk = sock.ReceiveText(5000)
'         If Len(chunk) = 0 Then Exit Do
'         ' ... accumulate chunk ...
'     Loop While sock.Connected
'     sock.CloseSocket
'==============================================================

Private m_Socket As Long
Private m_Connected As Boolean

Private Sub Class_Initialize()
    modWinsock.EnsureStarted
    m_Socket = INVALID_SOCKET
    m_Connected = False
End Sub

Private Sub Class_Terminate()
    CloseSocket
    modWinsock.ReleaseStarted
End Sub

''' Returns True while the socket believes it holds an open connection.
Public Property Get Connected() As Boolean
    Connected = m_Connected
End Property

''' Opens a blocking TCP connection to Host:Port, raising an error if the
''' connection cannot be established within TimeoutMs milliseconds.
Public Sub Connect(ByVal Host As String, ByVal Port As Integer, Optional ByVal TimeoutMs As Long = 10000)
    Dim addr As SOCKADDR_IN
    Dim connectResult As Long
    Dim readfds As FD_SET_TYPE
    Dim writefds As FD_SET_TYPE
    Dim exceptfds As FD_SET_TYPE
    Dim tv As TIMEVAL
    Dim selectResult As Long
    Dim sockErr As Long
    Dim sockErrLen As Long
    Dim lastErr As Long
    Dim ioctlArg As Long

    CloseSocket

    m_Socket = wsock_socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    If m_Socket = INVALID_SOCKET Then
        Err.Raise vbObjectError + 1010, "CSocket", "socket() failed with error " & WSAGetLastError()
    End If

    addr.sin_family = AF_INET
    addr.sin_port = wsock_htons(CInt(Port))
    addr.sin_addr = ResolveAddress(Host)

    ' Switch to non-blocking mode so connect() can be bounded by a timeout.
    ioctlArg = 1
    wsock_ioctlsocket m_Socket, FIONBIO, ioctlArg

    connectResult = wsock_connect(m_Socket, addr, Len(addr))
    If connectResult = SOCKET_ERROR Then
        lastErr = WSAGetLastError()
        If lastErr <> WSAEWOULDBLOCK Then
            CloseSocket
            Err.Raise vbObjectError + 1011, "CSocket", "connect() failed with error " & lastErr
        End If

        writefds.fd_count = 1
        writefds.fd_array(0) = m_Socket
        exceptfds.fd_count = 1
        exceptfds.fd_array(0) = m_Socket
        tv.tv_sec = TimeoutMs \ 1000
        tv.tv_usec = (TimeoutMs Mod 1000) * 1000

        selectResult = wsock_select(0, readfds, writefds, exceptfds, tv)
        If selectResult <= 0 Then
            CloseSocket
            Err.Raise vbObjectError + 1012, "CSocket", "connect() timed out after " & TimeoutMs & " ms"
        End If

        sockErrLen = 4
        wsock_getsockopt m_Socket, SOL_SOCKET, SO_ERROR, sockErr, sockErrLen
        If sockErr <> 0 Then
            CloseSocket
            Err.Raise vbObjectError + 1013, "CSocket", "connect() failed with error " & sockErr
        End If
    End If

    ' Restore blocking mode; Send/Receive below use select() purely to
    ' bound how long we wait rather than to drive non-blocking I/O.
    ioctlArg = 0
    wsock_ioctlsocket m_Socket, FIONBIO, ioctlArg

    m_Connected = True
End Sub

''' Sends Data in full, looping until every byte has been accepted by
''' the socket (send() is permitted to accept less than the full buffer
''' in one call).
Public Sub SendText(ByVal Data As String)
    Dim bytes() As Byte
    Dim totalSent As Long
    Dim sent As Long
    Dim remaining As Long
    Dim sendErr As Long

    If Not m_Connected Then
        Err.Raise vbObjectError + 1030, "CSocket", "SendText called on a socket that is not connected"
    End If

    If Len(Data) = 0 Then Exit Sub

    bytes = StrConv(Data, vbFromUnicode)
    remaining = UBound(bytes) - LBound(bytes) + 1
    totalSent = 0

    Do While totalSent < remaining
        sent = wsock_send(m_Socket, bytes(LBound(bytes) + totalSent), remaining - totalSent, 0)
        If sent = SOCKET_ERROR Then
            sendErr = WSAGetLastError()
            m_Connected = False
            Err.Raise vbObjectError + 1031, "CSocket", "send() failed with error " & sendErr
        End If
        totalSent = totalSent + sent
    Loop
End Sub

''' Waits up to TimeoutMs milliseconds for data, then returns one buffer's
''' worth of text (up to MaxLength bytes). Returns "" when the peer has
''' closed the connection (check Connected to tell this apart from an
''' empty-but-still-open read, which cannot happen with this API).
Public Function ReceiveText(Optional ByVal TimeoutMs As Long = 10000, Optional ByVal MaxLength As Long = 8192) As String
    Dim readfds As FD_SET_TYPE
    Dim writefds As FD_SET_TYPE
    Dim exceptfds As FD_SET_TYPE
    Dim tv As TIMEVAL
    Dim selectResult As Long
    Dim buffer() As Byte
    Dim received As Long
    Dim selErr As Long
    Dim recvErr As Long

    If Not m_Connected Then
        Err.Raise vbObjectError + 1040, "CSocket", "ReceiveText called on a socket that is not connected"
    End If

    readfds.fd_count = 1
    readfds.fd_array(0) = m_Socket
    tv.tv_sec = TimeoutMs \ 1000
    tv.tv_usec = (TimeoutMs Mod 1000) * 1000

    selectResult = wsock_select(0, readfds, writefds, exceptfds, tv)
    If selectResult = SOCKET_ERROR Then
        selErr = WSAGetLastError()
        m_Connected = False
        Err.Raise vbObjectError + 1041, "CSocket", "select() failed with error " & selErr
    ElseIf selectResult = 0 Then
        Err.Raise vbObjectError + 1042, "CSocket", "ReceiveText timed out after " & TimeoutMs & " ms"
    End If

    ReDim buffer(0 To MaxLength - 1)
    received = wsock_recv(m_Socket, buffer(0), MaxLength, 0)

    If received = SOCKET_ERROR Then
        recvErr = WSAGetLastError()
        m_Connected = False
        Err.Raise vbObjectError + 1043, "CSocket", "recv() failed with error " & recvErr
    ElseIf received = 0 Then
        ' Peer closed the connection gracefully.
        m_Connected = False
        ReceiveText = ""
    Else
        ReDim Preserve buffer(0 To received - 1)
        ReceiveText = StrConv(buffer, vbUnicode)
    End If
End Function

''' Closes the socket if open. Safe to call more than once.
Public Sub CloseSocket()
    If m_Socket <> INVALID_SOCKET Then
        wsock_shutdown m_Socket, SD_BOTH
        wsock_closesocket m_Socket
        m_Socket = INVALID_SOCKET
    End If
    m_Connected = False
End Sub

''' Resolves Host to a network-order IPv4 address, trying a literal
''' dotted-quad first and falling back to gethostbyname() for DNS names.
Private Function ResolveAddress(ByVal Host As String) As Long
    Dim addr As Long
    Dim hostentPtr As Long
    Dim hostentData As HOSTENT
    Dim firstAddrPtr As Long
    Dim resolvedAddr As Long

    addr = wsock_inet_addr(Host)
    If addr <> INVALID_SOCKET Then
        ResolveAddress = addr
        Exit Function
    End If

    hostentPtr = wsock_gethostbyname(Host)
    If hostentPtr = 0 Then
        Err.Raise vbObjectError + 1020, "CSocket", "Unable to resolve host '" & Host & "' (error " & WSAGetLastError() & ")"
    End If

    CopyMemory hostentData, ByVal hostentPtr, Len(hostentData)
    CopyMemory firstAddrPtr, ByVal hostentData.h_addr_list, 4
    CopyMemory resolvedAddr, ByVal firstAddrPtr, 4

    ResolveAddress = resolvedAddr
End Function
