Attribute VB_Name = "modWinsock"
Option Explicit

'==============================================================
' modWinsock
'
' Low-level Winsock 1.1 API declarations (WSOCK32.DLL). This is
' the Winsock stack shipped with the base install of Windows
' NT4 - Winsock 2 (WS2_32.DLL) was only introduced with NT4
' Service Pack 4, so it is deliberately not used here.
'
' Nothing in this module should be called directly by consumers
' of the library; it exists purely as the foundation CSocket is
' built on top of.
'==============================================================

Public Const AF_INET As Long = 2
Public Const SOCK_STREAM As Long = 1
Public Const IPPROTO_TCP As Long = 6
Public Const INVALID_SOCKET As Long = -1
Public Const SOCKET_ERROR As Long = -1
Public Const SD_BOTH As Long = 2

Public Const SOL_SOCKET As Long = &HFFFF&
Public Const SO_ERROR As Long = &H1007&
Public Const FIONBIO As Long = &H8004667E

Public Const WSAEWOULDBLOCK As Long = 10035

Public Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 256) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Public Type SOCKADDR_IN
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero(0 To 7) As Byte
End Type

Public Type HOSTENT
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Public Type TIMEVAL
    tv_sec As Long
    tv_usec As Long
End Type

Public Type FD_SET_TYPE
    fd_count As Long
    fd_array(0 To 63) As Long
End Type

' -- Winsock 1.1 (WSOCK32.DLL) -----------------------------------------
Public Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired As Integer, ByRef lpWSAData As WSADATA) As Long
Public Declare Function WSACleanup Lib "wsock32.dll" () As Long
Public Declare Function WSAGetLastError Lib "wsock32.dll" () As Long

Public Declare Function wsock_socket Lib "wsock32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
Public Declare Function wsock_connect Lib "wsock32.dll" Alias "connect" (ByVal s As Long, ByRef addr As SOCKADDR_IN, ByVal namelen As Long) As Long
Public Declare Function wsock_send Lib "wsock32.dll" Alias "send" (ByVal s As Long, ByRef buf As Byte, ByVal buflen As Long, ByVal flags As Long) As Long
Public Declare Function wsock_recv Lib "wsock32.dll" Alias "recv" (ByVal s As Long, ByRef buf As Byte, ByVal buflen As Long, ByVal flags As Long) As Long
Public Declare Function wsock_closesocket Lib "wsock32.dll" Alias "closesocket" (ByVal s As Long) As Long
Public Declare Function wsock_shutdown Lib "wsock32.dll" Alias "shutdown" (ByVal s As Long, ByVal how As Long) As Long
Public Declare Function wsock_ioctlsocket Lib "wsock32.dll" Alias "ioctlsocket" (ByVal s As Long, ByVal cmd As Long, ByRef argp As Long) As Long
Public Declare Function wsock_getsockopt Lib "wsock32.dll" Alias "getsockopt" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, ByRef optval As Long, ByRef optlen As Long) As Long
Public Declare Function wsock_htons Lib "wsock32.dll" Alias "htons" (ByVal hostshort As Integer) As Integer
Public Declare Function wsock_inet_addr Lib "wsock32.dll" Alias "inet_addr" (ByVal cp As String) As Long
Public Declare Function wsock_gethostbyname Lib "wsock32.dll" Alias "gethostbyname" (ByVal host_name As String) As Long
Public Declare Function wsock_select Lib "wsock32.dll" Alias "select" (ByVal nfds As Long, ByRef readfds As FD_SET_TYPE, ByRef writefds As FD_SET_TYPE, ByRef exceptfds As FD_SET_TYPE, ByRef timeout As TIMEVAL) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

' -- Reference counted startup/cleanup ---------------------------------
' WSAStartup/WSACleanup must be paired, but several CSocket instances
' may be alive at once, so we keep a simple module-level counter rather
' than requiring callers to manage Winsock lifetime themselves.
Private m_InitCount As Long

Public Sub EnsureStarted()
    Dim wsaData As WSADATA
    Dim lReturn As Long

    If m_InitCount = 0 Then
        lReturn = WSAStartup(&H101, wsaData) ' request Winsock 1.1
        If lReturn <> 0 Then
            Err.Raise vbObjectError + 1000, "modWinsock", "WSAStartup failed with error " & lReturn
        End If
    End If

    m_InitCount = m_InitCount + 1
End Sub

Public Sub ReleaseStarted()
    If m_InitCount > 0 Then
        m_InitCount = m_InitCount - 1
        If m_InitCount = 0 Then
            WSACleanup
        End If
    End If
End Sub
