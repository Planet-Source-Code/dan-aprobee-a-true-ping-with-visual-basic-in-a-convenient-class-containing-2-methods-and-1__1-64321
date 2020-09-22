VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const IP_STATUS_BASE = 11000
Private Const IP_SUCCESS = 0
Private Const IP_BUF_TOO_SMALL = (11000 + 1)
Private Const IP_DEST_NET_UNREACHABLE = (11000 + 2)
Private Const IP_DEST_HOST_UNREACHABLE = (11000 + 3)
Private Const IP_DEST_PROT_UNREACHABLE = (11000 + 4)
Private Const IP_DEST_PORT_UNREACHABLE = (11000 + 5)
Private Const IP_NO_RESOURCES = (11000 + 6)
Private Const IP_BAD_OPTION = (11000 + 7)
Private Const IP_HW_ERROR = (11000 + 8)
Private Const IP_PACKET_TOO_BIG = (11000 + 9)
Private Const IP_REQ_TIMED_OUT = (11000 + 10)
Private Const IP_BAD_REQ = (11000 + 11)
Private Const IP_BAD_ROUTE = (11000 + 12)
Private Const IP_TTL_EXPIRED_TRANSIT = (11000 + 13)
Private Const IP_TTL_EXPIRED_REASSEM = (11000 + 14)
Private Const IP_PARAM_PROBLEM = (11000 + 15)
Private Const IP_SOURCE_QUENCH = (11000 + 16)
Private Const IP_OPTION_TOO_BIG = (11000 + 17)
Private Const IP_BAD_DESTINATION = (11000 + 18)
Private Const IP_ADDR_DELETED = (11000 + 19)
Private Const IP_SPEC_MTU_CHANGE = (11000 + 20)
Private Const IP_MTU_CHANGE = (11000 + 21)
Private Const IP_UNLOAD = (11000 + 22)
Private Const IP_ADDR_ADDED = (11000 + 23)
Private Const IP_GENERAL_FAILURE = (11000 + 50)
Private Const MAX_IP_STATUS = 11000 + 50
Private Const IP_PENDING = (11000 + 255)
Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD = 1
Private Const SOCKET_ERROR = -1
Private Const MAX_WSADescription = 256
Private Const MAX_WSASYSStatus = 128

Private Type ICMP_OPTIONS
   Ttl As Byte
   Tos As Byte
   Flags As Byte
   OptionsSize As Byte
   OptionsData As Long
End Type

Dim ICMPOPT As ICMP_OPTIONS

Private Type ICMP_ECHO_REPLY
   Address As Long
   status As Long
   roundTripTime As Long
   DataSize As Integer
   Reserved As Integer
   DataPointer As Long
   Options As ICMP_OPTIONS
   Data As String * 250
End Type

Private Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLen As Integer
    hAddrList As Long
End Type

Private Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To MAX_WSADescription) As Byte
    szSystemStatus(0 To MAX_WSASYSStatus) As Byte
    wMaxSockets As Integer
    wMaxUDPDG As Integer
    dwVendorInfo As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Source As Any, ByVal nbytes As Long)
Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal szHost As String) As Long
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Private Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, ByVal RequestOptions As Long, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal Timeout As Long) As Long
Private Declare Function inet_ntoa Lib "wsock32" (ByVal addr As Long) As Long
Private Declare Function lstrcpyA Lib "kernel32" (ByVal RetVal As String, ByVal Ptr As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal Ptr As Any) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
 
Dim m_pingTimeout As Long

Event PingReturn(ipAddress As String, successStatus As String, roundTripMilliseconds As Long)

Public Property Get pingTimeout() As Long
  pingTimeout = m_pingTimeout
End Property
Public Property Let pingTimeout(ByVal vNewValue As Long)
  m_pingTimeout = vNewValue
End Property

Private Sub Class_Initialize()
  'set default ping timeout value of 1 second
  m_pingTimeout = 1000
End Sub



Private Function GetStatusCode(status As Long) As String
Dim msg As String

   Select Case status
     Case IP_SUCCESS
       msg = "success"
     Case Else
       msg = "failure"
   End Select
 
   GetStatusCode = msg

End Function

Private Function HiByte(ByVal wParam As Integer)
  HiByte = wParam \ &H100 And &HFF&
End Function

Private Function LoByte(ByVal wParam As Integer)
  LoByte = wParam And &HFF&
End Function

Private Function funcIsIpAddress( _
               sAddr As String, _
               Optional bCalledFromWithin As Boolean) _
               As String
          
Dim leftPartLarge  As String
Dim leftPartSmall  As String

 leftPartSmall = LCase$(Trim$(Left$(sAddr, 4)))
 leftPartLarge = LCase$(Trim$(Left$(sAddr, 11)))
 
 'if the left 4 letters of [sAddr] are
 ' "www." or "http" then we can assume we
 'need to convert [sAddr] to an ip address
 If leftPartSmall = "http" Or leftPartSmall = "www." Then
   If leftPartSmall = "http" Then
     'strip the "http://" part
     sAddr = Trim$(Mid$(sAddr, 8, Len(sAddr) - 7))
     
     If leftPartLarge = "http://www." Then
       If bCalledFromWithin Then
         funcIsIpAddress = sAddr
       Else
         funcIsIpAddress = funcGetIPFromHostName(sAddr)
       End If
     Else
       'insert a "www."
       sAddr = ("www." & sAddr)
       
       If bCalledFromWithin Then
         funcIsIpAddress = sAddr
       Else
         funcIsIpAddress = funcGetIPFromHostName(sAddr)
       End If
     End If
   Else 'if leftPartSmall ="www."
      If bCalledFromWithin Then
        funcIsIpAddress = sAddr
      Else
        funcIsIpAddress = funcGetIPFromHostName(sAddr)
      End If
   End If
 Else 'we can assume [sAddr] is already ip address
   funcIsIpAddress = sAddr
 End If
  
End Function

Public Function Ping(szAddress As String) As Long
Dim ECHO As ICMP_ECHO_REPLY
Dim hPort As Long
Dim dwAddress As Long
Dim sDataToSend As String
Dim iOpt As Long

 If SocketsInitialize Then
   'need to convert [sxAddress]to an ip address?
   szAddress = funcIsIpAddress(szAddress)
   
   sDataToSend = "Echo This"
   dwAddress = AddressStringToLong(szAddress)
   hPort = IcmpCreateFile()

   If IcmpSendEcho(hPort, dwAddress, sDataToSend, _
      Len(sDataToSend), 0, ECHO, Len(ECHO), m_pingTimeout) Then
     'the ping succeeded,.Status will be 0
     '.RoundTripTime is the time in ms for
     ' the ping to complete,.Data is the data returned
     '(NULL terminated)
     '.Address is the Ip address that actually replied
     '.DataSize is the size of the string in .Data
     Ping = ECHO.roundTripTime
   Else
     Ping = ECHO.status * -1
   End If
   
   With ECHO
     RaiseEvent PingReturn(szAddress, _
           GetStatusCode(.status), .roundTripTime)
   End With
   
   Call IcmpCloseHandle(hPort)
   Call SocketsCleanup
 Else
   RaiseEvent PingReturn(szAddress, "failure", 0)
 End If 'Sockets initialize()
End Function

Private Function AddressStringToLong(ByVal tmp As String) As Long
Dim i As Integer
Dim parts(1 To 4) As String

   i = 0
   'we have to extract each part of the
   '123.456.789.123 string, delimited by a period
   While InStr(tmp, ".") > 0
      i = i + 1
      parts(i) = Mid(tmp, 1, InStr(tmp, ".") - 1)
      tmp = Mid(tmp, InStr(tmp, ".") + 1)
   Wend

   i = i + 1
   parts(i) = tmp

   If i <> 4 Then
     AddressStringToLong = 0
     Exit Function
   End If

   'build the long value out of the
   'hex of the extracted strings
   AddressStringToLong = Val("&H" & Right("00" & Hex(parts(4)), 2) & _
             Right("00" & Hex(parts(3)), 2) & _
             Right("00" & Hex(parts(2)), 2) & _
             Right("00" & Hex(parts(1)), 2))

End Function
 
Function funcGetIPFromHostName( _
                      ByVal sHostName As String) _
                      As String
Dim ptrHosent As Long    'address of hostent structure
Dim ptrName As Long      'address of name pointer
Dim ptrAddress As Long   'address of address pointer
Dim ptrIPAddress As Long 'address of string holding final IP address
Dim dwAddress As Long    'the final IP address
 
 If SocketsInitialize Then
   'this function can be called from within this class
   'AND from outside this class since it is public.
   'If it is being called from within this class then
   'we know [sHostname] has already been formatted
   'properly.  Otherwise  we have to make sure
   '[shostName] is provided in the proper format which
   'is "www.address.com" and NOT "http://www.address.com"
   'OR "http://adress.com"
   If LCase$(Trim$(Left$(sHostName, 4))) <> "www." Then
      sHostName = funcIsIpAddress(sHostName, True)
   End If

   'converts a host name to an IP address.
   ptrHosent = gethostbyname(sHostName & vbNullChar)

   If ptrHosent <> 0 Then
     'assign pointer addresses and offset
     'ptrName is the official name of the host (PC).
     'If using the DNS or similar resolution system,
     'it is the Fully Qualified Domain Name (FQDN)
     'that caused the server to return a reply.
     'If using a local hosts file, it is the first
     'entry after the IP address.
      ptrName = ptrHosent
      
     'Null-terminated list of addresses for the host.
     'The Address is offset 12 bytes from the start of
     'the HOSENT structure. Addresses are returned
     'in network byte order.
      ptrAddress = (ptrHosent + 12)
      
     'get the actual IP address
      CopyMemory ptrAddress, ByVal ptrAddress, 4
      CopyMemory ptrIPAddress, ByVal ptrAddress, 4
      CopyMemory dwAddress, ByVal ptrIPAddress, 4
      funcGetIPFromHostName = GetIPFromAddress(dwAddress)
   End If
   
   SocketsCleanup
 End If

End Function

Private Function GetIPFromAddress(Address As Long) As String
Dim ptrString As Long
   
   ptrString = inet_ntoa(Address)
   GetIPFromAddress = GetStrFromPtrA(ptrString)
End Function

Private Function GetStrFromPtrA(ByVal lpszA As Long) As String

   GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
   Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)
End Function

Private Function SocketsInitialize() As Boolean
Dim WSAD As WSADATA
   
  'when the socket version returned == version
  'required, return True
   SocketsInitialize = WSAStartup(WS_VERSION_REQD, _
                      WSAD) = IP_SUCCESS
End Function

Private Sub SocketsCleanup()
 
  'clean up the sockets
  WSACleanup
End Sub
 

Private Sub Class_Terminate()
  SocketsCleanup
End Sub
