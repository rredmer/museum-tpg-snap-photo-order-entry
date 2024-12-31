Attribute VB_Name = "WinAPI"

Public Const MAX_DEMO_SHOTS As Integer = 10
Public MachineId As String

Private Type LUID
  LowPart As Long
  HighPart As Long
End Type

Private Type TOKEN_PRIVILEGES
   PrivilegeCount As Long
   LuidUDT As LUID
  Attributes As Long
End Type

Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2

Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4
Option Explicit

' Declarations needed for GetAdaptersInfo & GetIfTable
Private Const MIB_IF_TYPE_OTHER                   As Long = 1
Private Const MIB_IF_TYPE_ETHERNET                As Long = 6
Private Const MIB_IF_TYPE_TOKENRING               As Long = 9
Private Const MIB_IF_TYPE_FDDI                    As Long = 15
Private Const MIB_IF_TYPE_PPP                     As Long = 23
Private Const MIB_IF_TYPE_LOOPBACK                As Long = 24
Private Const MIB_IF_TYPE_SLIP                    As Long = 28

Private Const MIB_IF_ADMIN_STATUS_UP              As Long = 1
Private Const MIB_IF_ADMIN_STATUS_DOWN            As Long = 2
Private Const MIB_IF_ADMIN_STATUS_TESTING         As Long = 3

Private Const MIB_IF_OPER_STATUS_NON_OPERATIONAL  As Long = 0
Private Const MIB_IF_OPER_STATUS_UNREACHABLE      As Long = 1
Private Const MIB_IF_OPER_STATUS_DISCONNECTED     As Long = 2
Private Const MIB_IF_OPER_STATUS_CONNECTING       As Long = 3
Private Const MIB_IF_OPER_STATUS_CONNECTED        As Long = 4
Private Const MIB_IF_OPER_STATUS_OPERATIONAL      As Long = 5

Private Const MAX_ADAPTER_DESCRIPTION_LENGTH      As Long = 128
Private Const MAX_ADAPTER_DESCRIPTION_LENGTH_p    As Long = MAX_ADAPTER_DESCRIPTION_LENGTH + 4
Private Const MAX_ADAPTER_NAME_LENGTH             As Long = 256
Private Const MAX_ADAPTER_NAME_LENGTH_p           As Long = MAX_ADAPTER_NAME_LENGTH + 4
Private Const MAX_ADAPTER_ADDRESS_LENGTH          As Long = 8
Private Const DEFAULT_MINIMUM_ENTITIES            As Long = 32
Private Const MAX_HOSTNAME_LEN                    As Long = 128
Private Const MAX_DOMAIN_NAME_LEN                 As Long = 128
Private Const MAX_SCOPE_ID_LEN                    As Long = 256

Private Const MAXLEN_IFDESCR                      As Long = 256
Private Const MAX_INTERFACE_NAME_LEN              As Long = MAXLEN_IFDESCR * 2
Private Const MAXLEN_PHYSADDR                     As Long = 8

' Information structure returned by GetIfEntry/GetIfTable
Private Type MIB_IFROW
    wszName(0 To MAX_INTERFACE_NAME_LEN - 1) As Byte    ' MSDN Docs say pointer, but it is WCHAR array
    dwIndex             As Long
    dwType              As Long
    dwMtu               As Long
    dwSpeed             As Long
    dwPhysAddrLen       As Long
    bPhysAddr(MAXLEN_PHYSADDR - 1) As Byte
    dwAdminStatus       As Long
    dwOperStatus        As Long
    dwLastChange        As Long
    dwInOctets          As Long
    dwInUcastPkts       As Long
    dwInNUcastPkts      As Long
    dwInDiscards        As Long
    dwInErrors          As Long
    dwInUnknownProtos   As Long
    dwOutOctets         As Long
    dwOutUcastPkts      As Long
    dwOutNUcastPkts     As Long
    dwOutDiscards       As Long
    dwOutErrors         As Long
    dwOutQLen           As Long
    dwDescrLen          As Long
    bDescr As String * MAXLEN_IFDESCR
End Type

Private Type TIME_t
    aTime As Long
End Type

Private Type IP_ADDRESS_STRING
    IPadrString     As String * 16
End Type

Private Type IP_ADDR_STRING
    AdrNext         As Long
    IpAddress       As IP_ADDRESS_STRING
    IpMask          As IP_ADDRESS_STRING
    NTEcontext      As Long
End Type

' Information structure returned by GetIfEntry/GetIfTable
Private Type IP_ADAPTER_INFO
    Next As Long
    ComboIndex As Long
    AdapterName         As String * MAX_ADAPTER_NAME_LENGTH_p
    Description         As String * MAX_ADAPTER_DESCRIPTION_LENGTH_p
    MACadrLength        As Long
    MACaddress(0 To MAX_ADAPTER_ADDRESS_LENGTH - 1) As Byte
    AdapterIndex        As Long
    AdapterType         As Long             ' MSDN Docs say "UInt", but is 4 bytes
    DhcpEnabled         As Long             ' MSDN Docs say "UInt", but is 4 bytes
    CurrentIpAddress    As Long
    IpAddressList       As IP_ADDR_STRING
    GatewayList         As IP_ADDR_STRING
    DhcpServer          As IP_ADDR_STRING
    HaveWins            As Long             ' MSDN Docs say "Bool", but is 4 bytes
    PrimaryWinsServer   As IP_ADDR_STRING
    SecondaryWinsServer As IP_ADDR_STRING
    LeaseObtained       As TIME_t
    LeaseExpires        As TIME_t
End Type

     
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal numbytes As Long)

Public Declare Function GetAdaptersInfo Lib "iphlpapi.dll" (ByRef pAdapterInfo As Any, ByRef pOutBufLen As Long) As Long
Public Declare Function GetNumberOfInterfaces Lib "iphlpapi.dll" (ByRef pdwNumIf As Long) As Long
Public Declare Function GetIfEntry Lib "iphlpapi.dll" (ByRef pIfRow As Any) As Long
Private Declare Function GetIfTable Lib "iphlpapi.dll" (ByRef pIfTable As Any, ByRef pdwSize As Long, ByVal bOrder As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As Any, ReturnLength As Any) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long

' Shut down windows, and optional reboot it
' if the 2nd argument is True, no WM_QUERYENDSESSION and WM_ENDSESSION
' messages are sent to active applications

Sub ShutDownWindows(ByVal Reboot As Boolean, Optional ByVal Force As Boolean)
   Dim hToken As Long
   Dim tp As TOKEN_PRIVILEGES
   Dim flags As Long
   
   ' Windows NT/2000 require a special treatment
   ' to ensure that the calling process has the
   ' privileges to shut down the system
   
   ' under NT the high-order bit (that is, the sign bit)
   ' of the value retured by GetVersion is cleared
   If GetVersion() >= 0 Then
       ' Open this process for adjusting its privileges
       OpenProcessToken GetCurrentProcess(), (TOKEN_ADJUST_PRIVILEGES Or _
           TOKEN_QUERY), hToken
       
       ' Get the LUID for shutdown privilege.
       ' retrieves the locally unique identifier (LUID) used
       ' to locally represent the specified privilege name
       ' (first argument = "" means the local system)
       LookupPrivilegeValue "", "SeShutdownPrivilege", tp.LuidUDT
       
       ' complete the TOKEN_PRIVILEGES structure with the # of
       ' privileges and the desired attribute
       tp.PrivilegeCount = 1
       tp.Attributes = SE_PRIVILEGE_ENABLED
       
       ' enables or disables privileges in the specified access token
       ' last 3 arguments are zero because we aren't interested
       ' in previous privilege attributes.
       AdjustTokenPrivileges hToken, False, tp, 0, ByVal 0&, ByVal 0&
   End If
   
   ' prepare shutdown flags
   flags = EWX_SHUTDOWN
   If Reboot Then flags = flags Or EWX_REBOOT
   If Force Then flags = flags Or EWX_FORCE
   
   ' finally, you can shut down Windows
   ExitWindowsEx flags, &HFFFF
   
End Sub



'-----------------------------------------------------------------------------------
' Get the system's MAC address(es) via GetAdaptersInfo API function (IPHLPAPI.DLL)
'
' Note: GetAdaptersInfo returns information about physical adapters
'-----------------------------------------------------------------------------------
Public Function GetMACs_AdaptInfo() As String

    Dim AdapInfo As IP_ADAPTER_INFO, bufLen As Long, sts As Long
    Dim retStr As String, numStructs%, i%, IPinfoBuf() As Byte, srcPtr As Long
    
    
    ' Get size of buffer to allocate
    sts = GetAdaptersInfo(AdapInfo, bufLen)
    If (bufLen = 0) Then Exit Function
    numStructs = bufLen / Len(AdapInfo)
    retStr = numStructs & " Adapter(s):" & vbCrLf
    
    ' reserve byte buffer & get it filled with adapter information
    ' !!! Don't Redim AdapInfo array of IP_ADAPTER_INFO,
    ' !!! because VB doesn't allocate it contiguous (padding/alignment)
    ReDim IPinfoBuf(0 To bufLen - 1) As Byte
    sts = GetAdaptersInfo(IPinfoBuf(0), bufLen)
    If (sts <> 0) Then Exit Function
    
    ' Copy IP_ADAPTER_INFO slices into UDT structure
    srcPtr = VarPtr(IPinfoBuf(0))
    For i = 0 To numStructs - 1
        If (srcPtr = 0) Then Exit For
'        CopyMemory AdapInfo, srcPtr, Len(AdapInfo)
        CopyMemory AdapInfo, ByVal srcPtr, Len(AdapInfo)
        
        ' Extract Ethernet MAC address
        With AdapInfo
            If (.AdapterType = MIB_IF_TYPE_ETHERNET) Then
                retStr = retStr & vbCrLf & "[" & i & "] " & sz2string(.Description) _
                        & vbCrLf & MAC2String(.MACaddress) & vbCrLf
            End If
        End With
        srcPtr = AdapInfo.Next
    Next i
    
    ' Return list of MAC address(es)
    GetMACs_AdaptInfo = retStr
    
End Function


'-----------------------------------------------------------------------------------
' Get the system's MAC address(es) via GetIfTable API function (IPHLPAPI.DLL)
'
' Note: GetIfTable returns information also about the virtual loopback adapter
'-----------------------------------------------------------------------------------
Public Function GetMACs_IfTable() As String
    
    Dim NumAdapts As Long, nRowSize As Long, i%, retStr As String
    Dim IfInfo As MIB_IFROW, IPinfoBuf() As Byte, bufLen As Long, sts As Long
    
    
    ' Get # of interfaces defined (sometimes 1 more than GetIfTable)
    sts = GetNumberOfInterfaces(NumAdapts)
    
    ' Get size of buffer to allocate
    sts = GetIfTable(ByVal 0&, bufLen, 1)
    If (bufLen = 0) Then Exit Function

    ' reserve byte buffer & get it filled with adapter information
    ReDim IPinfoBuf(0 To bufLen - 1) As Byte
    sts = GetIfTable(IPinfoBuf(0), bufLen, 1)
    If (sts <> 0) Then Exit Function
    
    NumAdapts = IPinfoBuf(0)
    nRowSize = Len(IfInfo)
    retStr = NumAdapts & " Interface(s):" & vbCrLf

    For i = 1 To NumAdapts
        ' copy one IfRow chunk of byte data into an MIB_IFROW structure
        Call CopyMemory(IfInfo, IPinfoBuf(4 + (i - 1) * nRowSize), nRowSize)
        
        ' Take adapter address if correct type
        With IfInfo
            retStr = retStr & vbCrLf & "[" & i & "] " & Left$(.bDescr, .dwDescrLen - 1) & vbCrLf
            If (.dwType = MIB_IF_TYPE_ETHERNET) Then
                retStr = retStr & MAC2String(.bPhysAddr) & vbCrLf
            End If
        End With
    Next i

    GetMACs_IfTable = retStr
    
End Function


' Convert a byte array containing a MAC address to a hex string
Private Function MAC2String(AdrArray() As Byte) As String
    Dim aStr As String, hexStr As String, i%
    
    For i = 0 To 5
        If (i > UBound(AdrArray)) Then
            hexStr = "00"
        Else
            hexStr = Hex$(AdrArray(i))
        End If
        
        If (Len(hexStr) < 2) Then hexStr = "0" & hexStr
        aStr = aStr & hexStr
        If (i < 5) Then aStr = aStr & "-"
    Next i
    
    MAC2String = aStr
    
End Function


' Convert a zero-terminated fixed string to a dynamic VB string
Private Function sz2string(ByVal szStr As String) As String
    sz2string = Left$(szStr, InStr(1, szStr, Chr$(0)) - 1)
End Function


