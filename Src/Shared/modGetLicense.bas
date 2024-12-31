Attribute VB_Name = "modGetLicense"
Option Explicit
Public FS As New Scripting.FileSystemObject

'--- Identifying Numbers
Public MacStr As String
Public SerStr As String
Public ProcessorID As String
Public BrandName As String

'Public DemoMode As Boolean
'Public DemoExceeded As Boolean
Private Const MIN_ASC = 32     ' Space.
Private Const MAX_ASC = 126    ' ~.
Private Const NUM_ASC = MAX_ASC - MIN_ASC + 1
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private RegKey As String
Private LicKey As String
Public RegFile As String
Public LicFile As String


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

Public Sub SetKeys(RegistrationKey As String, LicenseKey As String, fname As String)
    Dim MStr As String, Arr As Variant, GetProcessorID As String, objWMIService As Object, colItems As Object, objItem As Object, ts As TextStream
    Dim RegU As String, RegL As String
    
    '---- Store keys local
    RegKey = RegistrationKey
    LicKey = LicenseKey
    
    '---- Get Computer Attributes to Public
    MStr = GetMACs_AdaptInfo()
    Arr = Split(MStr, vbCrLf)
    MacStr = Trim(Arr(UBound(Arr) - 1))
    SerStr = FS.Drives("C:").SerialNumber
    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystemProduct", , 48)
    For Each objItem In colItems
        BrandName = Trim(objItem.Name)
        ProcessorID = objItem.identifyingnumber
    Next
    RegFile = App.Path & "\" & fname
    LicFile = App.Path & "\" & fname & ".LIC"
    
    '---- Write registration file if it does not exist
    If FS.FileExists(RegFile) = False Then
        Set ts = FS.CreateTextFile(RegFile, False, False)
        RegU = App.CompanyName & "," & App.ProductName & "," & BrandName & "," & ProcessorID & "," & SerStr & "," & MacStr
        Cipher RegKey, RegU, RegL
        ts.Write RegL
        ts.Close
        Set ts = Nothing
    End If
    
End Sub

Public Function ReadLicFile() As Boolean
    Dim ts As TextStream, lkey As String, ukey As String, vkey As Variant
    If FS.FileExists(LicFile) Then
        Set ts = FS.OpenTextFile(LicFile, ForReading, False, TristateFalse)
        lkey = ts.ReadAll
        ts.Close
        Set ts = Nothing
        Decipher LicKey, lkey, ukey
        vkey = Split(ukey, ",")
        If UBound(vkey) = 7 Then
            If vkey(5) = SerStr And vkey(6) = MacStr And Trim(vkey(2)) = Trim(App.ProductName) And vkey(0) = vkey(7) Then
                '---- Check license expiration date
                If Date < #2/1/2009# Then
                    'DemoMode = False
                    'DemoExceeded = False
                    ReadLicFile = True
                Else
                    FS.DeleteFile LicFile, True
                    MsgBox "License has expired (2/1/2009).", vbApplicationModal + vbCritical + vbOKOnly, "ERROR"
                End If
                Exit Function
            End If
        Else
            MsgBox "Invalid license file or license has expired.", vbApplicationModal + vbCritical + vbOKOnly, "ERROR"
        End If
    End If
    ReadLicFile = False
End Function

Public Function GetComputerID() As String
    'GetComputerID = Right(Trim(FS.GetDrive("C:").SerialNumber), 4)
    GetComputerID = Mid(MacStr, 13, 2) & Right(MacStr, 2)
End Function

Public Function ComputerName() As String
   Dim strBuffer As String * 255
   If GetComputerName(strBuffer, 255&) <> 0 Then
      ComputerName = Left(strBuffer, InStr(strBuffer, vbNullChar) - 1)
   Else
      ComputerName = "N/A"
   End If
End Function

Public Sub Cipher(ByVal password As String, ByVal from_text As String, to_text As String)
    Dim offset As Long, str_len As Integer, i As Integer, ch As Integer
    offset = NumericPassword(password)          ' Initialize the random number generator.
    Rnd -1
    Randomize offset
    str_len = Len(from_text)
    For i = 1 To str_len
        ch = Asc(Mid$(from_text, i, 1))
        If ch >= MIN_ASC And ch <= MAX_ASC Then
            ch = ch - MIN_ASC
            offset = Int((NUM_ASC + 1) * Rnd)
            ch = ((ch + offset) Mod NUM_ASC)
            ch = ch + MIN_ASC
            to_text = to_text & Chr$(ch)
        End If
    Next i
End Sub

Public Sub Decipher(ByVal password As String, ByVal from_text As String, to_text As String)
    Dim offset As Long, str_len As Integer, i As Integer, ch As Integer
    offset = NumericPassword(password)          'Initialize the random number generator.
    Rnd -1
    Randomize offset
    str_len = Len(from_text)
    For i = 1 To str_len
        ch = Asc(Mid$(from_text, i, 1))
        If ch >= MIN_ASC And ch <= MAX_ASC Then
            ch = ch - MIN_ASC
            offset = Int((NUM_ASC + 1) * Rnd)
            ch = ((ch - offset) Mod NUM_ASC)
            If ch < 0 Then ch = ch + NUM_ASC
            ch = ch + MIN_ASC
            to_text = to_text & Chr$(ch)
        End If
    Next i
End Sub

' Translate a password into an offset value.
Private Function NumericPassword(ByVal password As String) As Long
    Dim value As Long, ch As Long, shift1 As Long, shift2 As Long, i As Integer, str_len As Integer
    str_len = Len(password)
    For i = 1 To str_len
        ch = Asc(Mid$(password, i, 1))
        value = value Xor (ch * 2 ^ shift1)
        value = value Xor (ch * 2 ^ shift2)
        shift1 = (shift1 + 7) Mod 19
        shift2 = (shift2 + 13) Mod 23
    Next i
    NumericPassword = value
End Function

Public Function CheckForUpdate(UpdatePath As String) As Boolean
    CheckForUpdate = False
    If FS.FileExists(UpdatePath) Then
        '---- Update folder found
        Dim Vers As String, Parts As Variant
        Vers = FS.GetFileVersion(UpdatePath)
        Parts = Split(Vers, ".")
        If UBound(Parts) = 3 Then
            If CInt(Parts(0)) > CInt(App.Major) Then
                'Major Update
                CheckForUpdate = True
            ElseIf CInt(Parts(0)) = CInt(App.Major) Then
                If CInt(Parts(1)) > CInt(App.Minor) Then
                    'Minor update
                    CheckForUpdate = True
                ElseIf CInt(Parts(1)) = CInt(App.Minor) Then
                    If CInt(Parts(3)) > CInt(App.Revision) Then
                        'Revision update
                        CheckForUpdate = True
                    End If
                End If
            End If
        End If
    End If
End Function

Public Function DoUpdate(NewFile As String) As Boolean
    Dim OldFile As String, CmdLine As String
    OldFile = Trim(App.Path) & "\" & Trim(App.EXEName) & ".exe"
    CmdLine = Trim(App.Path) & "\" & "UPDATE.EXE " & OldFile & "," & NewFile
    Shell CmdLine, vbNormalFocus
    End
End Function
