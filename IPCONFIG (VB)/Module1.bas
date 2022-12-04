Attribute VB_Name = "Module1"
Public Const MAX_HOSTNAME_LEN = 132
Public Const MAX_DOMAIN_NAME_LEN = 132
Public Const MAX_SCOPE_ID_LEN = 260
Public Const MAX_ADAPTER_NAME_LENGTH = 260
Public Const MAX_ADAPTER_ADDRESS_LENGTH = 8
Public Const MAX_ADAPTER_DESCRIPTION_LENGTH = 132
Public Const ERROR_BUFFER_OVERFLOW = 111
Public Const MIB_IF_TYPE_ETHERNET = 1
Public Const MIB_IF_TYPE_TOKENRING = 2
Public Const MIB_IF_TYPE_FDDI = 3
Public Const MIB_IF_TYPE_PPP = 4
Public Const MIB_IF_TYPE_LOOPBACK = 5
Public Const MIB_IF_TYPE_SLIP = 6

Type IP_ADDR_STRING
            Next As Long
            IpAddress As String * 16
            IpMask As String * 16
            Context As Long
End Type

Type IP_ADAPTER_INFO
            Next As Long
            ComboIndex As Long
            AdapterName As String * MAX_ADAPTER_NAME_LENGTH
            Description As String * MAX_ADAPTER_DESCRIPTION_LENGTH
            AddressLength As Long
            Address(MAX_ADAPTER_ADDRESS_LENGTH - 1) As Byte
            Index As Long
            Type As Long
            DhcpEnabled As Long
            CurrentIpAddress As Long
            IpAddressList As IP_ADDR_STRING
            GatewayList As IP_ADDR_STRING
            DhcpServer As IP_ADDR_STRING
            HaveWins As Boolean
            PrimaryWinsServer As IP_ADDR_STRING
            SecondaryWinsServer As IP_ADDR_STRING
            LeaseObtained As Long
            LeaseExpires As Long
End Type

Type FIXED_INFO
            HostName As String * MAX_HOSTNAME_LEN
            DomainName As String * MAX_DOMAIN_NAME_LEN
            CurrentDnsServer As Long
            DnsServerList As IP_ADDR_STRING
            NodeType As Long
            ScopeId  As String * MAX_SCOPE_ID_LEN
            EnableRouting As Long
            EnableProxy As Long
            EnableDns As Long
End Type

Public Declare Function GetNetworkParams Lib "IPHlpApi" (FixedInfo As Any, pOutBufLen As Long) As Long
Public Declare Function GetAdaptersInfo Lib "IPHlpApi" (IpAdapterInfo As Any, pOutBufLen As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Sub main()
    Dim error As Long
    Dim FixedInfoSize As Long
    Dim AdapterInfoSize As Long
    Dim i As Integer
    Dim PhysicalAddress  As String
    Dim NewTime As Date
    Dim AdapterInfo As IP_ADAPTER_INFO
    Dim Adapt As IP_ADAPTER_INFO
    Dim AddrStr As IP_ADDR_STRING
    Dim FixedInfo As FIXED_INFO
    Dim Buffer As IP_ADDR_STRING
    Dim pAddrStr As Long
    Dim pAdapt As Long
    Dim Buffer2 As IP_ADAPTER_INFO
    Dim FixedInfoBuffer() As Byte
    Dim AdapterInfoBuffer() As Byte
    
    'Get the main IP configuration information for this machine using a FIXED_INFO structure
    FixedInfoSize = 0
    error = GetNetworkParams(ByVal 0&, FixedInfoSize)
    If error <> 0 Then
        If error <> ERROR_BUFFER_OVERFLOW Then
           MsgBox "GetNetworkParams sizing failed with error " & error
           Exit Sub
        End If
    End If
    ReDim FixedInfoBuffer(FixedInfoSize - 1)
    

    error = GetNetworkParams(FixedInfoBuffer(0), FixedInfoSize)
    If error = 0 Then
            CopyMemory FixedInfo, FixedInfoBuffer(0), Len(FixedInfo)
            MsgBox "Host Name:  " & FixedInfo.HostName
            MsgBox "DNS Servers:  " & FixedInfo.DnsServerList.IpAddress
            pAddrStr = FixedInfo.DnsServerList.Next
            Do While pAddrStr <> 0
                  CopyMemory Buffer, ByVal pAddrStr, Len(Buffer)
                  MsgBox "DNS Servers:  " & Buffer.IpAddress
                  pAddrStr = Buffer.Next
            Loop
            
            Select Case FixedInfo.NodeType
                       Case 1
                                  MsgBox "Node type: Broadcast"
                       Case 2
                                   MsgBox "Node type: Peer to peer"
                       Case 4
                                    MsgBox "Node type: Mixed"
                       Case 8
                                    MsgBox "Node type: Hybrid"
                       Case Else
                                    MsgBox "Unknown node type"
            End Select
            
            MsgBox "NetBIOS Scope ID:  " & FixedInfo.ScopeId
            If FixedInfo.EnableRouting Then
                       MsgBox "IP Routing Enabled "
            Else
                       MsgBox "IP Routing not enabled"
            End If
            If FixedInfo.EnableProxy Then
                       MsgBox "WINS Proxy Enabled "
            Else
                       MsgBox "WINS Proxy not Enabled "
            End If
            If FixedInfo.EnableDns Then
                      MsgBox "NetBIOS Resolution Uses DNS "
            Else
                      MsgBox "NetBIOS Resolution Does not use DNS  "
            End If
    Else
            MsgBox "GetNetworkParams failed with error " & error
            Exit Sub
    End If
    
    'Enumerate all of the adapter specific information using the IP_ADAPTER_INFO structure.
    'Note:  IP_ADAPTER_INFO contains a linked list of adapter entries.
    
    AdapterInfoSize = 0
    error = GetAdaptersInfo(ByVal 0&, AdapterInfoSize)
    If error <> 0 Then
        If error <> ERROR_BUFFER_OVERFLOW Then
           MsgBox "GetAdaptersInfo sizing failed with error " & error
           Exit Sub
        End If
    End If
   ReDim AdapterInfoBuffer(AdapterInfoSize - 1)
 
 ' Get actual adapter information
   error = GetAdaptersInfo(AdapterInfoBuffer(0), AdapterInfoSize)
   If error <> 0 Then
      MsgBox "GetAdaptersInfo failed with error " & error
      Exit Sub
   End If
   CopyMemory AdapterInfo, AdapterInfoBuffer(0), Len(AdapterInfo)
   pAdapt = AdapterInfo.Next

   Do While pAdapt <> 0
        CopyMemory Buffer2, AdapterInfo, Len(Buffer2)
          Select Case Buffer2.Type
                Case MIB_IF_TYPE_ETHERNET
                    MsgBox "Ethernet adapter "
                Case MIB_IF_TYPE_TOKENRING
                    MsgBox "Token Ring adapter "
                Case MIB_IF_TYPE_FDDI
                    MsgBox "FDDI adapter "
                Case MIB_IF_TYPE_PPP
                    MsgBox "PPP adapter"
                Case MIB_IF_TYPE_LOOPBACK
                    MsgBox "Loopback adapter "
                Case MIB_IF_TYPE_SLIP
                    MsgBox "Slip adapter "
                Case Else
                    MsgBox "Other adapter "
        End Select
    MsgBox " AdapterName: " & Buffer2.AdapterName
    MsgBox "AdapterDescription: " & Buffer2.Description

    For i = 0 To Buffer2.AddressLength - 1
           PhysicalAddress = PhysicalAddress & Hex(Buffer2.Address(i))
            If i < Buffer2.AddressLength - 1 Then
             PhysicalAddress = PhysicalAddress & "-"
            End If

    Next
    MsgBox "Physical Address: " & PhysicalAddress
    If Buffer2.DhcpEnabled Then
            MsgBox "DHCP Enabled "
    Else
            MsgBox "DHCP disabled"
    End If

    pAddrStr = Buffer2.IpAddressList.Next
    Do While pAddrStr <> 0
           CopyMemory Buffer, Buffer2.IpAddressList, LenB(Buffer)
           MsgBox "IP Address: " & Buffer.IpAddress
           MsgBox "Subnet Mask: " & Buffer.IpMask
           pAddrStr = Buffer.Next
           If pAddrStr <> 0 Then
            CopyMemory Buffer2.IpAddressList, ByVal pAddrStr, Len(Buffer2.IpAddressList)
           End If
   Loop
    MsgBox "Default Gateway: " & Buffer2.GatewayList.IpAddress
    pAddrStr = Buffer2.GatewayList.Next
    Do While pAddrStr <> 0
            CopyMemory Buffer, Buffer2.GatewayList, Len(Buffer)
            MsgBox "IP Address: " & Buffer.IpAddress
            pAddrStr = Buffer.Next
            If pAddrStr <> 0 Then
            CopyMemory Buffer2.GatewayList, ByVal pAddrStr, Len(Buffer2.GatewayList)
            End If
    Loop

    MsgBox "DHCP Server: " & Buffer2.DhcpServer.IpAddress
    MsgBox "Primary WINS Server: " & Buffer2.PrimaryWinsServer.IpAddress
    MsgBox "Secondary WINS Server: " & Buffer2.SecondaryWinsServer.IpAddress

    ' Display time
    NewTime = CDate(Adapt.LeaseObtained)
    MsgBox "Lease Obtained: " & CStr(NewTime)

    NewTime = CDate(Adapt.LeaseExpires)
    MsgBox "Lease Expires :  " & CStr(NewTime)
    pAdapt = Buffer2.Next
    If pAdapt <> 0 Then
        CopyMemory AdapterInfo, ByVal pAdapt, Len(AdapterInfo)
    End If

   Loop
   
End Sub
