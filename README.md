<div align="center">

## IPCONFIG \(VB\)


</div>

### Description

Returns all Information about all Network devices in the computer (even Virtual ones like the PPPoE Adapter for DSL connections). Displays IP address, dns servers, wins servers, host name, etc...
 
### More Info
 
This code works under Windows 3.11, Windows 95, Windows 98, Windows NT, Windows 2000. Using Visual Basic 6.0 (no service packs).

As this code returns everything to message boxes, it will be easy for you to modify it to use in your application.


<span>             |<span>
---                |---
**Submitted On**   |2000-10-28 19:07:20
**By**             |[MicroVB INC](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/microvb-inc.md)
**Level**          |Advanced
**User Rating**    |4.3 (26 globes from 6 users)
**Compatibility**  |VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD1103310282000\.zip](https://github.com/Planet-Source-Code/microvb-inc-ipconfig-vb__1-12361/archive/master.zip)

### API Declarations

```
Public Declare Function GetNetworkParams Lib "IPHlpApi" (FixedInfo As Any, pOutBufLen As Long) As Long
Public Declare Function GetAdaptersInfo Lib "IPHlpApi" (IpAdapterInfo As Any, pOutBufLen As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
```





