Attribute VB_Name = "modStuffX8"
'-----------------------------------
'-
'-     These function are some nice
'-      internet/network functions
'-
'-
'-   By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'-----------------------------------

Private Declare Function DoFileDownload Lib "shdocvw.dll" (ByVal lpszFile As String) As Long
'--
Private Declare Function InetIsOffline Lib "url.dll" (ByVal dwFlags As Long) As Long
'--
Const NETWORK_ALIVE_AOL = &H4
Const NETWORK_ALIVE_LAN = &H1
Const NETWORK_ALIVE_WAN = &H2
Private Declare Function IsNetworkAlive Lib "SENSAPI.DLL" (ByRef lpdwFlags As Long) As Long
'-
Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) As Long
Dim sConnType As String * 255
'--
Public Function TvGetConnectState() As String
    Dim Ret As Long
    Ret = InternetGetConnectedStateEx(Ret, sConnType, 254, 0)
    If Ret = 1 Then
      TvGetConnectState = "1- You are connected to Internet via a " & sConnType
    Else
      TvGetConnectState = "0- You are not connected to internet"
    End If
End Function

Public Function IsOnNetwork() As String
    Dim Ret As Long
    If IsNetworkAlive(Ret) = 0 Then
IsOnNetwork = "0- The local system is not connected to a network!"
    Else
IsOnNetwork = "1- The local system is connected to a " + IIf(Ret = NETWORK_ALIVE_AOL, "AOL", IIf(Ret = NETWORK_ALIVE_LAN, "LAN", "WAN")) + " network!"
    End If
End Function
Public Function IsOnline() As Boolean
If InetIsOffline(0) = 0 Then IsOnline = True Else IsOnline = False

End Function

Public Sub OpenDownloadFile(FileURL As String)
'-----------------------------------
'-   OpenDownloadFile(FileURL)
'-
'-  FileURL is a file located on the internet
'-
'- It will launch the Download File dialog
'-     and it will download the file
'-
'-   By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'- Required= Internet Explorer (4+?)
'-
'-----------------------------------
'- Credits to: AllApi Network
'-----------------------------------
   DoFileDownload StrConv(FileURL, vbUnicode)
End Sub


