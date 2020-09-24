Attribute VB_Name = "modCodeX8"

Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
Public Function TvGetFreeDisk(DiskD As String) As Long
'-----------------------------------
'-    TvGetFreeDisk()
'-
'- It'll return the free bytes on a drive
'-
'-     By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'-----------------------------------
    Dim r As Long
    Dim RPN As String
    Dim TFB As Currency
    RPN = DiskD
    GetDiskFreeSpaceEx RPN, 0, 0, TFB

TvGetFreeDisk = TFB * 10000 ' bytes

End Function
