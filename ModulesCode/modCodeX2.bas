Attribute VB_Name = "modCodeX2"
'-----------------------------------
'-
'- Get/Set computer name and get Username
'-
'-     By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'-----------------------------------
Const MAX_COMPUTERNAME_LENGTH As Long = 31
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function SetComputerName Lib "kernel32" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Public Function TvGetUserName() As String
Dim X As String
X = String(100, Chr$(0)) ' Create Buffer
GetUserName X, 100
TvGetUserName = X
End Function
Public Sub TvSetComputerName(Name As String)
    Dim sNewName As String
    sNewName = Name
    
    If sNewName = "" Then
    sNewName = InputBox("Please enter a new computer name.", "Please input new computername!", "")

    End If

    'Ask for a new computer name
    
    If sNewName = "" Then Exit Sub
    'Set the new computer name
    SetComputerName sNewName
End Sub


Public Function TvGetComputerName() As String
    Dim dwLen As Long
    Dim strString As String
    'Create a buffer
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwLen, "X")
    'Get the computer name
    GetComputerName strString, dwLen
    'get only the actual data
    strString = Left(strString, dwLen)
    TvGetComputerName = strString
End Function
