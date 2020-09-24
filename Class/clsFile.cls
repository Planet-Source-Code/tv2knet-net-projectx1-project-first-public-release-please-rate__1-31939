VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------
'-     By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'-----------------------------------
'- Credits: AllAPI-Guide
'-----------------------------------

Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type

Private Const FO_DELETE = &H3
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Dim SHDirOp As SHFILEOPSTRUCT

Public Sub TvRename(FromDest As String, ToDest As String)
    MoveFile FromDest, ToDest
End Sub
Public Sub TvCreateDir(Dir As String)
    CreateDirectory Dir, ByVal &H0
End Sub
Public Sub TvDeleteFile(File As String)
    DeleteFile File
End Sub
Public Sub TvDeleteDir(Dir As String)
    With SHDirOp
        .wFunc = FO_DELETE
        .pFrom = Dir
    End With
    SHFileOperation SHDirOp
End Sub

Public Sub TvCopyFile(FromDest As String, ToDest As String, Overwrite As Boolean)
    If Overwrite = True Then
    Overwrite = False
    GoTo 10
    End If
    If Overwrite = False Then
    Overwrite = True
    GoTo 10
    End If
10
    CopyFile FromDest, ToDest, Overwrite
    
End Sub
