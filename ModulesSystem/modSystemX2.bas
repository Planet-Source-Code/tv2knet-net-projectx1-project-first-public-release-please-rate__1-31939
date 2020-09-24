Attribute VB_Name = "modSystemX2"
'-----------------------------------
'-   You can change/restrict Windows a little:
'-
'-  Disable/Enable Ctrl+Alt+Delete
'-  Hide/Show Stuff
'-  Set forms on top/not on top
'-
'-   By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'-----------------------------------

Option Explicit
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
'--
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const FLAGS = SWP_NOSIZE Or SWP_NOMOVE
'--
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long
    Const SPI_SCREENSAVERRUNNING = 97
    Const RSP_SIMPLE_SERVICE = 1
    Const RSP_UNREGISTER_SERVICE = 0
'--
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long



'--
Public Sub HideStart(FormT As Form)
' Put HideStart in a timer
   Dim desktop, tophwnd As Long
   
   desktop = GetDesktopWindow()
   tophwnd = GetTopWindow(desktop)
   If tophwnd <> FormT.hwnd Then SetForegroundWindow (FormT.hwnd)

End Sub

Public Sub TvHideFromCAD()
    Dim lngProcessID As Long
    Dim lngid As Long
    
    lngid = GetCurrentProcessId()
    Call RegisterServiceProcess(lngid, RSP_SIMPLE_SERVICE)
End Sub


Public Sub TvShowInCAD()
    Dim lngProcessID As Long
    Dim lngid As Long
    
    lngid = GetCurrentProcessId()
    Call RegisterServiceProcess(lngid, RSP_UNREGISTER_SERVICE)
End Sub

Public Function TvDisableCAD()
Dim bck As Integer
 Dim Old As Boolean
 bck = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, Old, 0)
End Function
Public Function TvEnableCAD()
Dim bck As Integer
Dim Old As Boolean
 bck = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, Old, 0)
End Function

Public Function TvOnTop(TForm As Form)
Dim WinOnTop As Long
WinOnTop = SetWindowPos(TForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Function
Public Function TvNotOnTop(TForm As Form)
Dim WinOnTop As Long
WinOnTop = SetWindowPos(TForm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Function


