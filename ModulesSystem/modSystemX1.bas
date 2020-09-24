Attribute VB_Name = "modSystemX1"
'-----------------------------------
'-
'-  WARNING: These module has some DANGEROUS
'-     functions in it:
'-
'-  TvBlockInput (Can't press a button/Move mouse)
'-  QWindows (Quit Windows)
'-
'-   By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'-----------------------------------

Const EWX_FORCE = 4
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function fCreateShellLink Lib "VB5STKIT.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long
Private Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Sub TvBlockInput(Always As Boolean, Optional TimeToSleep As Long, Optional IsInSeconds As Boolean)

    DoEvents
    If Always = True Then
    'block the mouse and keyboard input
      BlockInput True
      Exit Sub
    End If
    If TimeToSleep = 0 Then Exit Sub
          BlockInput True
    'block the mouse and keyboard input
    BlockInput True
    If IsInSeconds = True Then TimeToSleep = TimeToSleep * 1000
Debug.Print TimeToSleep
    Sleep TimeToSleep
    'unblock the mouse and keyboard input
    BlockInput False
End Sub
Public Sub TvEnableInput()
    'unblock the mouse and keyboard input
    BlockInput False
End Sub

Public Sub TvCreateLink(PutLinkIn As String, LinkName As String, LinksTo As String)
lng = fCreateShellLink(PutLinkIn, LinkName, LinksTo, "")
End Sub

Public Sub QWindows(DoWhat As Integer, ShowWarning As Boolean)
Dim Message As String

If DoWhat = 0 Then Message = "LogOff you from Windows"
If DoWhat = 1 Then Message = "Shut Down the computer"
If DoWhat = 2 Then Message = "Reboot the computer"
If DoWhat > 2 Then Exit Sub
If DoWhat < 0 Then Exit Sub
If ShowWarning = True Then
    msg = MsgBox("This program is going to " + Message + ". Press OK to continue or Cancel to stop.", vbCritical + vbOKCancel + 256, "WARNING!")
    If msg = vbCancel Then Exit Sub
End If
    Ret& = ExitWindowsEx(EWX_FORCE Or DoWhat, 0)
End Sub

