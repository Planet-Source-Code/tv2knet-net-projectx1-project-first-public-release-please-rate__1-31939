Attribute VB_Name = "modCodeX5"

'In general section
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type

'Declare the API-Functions
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Function IsOnPic(PControl As PictureBox) As Boolean
'-----------------------------------
'-    IsOnPic()
'-
'- It'll return true if mouse going over picbox
'-
'-     By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'-----------------------------------

    Dim Rec As RECT, Point As POINTAPI
    GetWindowRect PControl.hwnd, Rec
    GetCursorPos Point
    
If Point.X < Rec.Left Or Point.X > Rec.Right Or Point.Y < Rec.Top Or Point.Y > Rec.Bottom Then
    IsOnPic = False
    Exit Function
End If

    IsOnPic = True
End Function




Public Function IsInIDE() As Boolean
'-----------------------------------
'-    IsInIDE()
'-
'- It'll return true if running in IDE
'-
'-     By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'-----------------------------------
'-TESTED: In VB6.0 SP5
'-----------------------------------
Dim X As Long
On Error Resume Next
X = VB.App.LogMode()
If X = 1 Then IsInIDE = False Else IsInIDE = True
End Function

