Attribute VB_Name = "modCodeX6"
Private Declare Function CreateCaret Lib "user32" (ByVal hwnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function ShowCaret Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long

Public Sub FocusOnTextChar(TextB As TextBox)
'-----------------------------------
'-    FocusOnTextChar()
'-
'- It'll set a custom 'pointer'
'-
'-     By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'-----------------------------------
    Dim h As Long
    'sets/retrieves the window which has the focus
    TextB.SetFocus
    h = GetFocus()
    'Create a new cursor
    'Call CreateCaret(H&, 0, 5, 15) ' Use this for constant size
    Call CreateCaret(h&, 0, 5, TextB.FontSize * 2) ' Isn't 100% good but it works
    'Show the new cursor
    X& = ShowCaret&(h&)
End Sub
