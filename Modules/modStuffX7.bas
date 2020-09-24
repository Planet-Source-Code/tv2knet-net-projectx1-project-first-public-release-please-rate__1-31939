Attribute VB_Name = "modStuffX7"
Const AC_SRC_OVER = &H0
Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
Public Sub BlendPic(Picture1 As PictureBox, Picture2 As PictureBox)
'-----------------------------------
'-    BlendPic(Picture1, Picture2)
'-
'-  Picture1 and Picture2 are Pictures
'-
'-  It will Blen Picture1 with Picture2
'-     and show it in Picture2
'-
'-   By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'- Requires: >= Windows 98
'-----------------------------------
'- Credits to: AllApi Network
'-----------------------------------
    Dim BF As BLENDFUNCTION, lBF As Long
    'Set the graphics mode to persistent
    Picture1.AutoRedraw = True
    Picture2.AutoRedraw = True
    'API uses pixels
    Picture1.ScaleMode = vbPixels
    Picture2.ScaleMode = vbPixels
    'set the parameters
    With BF
        .BlendOp = AC_SRC_OVER
        .BlendFlags = 0
        .SourceConstantAlpha = 128
        .AlphaFormat = 0
    End With
    'copy the BLENDFUNCTION-structure to a Long
    RtlMoveMemory lBF, BF, 4
    'AlphaBlend the picture from Picture1 over the picture of Picture2
    AlphaBlend Picture2.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, lBF
End Sub

