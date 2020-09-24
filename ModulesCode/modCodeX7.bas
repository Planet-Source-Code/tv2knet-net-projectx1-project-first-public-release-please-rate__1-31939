Attribute VB_Name = "modCodeX7"
'-----------------------------------
'-    Encode/Decode
'-
'- It will encode/decode a text
'-  Warning: Don't put encoded text in textbox
'-     The text will be corrupted
'-   Put it in a string like Dim EncT as string
'-
'-     By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'-----------------------------------
Public Function TvEncode(Data As String, Password As String) As String
Dim X As String
Dim Y As String
Dim t As String
Dim p As Long
Dim i As Long
Dim q As Long
If CheckTextOnly(Password) = False Then Exit Function
q = 1
For i = 1 To Len(Data)
X = Mid$(Data, i, 1)
Y = ChrW(AscW(X) + AscW(Mid$(Password, q, 1)))
t = t + Y
q = q + 1
If q - 1 = Len(Password) Then
q = 1
End If
Next
TvEncode = t
End Function

Public Function TvDecode(Data As String, Password As String) As String
Dim X As String
Dim Y As String
Dim t As String
Dim p As Long
Dim i As Long
Dim q As Long
If CheckTextOnly(Password) = False Then Exit Function
q = 1
For i = 1 To Len(Data)
X = Mid$(Data, i, 1)
Y = ChrW(AscW(X) - AscW(Mid$(Password, q, 1)))
t = t + Y
q = q + 1
If q - 1 = Len(Password) Then
q = 1
End If

Next
TvDecode = t
End Function


Function CheckTextOnly(Data As String) As Boolean
Dim X As String
Dim i As Long
CheckTextOnly = True
Dim t(0 To 9) As String

t(0) = Replace(Data, "0", "", 1, Len(Data), vbBinaryCompare)
t(1) = Replace(Data, "1", "", 1, Len(Data), vbBinaryCompare)
t(2) = Replace(Data, "2", "", 1, Len(Data), vbBinaryCompare)
t(3) = Replace(Data, "3", "", 1, Len(Data), vbBinaryCompare)
t(4) = Replace(Data, "4", "", 1, Len(Data), vbBinaryCompare)
t(5) = Replace(Data, "5", "", 1, Len(Data), vbBinaryCompare)
t(6) = Replace(Data, "6", "", 1, Len(Data), vbBinaryCompare)
t(7) = Replace(Data, "7", "", 1, Len(Data), vbBinaryCompare)
t(8) = Replace(Data, "8", "", 1, Len(Data), vbBinaryCompare)
t(9) = Replace(Data, "9", "", 1, Len(Data), vbBinaryCompare)
If Len(t(0)) <> Len(Data) Or Len(t(1)) <> Len(Data) Or Len(t(2)) <> Len(Data) Or Len(t(3)) <> Len(Data) Or Len(t(4)) <> Len(Data) Or Len(t(5)) <> Len(Data) Or Len(t(6)) <> Len(Data) Or Len(t(7)) <> Len(Data) Or Len(t(8)) <> Len(Data) Or Len(t(9)) <> Len(Data) Then
CheckTextOnly = False
End If
End Function
