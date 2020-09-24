Attribute VB_Name = "modStuffX1"
Option Explicit


Public Function GetVBCRLF(Data As String) As Long
'-----------------------------------
'-    GetVBCRLF2(Data)
'-   Data is a string
'-
'- It'll return the amount of
'-     VBCrLF's in the string
'-
'-   By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'-----------------------------------


GetVBCRLF2 = Len(Data) - Len(Replace(Data, Chr$(13) + Chr$(10), " ", 1, Len(Data), vbTextCompare))
End Function


Public Function GetString(Data As String, Find As String) As Long
'-----------------------------------
'-    GetString(Data, Find)
'-   Data is a string
'-   Find is a string
'-
'- It'll return the amount of
'-     'Find' in the string
'-
'-   By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'- Credits To: Andrew Murphy
'-
'-----------------------------------
Dim t As String
Dim X As String
Dim Y As String
Dim z As Long
Dim i As Integer
On Error GoTo 10
X = Data
For i = 1 To Len(X)
t = Mid$(X, i, Len(Find))
If t = Find Then
z = z + 1
i = i + Len(Find) - 1
End If
10
On Error Resume Next
Next

GetString = z
End Function

'Public Function GetWords(Data As String) As Long
'-----------------------------------
'-     GetWords(Data)
'-   Data is a string
'-
'- It'll return the amount of
'-     Words in the string
'-
'-   By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'-----------------------------------
'
' TO DO:
'
' Detection of VBCRLF's and " "
'
'-----------------------------------
'Dim t As String
'Dim x As String
'Dim o As String
'Dim z As Integer
'Dim i As Integer
'Dim p As Integer
'Dim j As Long
'Dim SpaceC As Integer
'Dim LineC As Integer
'On Error GoTo 10
'x = Data
'If x = "" Then GoTo 123
'If Replace(x, " ", "", 1, Len(x)) = "" Then GoTo 123
'x = x + "x"
'x = Replace(x, vbCrLf, " ", 1, Len(x))
'''x = Replace(x, "  ", " ", 1, Len(x))
'
'For i = 1 To Len(x)
't = Mid$(x, i, 2)
'If t = vbCrLf Then
'i = i + 1
'p = p + 1
'End If
'

''t = Mid$(x, i, 1)
'If t = " " And SpaceC = 0 Then
'SpaceC = 1
'z = z + 1
'GoTo 20
'End If
'

'
'20
'SpaceC = 0
'10
'
'DoEvents
'Next
'
'If Replace(x, " ", "") = "" Then
'p = p - 1
'End If
'If Right(x, 2) = " x" Then
'p = p - 1
'End If
'If Left(x, 1) = " " Then
'p = p - 1
'End If
'GetWords = z + p + 1
'Exit Function
'123
'GetWords = 0
'End Function




Public Function TvReplaceStr(Data As String, Find As String, Replace As String, Start As Integer) As String
'-----------------------------------
'-     TvReplaceStr(Data, Find, Replace, Start)
'-   Data is a string
'-   Find is a string
'-   Replace is a string
'-   Start is an integer
'-
'- It'll return the string(text)
'- with the find string replaced
'- with the replace string
'-
'-   By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'-----------------------------------
Dim t As String
Dim X As String
Dim Y As String
Dim z As Long
Dim i As Long
Dim p As Long
Dim d As Long
Dim O As String
p = Len(Find)
d = Len(Replace)
If p = 0 Or d = 0 Or Len(Data) = 0 Then Exit Function
10
For i = Start To Len(Data)
X = Mid$(Data, i, p)
If X = Find Then
O = Replace
i = i + p - 1
Else
O = Mid$(Data, i, 1)
End If
Y = Y + O
Next
TvReplaceStr = Y
End Function




Public Function GetWordsBeta(Data As String) As Long 'Won't work perfect.... Yet!
Dim X As String
Dim z As String
Dim p As Long
Dim m As Integer
Dim q As Integer
X = Data
z = ""
p = 0
m = 0
If Data <> "" Then
m = 1
End If
z = Replace(X, " ", Chr(255), 1, Len(Data), vbTextCompare)
If z <> Replace(z, vbCrLf, Chr(255), 1, Len(z), vbTextCompare) Then
z = Replace(z, vbCrLf, Chr(255), 1, Len(z), vbTextCompare)
p = 1
End If

Dim i As Integer
For i = 1 To 300
z = Replace(z, Chr(255) + Chr(255), Chr(255), 1, Len(z), vbTextCompare)
Next
If Right(Data, 1) = " " Then q = 1
GetWordsBeta = GetString(z, Chr(255)) + m - q

End Function



Public Function TvReverse(StringT As String) As String
'-----------------------------------
'-     TvReverse(StringT as string)
'-   StringT is the original text string
'-     that should be reversed
'-
'- It'll return the string(text)
'-   with the reversed text
'-
'-   By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'-----------------------------------
If StringT = "" Then Exit Function
Dim ReversedString As String
Dim CurrentChar As String
Dim i As Long
ReversedString = ""
CurrentChar = ""
For i = 1 To Len(StringT)
CurrentChar = Mid$(StringT, Len(StringT) - i + 1, 1)
ReversedString = ReversedString + CurrentChar
Next
TvReverse = ReversedString
End Function


Public Function GetChars(Data As String) As Long
'-----------------------------------
'-     GetChars(Data as string)
'-     Data is a string
'-
'- It'll return the number of chars
'-      in the string
'-
'-   By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'-----------------------------------
GetChars = Len(Data)


End Function


Public Function TvRandomString(Size As Long) As String
'-----------------------------------
'-     TvRandomString(Size as long)
'-   Size is the lenght of the new string
'-
'- It'll return the string(text)
'-   that is generated
'-
'-   By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'-----------------------------------
Dim i As Long
Dim X As String
Dim Y As String
Dim f As Long
Randomize Timer
For i = 1 To Size
10
Y = Chr(1 + Rnd(Timer) * 124) ' No special chars and empty spaces
For f = 48 To 57 ' To remove numbers
If Y = Chr(f) Then
f = 56
GoTo 10
End If
Next
X = X + Y
Next
TvRandomString = X
End Function


Public Function TvIsInStr(Data As String, Find As String) As Long
'-----------------------------------
'-    TvIsInStr(Data, Find)
'-   Data is a string
'-   Find is a string
'-
'- It'll return the amount of
'-     'Find' in the string
'-
'-   By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'-----------------------------------
Dim TempB As String
TempB = String(Len(Find) - 1, Chr(0))
TvIsInStr = Len(Data) - Len(Replace(Data, Find, TempB, 1, Len(Data), vbBinaryCompare))
End Function

