Attribute VB_Name = "modCodeX9"

Option Explicit








Public Function TvExtractT(ExtractType As Integer, InputText As String) As String
'-----------------------------------
'-   TvExtractT(ExtractType, InputText)
'-
'- ExtractType is 1/2 (1=filelocation,2=pathlocation)
'-   InputText is the file path
'-
'- It'll return the path or file
'-
'-   By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'-----------------------------------
If ExtractType <> 1 And ExtractType <> 2 Then Exit Function
If Len(InputText) = Len(Replace(InputText, "\", "", 1, Len(InputText), vbTextCompare)) Then
TvExtractT = InputText
Exit Function
End If

'--
Dim X As String
Dim p As Long
'--
'--
p = ExtractType
X = InputText
'--
'--
Dim z As String
Dim i As Long
Dim t As String
'--
If p = 1 Then
For i = Len(X) To 1 Step -1
z = Mid$(X, i, 1)
If z = "\" Then
t = Right$(X, Len(X) - i)
GoTo 10
End If
Next
End If

If p = 2 Then
For i = Len(X) To 1 Step -1
z = Mid$(X, i, 1)
If z = "\" Then
t = Left$(X, i)
GoTo 10
End If
Next
End If

10
TvExtractT = t
End Function
























