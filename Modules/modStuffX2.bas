Attribute VB_Name = "modStuffX2"
Option Explicit

Public Function TvParseIt(HTMLData As String) As String
'-----------------------------------
'-    TvParseIt(HTMLData)
'-   HTMLData is the string with the HTML code in
'-
'-  It'll return the extracted text
'-
'-   By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'-----------------------------------

Dim X As String
Dim p As Long
Dim i As Long
Dim s As String
Dim d As String
Dim q As Long
X = HTMLData
X = X + ">}" 'To Be Sure
p = Len(X)


For i = 1 To p
s = Mid$(X, i, 1)
If s = "<" Then
For q = i To p
s = Mid$(X, q, 1)
If s = ">" Then
i = i + q - i
GoTo 10
End If

Next
End If
If s = "{" Then
For q = i To p
s = Mid$(X, q, 1)
If s = "}" Then
i = i + q - i
GoTo 10
End If

Next
End If
d = d + s
10
Next
d = Replace(d, "<", "", 1, Len(d), vbTextCompare)
d = Replace(d, ">", "", 1, Len(d), vbTextCompare)
d = Replace(d, "{", "", 1, Len(d), vbTextCompare)
d = Replace(d, "}", "", 1, Len(d), vbTextCompare)
TvParseIt = d
End Function

Public Sub SaveHTML2TXT(HTMLText As String, TXTLocation As String)
'-----------------------------------
'-    SaveHTML2TXT(HTMLText,TXTLocation)
'-   HTMLText is the string with the HTML code in
'- TXTLocation is a string pointing to the location where to save to
'-
'-  It'll save the extracted text to a file
'- It needs the TvParseIt Function to work
'-
'-   By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'-----------------------------------
On Error GoTo 10
Open TXTLocation For Output As #1
 Print #1, TvParseIt(HTMLText)
Close #1
Exit Sub
10
On Error Resume Next
Close #1
End Sub


Public Sub SaveHTMLFile2TXTFile(HTMLLocation As String, TXTLocation As String)
'-----------------------------------
'-    SaveHTML2TXT(HTMLText,TXTLocation)
'-   HTMLText is the location of the HTML file(a string)
'- TXTLocation is a string pointing to the location where to save to
'-
'-  It'll save the extracted text to a file
'- It needs the TvParseIt Function to work
'-
'-   By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'-----------------------------------
Dim X As String
Dim d As String
Dim z As String
Open HTMLLocation For Input As #1
While EOF(1) = False
Line Input #1, z
If d = "" Then
d = z
End If
Debug.Print z
d = d + vbCrLf + z
'Debug.Print d
'
Wend


Close #1

Open TXTLocation For Output As #1
 Print #1, TvParseIt(d)
Close #1
Exit Sub

Close #1
End Sub

