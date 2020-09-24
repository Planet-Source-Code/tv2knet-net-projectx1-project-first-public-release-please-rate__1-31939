Attribute VB_Name = "modStuffX6"
Option Explicit
Public Function GetLast(Data As String) As Long
'-----------------------------------
'-   Public Sub GetLast(Data)
'-
'-    Data is the number containing a '.'
'-     with a number after it
'-      (Only 1 after supported)
'-
'-   Will return the Rounded number
'-
'-   By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'-
'-----------------------------------

Dim i As Long
Dim Y As String
Dim z As String
Dim t As String
Dim PP As Integer
Dim p(1 To 2) As Long
GetLast = 0
On Error GoTo 10

If Replace(Data, ".", "") = Data Then
Data = Data + ".0"
End If


For i = 1 To Len(Data)
Y = Mid$(Data, i, 1)
If PP = 1 Then
p(1) = Val(t)
Y = Mid$(Data, i, 1)
p(2) = Val(Y)
If p(2) < 5 Then
p(1) = p(1)
Else
p(1) = p(1) + 1
End If
GoTo 10
End If
If Y <> "." Then
t = t + Y
Else
PP = 1
End If
Next
10
GetLast = p(1)
End Function



