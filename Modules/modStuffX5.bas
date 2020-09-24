Attribute VB_Name = "modStuffX5"
Option Explicit
Public h(1 To 10) As Long

Public Sub SaveForm(Height As Long, FormNumber As Long)
'-----------------------------------
'-   Public Sub SaveForm(Height, FormNumber)
'-
'-        Height is the height of the form
'-  FormNumber is just a number to indicate the form
'-
'-   By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'-
'-----------------------------------
h(FormNumber) = Height
End Function
Public Sub RollInForm(FormD As Form, FormNumber As Long, Freeze As Boolean)
'-----------------------------------
'-   Public Sub RollInForm(FormD, FormNumber, Freeze)
'-
'-     This will 'Roll' a form in
'-
'-       FormD is a form, FormNumber is the number
'-     you used to indicate the form and Freeze is a boolean
'-      that indicates to don't use DoEvents or not
'-
'-   By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'-
'-----------------------------------
FormD.Hide
Dim i As Long
Dim x As Long
x = h(FormNumber)

FormD.Show
For i = x To 1 Step -1

FormD.Height = i
i = i - 9
If Freeze <> True Then
DoEvents
End If
Next
End Sub
Public Sub RollOutForm(FormD As Form, FormNumber As Long, Freeze As Boolean)
'-----------------------------------
'-   Public Sub RollOutForm(FormD, FormNumber, Freeze)
'-
'-     This will 'Roll' a form out
'-
'-       FormD is a form, FormNumber is the number
'-     you used to indicate the form and Freeze is a boolean
'-      that indicates to don't use DoEvents or not
'-
'-   By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'-
'-----------------------------------
FormD.Hide
Dim i As Long
Dim x As Long
x = h(FormNumber)

FormD.Show
For i = 1 To x Step 1


FormD.Height = i
i = i + 9
If Freeze <> True Then
DoEvents
End If
Next

End Sub






'-----------------




Public Sub FadeIN(FormS As Form, Start As Long, EndN As Long)

With FormS
For i = Start To EndN
.BackColor = RGB(i, i, i)
DoEvents
Next
End With
End Sub

Public Sub FadeOUT(FormS As Form, Start As Long, EndN As Long)

With FormS
For i = EndN To Start Step -1
.BackColor = RGB(i, i, i)
DoEvents
Next
End With
End Sub

