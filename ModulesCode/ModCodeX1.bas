Attribute VB_Name = "ModCodeX1"
Private Declare Sub FatalAppExit Lib "kernel32" Alias "FatalAppExitA" (ByVal uAction As Long, ByVal lpMessageText As String)
Private Declare Sub FatalExit Lib "kernel32" (ByVal code As Long)

Public Sub FatalMessage(Message As String)
'-----------------------------------
'-    FatalMessage(Message)
'-
'-   Message is the string containing the
'-      message that should be viewed
'-
'-   Will 'crash' your program with
'-    the message that you give
'-
'-   By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'-----------------------------------
'-WARNING: Will Also Close IDE
'-----------------------------------

    FatalAppExit 0, Message
End Sub

Public Sub FatalClose()
'-----------------------------------
'-    FatalClose()
'-
'-   Will 'crash' your program
'-
'-   By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'-----------------------------------
'-WARNING: Will Also Close IDE
'-----------------------------------
    FatalExit 1
End Sub
