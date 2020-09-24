Attribute VB_Name = "modStuffX4"
Option Explicit
Public Declare Function sndPlaySound Lib "mmsystem.dll" (ByVal lpszSoundName As String, ByVal uFlags As Integer)
Public Function PlayWave(Wave As String) As Long
'-----------------------------------
'-    PlayWave(Wave)
'-  Wave is the location of the WAV
'-   file that you wanne play
'-
'- Returns the sndPlaySound return
'-
'-   By T-Virus Creations
'- http://www.tvirusonline.be
'- email: tvirus4ever@yahoo.co.uk
'-
'- API: -sndPlaySound(mmsystem.dll)
'-
'-----------------------------------
PlayWave = sndPlaySound(Wave, 3)
End Function
