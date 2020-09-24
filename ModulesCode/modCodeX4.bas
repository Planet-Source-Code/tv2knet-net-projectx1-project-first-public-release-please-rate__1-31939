Attribute VB_Name = "modCodeX4"
Declare Sub SHAddToRecentDocs Lib "shell32.dll" (ByVal uFlags As Long, ByVal pv As String)
Public Sub TvAddRecent(File As String)
    Call SHAddToRecentDocs(2, File)
End Sub


Function IsIDE() As Boolean
On Error GoTo 10
Debug.Print , 1, 1, X1
IsIDE = True
Exit Function
10
IsIDE = False

End Function
