	Function FileOrFolderExistsOnMac(FileOrFolderstr As String) As Boolean
	'Ron de Bruin : 26-June-2015
	'Function to test whether a file or folder exist on a Mac in office 2011 and up
	'Uses AppleScript to avoid the problem with long names in Office 2011,
	'limit is max 32 characters including the extension in 2011.
	    Dim ScriptToCheckFileFolder As String
	    Dim TestStr As String
	
	    If Val(Application.Version) < 15 Then
	        ScriptToCheckFileFolder = "tell application " & Chr(34) & "System Events" & Chr(34) & _
	         "to return exists disk item (" & Chr(34) & FileOrFolderstr & Chr(34) & " as string)"
	        FileOrFolderExistsOnMac = MacScript(ScriptToCheckFileFolder)
	    Else
	        On Error Resume Next
	        TestStr = Dir(FileOrFolderstr, vbDirectory)
	        On Error GoTo 0
	        If Not TestStr = vbNullString Then FileOrFolderExistsOnMac = True
	    End If
	End Function
