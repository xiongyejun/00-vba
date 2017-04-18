

	Function SpecialFolderPath() As String
	    Dim strSpecialFolderPath As String
	
		' home, documents, desktop, music, pictures, movies, applications
	    If Int(Val(Application.Version)) > 14 Then
	        SpecialFolderPath = _
	        MacScript("return POSIX path of (path to desktop folder) as string")
	        'Replace line needed for the special folders Home and documents
	        SpecialFolder = _
	        Replace(SpecialFolder, "/Library/Containers/com.microsoft.Excel/Data", "")
	    Else
	        SpecialFolderPath = MacScript("return (path to desktop folder) as string")
	    End If
	End Function