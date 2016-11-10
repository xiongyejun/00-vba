[http://www.rondebruin.nl/mac/mac015.htm](http://www.rondebruin.nl/mac/mac015.htm)

# Select files on a Mac (GetOpenFilename) #
In Windows we can use for example GetOpenFilename to select files and do what we want with the path results, you can use filefilter to only display the files you want and use MultiSelect to select more then one file. Also it is possible with ChDrive and ChDir to set the folder that is selected when GetOpenFilename opens, see a example on the bottom of this page for Excel for Windows.

But on a Mac the filefilter is not working and it is not possible to select more then one file. Also ChDir is not working like in Windows to set the folder that will open with GetOpenFilename. But we can use a combination of VBA and Applescript, see example below that only let you select xlsx files and you can set the start folder.

**Important :** The file location can be very important in Mac Excel 2016, read : Problems with Apple’s sandbox requirements and Mac Office 2016 with VBA code

 

## Example for Mac Excel 2011 and 2016 ##

You can run the macro below without changing it, it opens the Desktop in the file select dialog and you can only select one xlsx file now, see the code how to change this.

**Note :** If you got problems with the code please report it to me so i can fix it.

	Sub Select_File_Or_Files_Mac()
	    'Select files in Mac Excel with the format that you want
	    'Working in Mac Excel 2011 and 2016
	    'Ron de Bruin, 20 March 2016
	    Dim MyPath As String
	    Dim MyScript As String
	    Dim MyFiles As String
	    Dim MySplit As Variant
	    Dim N As Long
	    Dim Fname As String
	    Dim mybook As Workbook
	    Dim OneFile As Boolean
	    Dim FileFormat As String
	
	    'In this example you can only select xlsx files
	    'See my webpage how to use other and more formats.
	    FileFormat = "{""org.openxmlformats.spreadsheetml.sheet""}"
	
	    ' Set to True if you only want to be able to select one file
	    ' And to False to be able to select one or more files
	    OneFile = True
	
	    On Error Resume Next
	    MyPath = MacScript("return (path to desktop folder) as String")
	    'Or use A full path with as separator the :
	    'MyPath = "HarddriveName:Users:<UserName>:Desktop:YourFolder:"
	
	    'Building the applescript string, do not change this
	    If Val(Application.Version) < 15 Then
	        'This is Mac Excel 2011
	        If OneFile = True Then
	            MyScript = _
	                "set theFile to (choose file of type" & _
	                " " & FileFormat & " " & _
	                "with prompt ""Please select a file"" default location alias """ & _
	                MyPath & """ without multiple selections allowed) as string" & vbNewLine & _
	                "return theFile"
	        Else
	            MyScript = _
	                "set applescript's text item delimiters to {ASCII character 10} " & vbNewLine & _
	                "set theFiles to (choose file of type" & _
	                " " & FileFormat & " " & _
	                "with prompt ""Please select a file or files"" default location alias """ & _
	                MyPath & """ with multiple selections allowed) as string" & vbNewLine & _
	                "set applescript's text item delimiters to """" " & vbNewLine & _
	                "return theFiles"
	        End If
	    Else
	        'This is Mac Excel 2016
	        If OneFile = True Then
	            MyScript = _
	                "set theFile to (choose file of type" & _
	                " " & FileFormat & " " & _
	                "with prompt ""Please select a file"" default location alias """ & _
	                MyPath & """ without multiple selections allowed) as string" & vbNewLine & _
	                "return posix path of theFile"
	        Else
	            MyScript = _
	                "set theFiles to (choose file of type" & _
	                " " & FileFormat & " " & _
	                "with prompt ""Please select a file or files"" default location alias """ & _
	                MyPath & """ with multiple selections allowed)" & vbNewLine & _
	                "set thePOSIXFiles to {}" & vbNewLine & _
	                "repeat with aFile in theFiles" & vbNewLine & _
	                "set end of thePOSIXFiles to POSIX path of aFile" & vbNewLine & _
	                "end repeat" & vbNewLine & _
	                "set {TID, text item delimiters} to {text item delimiters, ASCII character 10}" & vbNewLine & _
	                "set thePOSIXFiles to thePOSIXFiles as text" & vbNewLine & _
	                "set text item delimiters to TID" & vbNewLine & _
	                "return thePOSIXFiles"
	        End If
	    End If
	
	    MyFiles = MacScript(MyScript)
	    On Error GoTo 0
	
	    'If you select one or more files MyFiles is not empty
	    'We can do things with the file paths now like I show you below
	    If MyFiles <> "" Then
	        With Application
	            .ScreenUpdating = False
	            .EnableEvents = False
	        End With
	
	        MySplit = Split(MyFiles, Chr(10))
	        For N = LBound(MySplit) To UBound(MySplit)
	
	            'Get file name only and test if it is open
	            Fname = Right(MySplit(N), Len(MySplit(N)) - InStrRev(MySplit(N), _
	                Application.PathSeparator, , 1))
	
	            If bIsBookOpen(Fname) = False Then
	
	                Set mybook = Nothing
	                On Error Resume Next
	                Set mybook = Workbooks.Open(MySplit(N))
	                On Error GoTo 0
	
	                If Not mybook Is Nothing Then
	                    MsgBox "You open this file : " & MySplit(N) & vbNewLine & _
	                    "And after you press OK it will be closed" & vbNewLine & _
	                    "without saving, replace this line with your own code."
	                    mybook.Close savechanges:=False
	                End If
	            Else
	                MsgBox "We skip this file : " & MySplit(N) & " because it Is already open"
	            End If
	
	            Next N
	        With Application
	            .ScreenUpdating = True
	            .EnableEvents = True
	        End With
	    End If
	End Sub
	
	Function bIsBookOpen(ByRef szBookName As String) As Boolean
	    ' Rob Bovey
	    On Error Resume Next
	    bIsBookOpen = Not (Application.Workbooks(szBookName) Is Nothing)
	End Function

## Other file formats : ##

In the macro you see this code line that say which file format you can select (xlsx).

FileFormat = "{""org.openxmlformats.spreadsheetml.sheet""}"

If you want more then one format you can use this to be able to also select xls files.

FileFormat = "{""org.openxmlformats.spreadsheetml.sheet"",""com.microsoft.Excel.xls""}"

This is a list of a few formats that you can use :

- xls : com.microsoft.Excel.xls
- xlsx : org.openxmlformats.spreadsheetml.sheet
- xlsm : org.openxmlformats.spreadsheetml.sheet.macroenabled
- xlsb : com.microsoft.Excel.sheet.binary.macroenabled
- csv : public.comma-separated-values-text
- doc : com.microsoft.word.doc
- docx : org.openxmlformats.wordprocessingml.document 
- docm : org.openxmlformats.wordprocessingml.document.macroenabled
- ppt : com.microsoft.powerpoint.ppt
- pptx : org.openxmlformats.presentationml.presentation
- pptm : org.openxmlformats.presentationml.presentation.macroenabled
- txt : public.plain-text