[http://www.rondebruin.nl/win/s7/win001.htm](http://www.rondebruin.nl/win/s7/win001.htm)

# Zip file(s) with the default Windows zip program (VBA)
Information #

Copy the code in a Standard module of your workbook, if you just started with VBA see this page.
Where do I paste the code that I find on the internet

Warning: The code below is not supported by Microsoft.
It is not possible to hide the copy dialog when you copy to a zip folder (this is only working with normal folders as far as I know). Also there is no possibility to avoid that someone can cancel the CopyHere operation or that your VBA code will be notified that the operation has been cancelled.

Note: Do not Dim for example FileNameZip as String in the code examples. This must be a Variant, if you change this the code will not work.

If you want to Unzip files see this page on my site.
Unzip file(s) with the default Windows zip program (VBA)

 

Code used by every example macro on this page

Every macro use the sub NewZip and the first example also use both functions.

	Sub NewZip(sPath)
	'Create empty Zip File
	'Changed by keepITcool Dec-12-2005
	    If Len(Dir(sPath)) > 0 Then Kill sPath
	    Open sPath For Output As #1
	    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
	    Close #1
	End Sub
	
	
	Function bIsBookOpen(ByRef szBookName As String) As Boolean
	' Rob Bovey
	    On Error Resume Next
	    bIsBookOpen = Not (Application.Workbooks(szBookName) Is Nothing)
	End Function
	
	
	Function Split97(sStr As Variant, sdelim As String) As Variant
	'Tom Ogilvy
	    Split97 = Evaluate("{""" & _
	                       Application.Substitute(sStr, sdelim, """,""") & """}")
	End Function
 

Examples

There are five examples on this page that you can copy in a normal module of your workbook.
Please read the information good above before you start testing the code below.

## Browse to the folder you want and select the file or files ##

	Sub Zip_File_Or_Files()
	    Dim strDate As String, DefPath As String, sFName As String
	    Dim oApp As Object, iCtr As Long, I As Integer
	    Dim FName, vArr, FileNameZip
	
	    DefPath = Application.DefaultFilePath
	    If Right(DefPath, 1) <> "\" Then
	        DefPath = DefPath & "\"
	    End If
	
	    strDate = Format(Now, " dd-mmm-yy h-mm-ss")
	    FileNameZip = DefPath & "MyFilesZip " & strDate & ".zip"
	
	    'Browse to the file(s), use the Ctrl key to select more files
	    FName = Application.GetOpenFilename(filefilter:="Excel Files (*.xl*), *.xl*", _
	                    MultiSelect:=True, Title:="Select the files you want to zip")
	    If IsArray(FName) = False Then
	        'do nothing
	    Else
	        'Create empty Zip File
	        NewZip (FileNameZip)
	        Set oApp = CreateObject("Shell.Application")
	        I = 0
	        For iCtr = LBound(FName) To UBound(FName)
	            vArr = Split97(FName(iCtr), "\")
	            sFName = vArr(UBound(vArr))
	            If bIsBookOpen(sFName) Then
	                MsgBox "You can't zip a file that is open!" & vbLf & _
	                       "Please close it and try again: " & FName(iCtr)
	            Else
	                'Copy the file to the compressed folder
	                I = I + 1
	                oApp.Namespace(FileNameZip).CopyHere FName(iCtr)
	
	                'Keep script waiting until Compressing is done
	                On Error Resume Next
	                Do Until oApp.Namespace(FileNameZip).items.Count = I
	                    Application.Wait (Now + TimeValue("0:00:01"))
	                Loop
	                On Error GoTo 0
	            End If
	        Next iCtr
	
	        MsgBox "You find the zipfile here: " & FileNameZip
	    End If
	End Sub
 

## Browse to a folder and zip all files in it ##

	Sub Zip_All_Files_in_Folder_Browse()
	    Dim FileNameZip, FolderName, oFolder
	    Dim strDate As String, DefPath As String
	    Dim oApp As Object
	
	    DefPath = Application.DefaultFilePath
	    If Right(DefPath, 1) <> "\" Then
	        DefPath = DefPath & "\"
	    End If
	
	    strDate = Format(Now, " dd-mmm-yy h-mm-ss")
	    FileNameZip = DefPath & "MyFilesZip " & strDate & ".zip"
	
	    Set oApp = CreateObject("Shell.Application")
	
	    'Browse to the folder
	    Set oFolder = oApp.BrowseForFolder(0, "Select folder to Zip", 512)
	    If Not oFolder Is Nothing Then
	        'Create empty Zip File
	        NewZip (FileNameZip)
	
	        FolderName = oFolder.Self.Path
	        If Right(FolderName, 1) <> "\" Then
	            FolderName = FolderName & "\"
	        End If
	
	        'Copy the files to the compressed folder
	        oApp.Namespace(FileNameZip).CopyHere oApp.Namespace(FolderName).items
	
	        'Keep script waiting until Compressing is done
	        On Error Resume Next
	        Do Until oApp.Namespace(FileNameZip).items.Count = _
	        oApp.Namespace(FolderName).items.Count
	            Application.Wait (Now + TimeValue("0:00:01"))
	        Loop
	        On Error GoTo 0
	
	        MsgBox "You find the zipfile here: " & FileNameZip
	
	    End If
	End Sub
 

## Zip all files in the folder that you enter in the code ##

Note: Before you run the macro below change the folder in this macro line
FolderName = "C:\Users\Ron\test\"

	Sub Zip_All_Files_in_Folder()
	    Dim FileNameZip, FolderName
	    Dim strDate As String, DefPath As String
	    Dim oApp As Object
	
	    DefPath = Application.DefaultFilePath
	    If Right(DefPath, 1) <> "\" Then
	        DefPath = DefPath & "\"
	    End If
	
	    FolderName = "C:\Users\Ron\test\"    '<< Change
	
	    strDate = Format(Now, " dd-mmm-yy h-mm-ss")
	    FileNameZip = DefPath & "MyFilesZip " & strDate & ".zip"
	
	    'Create empty Zip File
	    NewZip (FileNameZip)
	
	    Set oApp = CreateObject("Shell.Application")
	    'Copy the files to the compressed folder
	    oApp.Namespace(FileNameZip).CopyHere oApp.Namespace(FolderName).items
	
	    'Keep script waiting until Compressing is done
	    On Error Resume Next
	    Do Until oApp.Namespace(FileNameZip).items.Count = _
	       oApp.Namespace(FolderName).items.Count
	        Application.Wait (Now + TimeValue("0:00:01"))
	    Loop
	    On Error GoTo 0
	
	    MsgBox "You find the zipfile here: " & FileNameZip
	End Sub
 

## Zip the ActiveWorkbook ##

This sub will make a copy of the Activeworkbook and zip it in "C:\Users\Ron\test\" with a date-time stamp. Change this folder or use your default path Application.DefaultFilePath

	Sub Zip_ActiveWorkbook()
	    Dim strDate As String, DefPath As String
	    Dim FileNameZip, FileNameXls
	    Dim oApp As Object
	    Dim FileExtStr As String
	
	    DefPath = "C:\Users\Ron\test\"    '<< Change
	    If Right(DefPath, 1) <> "\" Then
	        DefPath = DefPath & "\"
	    End If
	
	    'Create date/time string and the temporary xl* and Zip file name
	    If Val(Application.Version) < 12 Then
	        FileExtStr = ".xls"
	    Else
	        Select Case ActiveWorkbook.FileFormat
	        Case 51: FileExtStr = ".xlsx"
	        Case 52: FileExtStr = ".xlsm"
	        Case 56: FileExtStr = ".xls"
	        Case 50: FileExtStr = ".xlsb"
	        Case Else: FileExtStr = "notknown"
	        End Select
	        If FileExtStr = "notknown" Then
	            MsgBox "Sorry unknown file format"
	            Exit Sub
	        End If
	    End If
	
	    strDate = Format(Now, " yyyy-mm-dd h-mm-ss")
	    
	    FileNameZip = DefPath & Left(ActiveWorkbook.Name, _
	    Len(ActiveWorkbook.Name) - Len(FileExtStr)) & strDate & ".zip"
	    
	    FileNameXls = DefPath & Left(ActiveWorkbook.Name, _
	    Len(ActiveWorkbook.Name) - Len(FileExtStr)) & strDate & FileExtStr
	
	    If Dir(FileNameZip) = "" And Dir(FileNameXls) = "" Then
	
	        'Make copy of the activeworkbook
	        ActiveWorkbook.SaveCopyAs FileNameXls
	
	        'Create empty Zip File
	        NewZip (FileNameZip)
	
	        'Copy the file in the compressed folder
	        Set oApp = CreateObject("Shell.Application")
	        oApp.Namespace(FileNameZip).CopyHere FileNameXls
	
	        'Keep script waiting until Compressing is done
	        On Error Resume Next
	        Do Until oApp.Namespace(FileNameZip).items.Count = 1
	            Application.Wait (Now + TimeValue("0:00:01"))
	        Loop
	        On Error GoTo 0
	        'Delete the temporary xls file
	        Kill FileNameXls
	
	        MsgBox "Your Backup is saved here: " & FileNameZip
	
	    Else
	        MsgBox "FileNameZip or/and FileNameXls exist"
	
	    End If
	End Sub
 

## Zip and mail the ActiveWorkbook ##

This will only work if you use Outlook as your mail program.

This sub will send a newly created workbook (copy of the Activeworkbook). It save and zip the workbook before mailing it with a date/time stamp. After the zip file is sent the zip file and the workbook will be deleted from your hard disk.

	Sub Zip_Mail_ActiveWorkbook()
	    Dim strDate As String, DefPath As String, strbody As String
	    Dim oApp As Object, OutApp As Object, OutMail As Object
	    Dim FileNameZip, FileNameXls
	    Dim FileExtStr As String
	
	    DefPath = Application.DefaultFilePath
	    If Right(DefPath, 1) <> "\" Then
	        DefPath = DefPath & "\"
	    End If
	
	    'Create date/time string and the temporary xl* and zip file name
	    If Val(Application.Version) < 12 Then
	        FileExtStr = ".xls"
	    Else
	        Select Case ActiveWorkbook.FileFormat
	        Case 51: FileExtStr = ".xlsx"
	        Case 52: FileExtStr = ".xlsm"
	        Case 56: FileExtStr = ".xls"
	        Case 50: FileExtStr = ".xlsb"
	        Case Else: FileExtStr = "notknown"
	        End Select
	        If FileExtStr = "notknown" Then
	            MsgBox "Sorry unknown file format"
	            Exit Sub
	        End If
	    End If
	
	    strDate = Format(Now, " yyyy-mm-dd h-mm-ss")
	
	    FileNameZip = DefPath & Left(ActiveWorkbook.Name, _
	    Len(ActiveWorkbook.Name) - Len(FileExtStr)) & strDate & ".zip"
	
	    FileNameXls = DefPath & Left(ActiveWorkbook.Name, _
	    Len(ActiveWorkbook.Name) - Len(FileExtStr)) & strDate & FileExtStr
	
	
	    If Dir(FileNameZip) = "" And Dir(FileNameXls) = "" Then
	
	        'Make copy of the activeworkbook
	        ActiveWorkbook.SaveCopyAs FileNameXls
	
	        'Create empty Zip File
	        NewZip (FileNameZip)
	
	        'Copy the file in the compressed folder
	        Set oApp = CreateObject("Shell.Application")
	        oApp.Namespace(FileNameZip).CopyHere FileNameXls
	
	        'Keep script waiting until Compressing is done
	        On Error Resume Next
	        Do Until oApp.Namespace(FileNameZip).items.Count = 1
	            Application.Wait (Now + TimeValue("0:00:01"))
	        Loop
	        On Error GoTo 0
	
	        'Create the mail
	        Set OutApp = CreateObject("Outlook.Application")
	        Set OutMail = OutApp.CreateItem(0)
	        strbody = "Hi there" & vbNewLine & vbNewLine & _
	                  "This is line 1" & vbNewLine & _
	                  "This is line 2" & vbNewLine & _
	                  "This is line 3" & vbNewLine & _
	                  "This is line 4"
	
	        On Error Resume Next
	        With OutMail
	            .To = "ron@debruin.nl"
	            .CC = ""
	            .BCC = ""
	            .Subject = "This is the Subject line"
	            .Body = strbody
	            .Attachments.Add FileNameZip
	            .Send   'or use .Display
	        End With
	        On Error GoTo 0
	
	        'Delete the temporary Excel file and Zip file you send
	        Kill FileNameZip
	        Kill FileNameXls
	    Else
	        MsgBox "FileNameZip or/and FileNameXls exist"
	    End If
	End Sub