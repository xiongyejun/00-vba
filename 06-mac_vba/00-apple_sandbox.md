[http://www.rondebruin.nl/mac/mac034.htm](http://www.rondebruin.nl/mac/mac034.htm)
# Problems with Apple’s sandbox requirements and Mac Office 2016 with VBA code #
# 苹果沙盒和Mac Office2016 VBA 之间的问题 #

In Windows Excel 97-2016 and in Mac Excel 2011 you can open files or save files where you want in almost every folder on your system without warnings or requests for permission. But in Mac Office 2016 Microsoft have to deal with Apple’s sandbox requirements. When you use VBA in Mac Excel 2016 that Save or Open files you will notice that it is possible that it will ask you permission to access the file or folder (Grant File Access Prompt), this is because of Apple’s sandbox requirements.This means that when you want to save/open files or check if it exists with VBA code the first time you will be prompted to allow access on the first attempt to access such a folder or file.

在Windows Excel 97 - 2016、Mac Excel 2011打开文件或保存文件,在每一个文件夹都没有警告或请求许可。但在Mac Office 2016微软不得不处理苹果的沙箱的要求。当使用VBA Mac Excel 2016中保存或打开文件,会提示许可来访问文件或文件夹(批准文件访问提示)。这意味着当你第一次使用VBA代码想要保存/打开文件或检查它是否存在，系统将提示您在第一次尝试访问允许访问这样一个文件夹或文件。

## How to avoid problems ##
## 如何避免这个问题 ##

There are a few places on your Mac that you can use to avoid the prompts and let your code do what it needs to do without user interaction. But these folders are not in a place that a user can easily find so below are some steps that I hope to make it easier for you to access the folder manual if you want.

在Mac电脑上有几个地方可以避免提示,让代码做需要做的事,不需要用户交互。但这些文件夹不是在一个用户可以很容易的找到的地方,所以下面是一些步骤,我希望能使你更容易访问的文件夹手动如果你想。

This is the Root folder on my machine that we use in the examples on this page:
/Users/rondebruin/Library/Group Containers/UBF8T346G9.Office
**Note:** rondebruin is the user name in this path and I agree that the naming of the folder for Office(UBF8T346G9.Office) is not so nice, but Microsoft must use that of Apple.

The folder above you can use to share data between Office programs or with a third party application, so this location will always work if you want to have read and write access. If you want to have a location only for Excel for example use this path : /Users/rondebruin/Library/Containers/com.microsoft.Excel/Data
I not use this location on this example page to be sure that every Office program can access my files if this is needed.

## Manual create a folder for your Excel files in the Office folder ##
## 手动创建一个文件夹给你办公室的Excel文件的文件夹 ##

1. Open a Finder Window
1. Hold the Alt key when you press on Go in the Finder menu bar
1. Click on Library
1. Open the Group Containers folder
1. Open the UBF8T346G9.Office folder
1. Create a Folder inside this folder named MyExcelFolder for example
1. Select this folder

This are three ways to easily open the folder manual :

- Add it to your Favorites in Finder by dragging it to it.
- Add it to your Favorites in Finder with the shortcut : cmd Ctrl T
- Drag the folder to the Desktop with the CMD and Alt key down. You now have a link(alias) to the folder on your desktop so it is easy to find it and open it in the future.

** Note :** Adding the folder to your Favorites is my favorite because you see the folder in your open and save dialogs in Excel.

## How to create a Folder in the Office folder with VBA code ##

Below you find a macro and a function that you can use to create a folder if it not exists in the Root folder named : UBF8T346G9.Office

In the macro you see one line that call the function and the argument is the name of the folder that you want to create if it not exists. Change "MyProject" to something else to create another folder.

	Sub MakeFolderinMacOffice2016()
	    'Create folder if it not exists in the Microsoft Office Folder
	    'This macro use the custom function
	    'named : CreateFolderinMacOffice2016
	    Call CreateFolderinMacOffice2016(NameFolder:="MyProject")
	End Sub
	
	Function CreateFolderinMacOffice2016(NameFolder As String) As String
	    'Function to create folder if it not exists in the Microsoft Office Folder
	    'Ron de Bruin : 8-Jan-2016
	    Dim OfficeFolder As String
	    Dim PathToFolder As String
	    Dim TestStr As String
	
	    OfficeFolder = MacScript("return POSIX path of (path to desktop folder) as string")
	    OfficeFolder = Replace(OfficeFolder, "/Desktop", "") & _
	        "Library/Group Containers/UBF8T346G9.Office/"
	
	    PathToFolder = OfficeFolder & NameFolder
	
	    On Error Resume Next
	    TestStr = Dir(PathToFolder, vbDirectory)
	    On Error GoTo 0
	    If TestStr = vbNullString Then
	        MkDir PathToFolder
	        'You can use this msgbox line for testing if you want
	        'MsgBox "You find the new folder in this location :" & PathToFolder
	    End If
	    CreateFolderinMacOffice2016 = PathToFolder
	End Function


**Note:** On this page there is a example to create a shortcut on the Desktop to the folder.

[Make folder in the Office folder in Office 2016 and create shortcut on the Desktop with VBA](http://www.rondebruin.nl/mac/mac035.htm)

## How do I open files with VBA code in my folder ? ##

Below you find a macro and a function that you can use to open a file in one of the sub folders of the UBF8T346G9.Office folder. In the macro you see one line that call the function and there are two arguments :

1. Name of the sub folder
1. Name of the file

**Note :** You can also add code in the macro to test if the file is already open, I use that also in the code example in this section : Browse to a file or files in a sub folder of the Office folder.

	Sub OpenFileinMacOffice2016()
	    'Open a file in a sub folder of the Office folder
	    'This macro use the custom function
	    'named : MacOffice2016OpenFile
	    Dim FileString As String
	    Dim wb As Workbook
	    FileString = MacOffice2016OpenFile(FolderName:="MyExcelFolder", FileName:="ron.xlsm")
	    If FileString <> "Error" Then
	        Set wb = Workbooks.Open(FileString)
	        'You can do with wb now what you want
	    End If
	End Sub
	
	Function MacOffice2016OpenFile(FolderName As String, FileName As String) As String
	    'Function to open a file in a sub folder of the Office folder
	    'Ron de Bruin : 8-Jan-2016
	    Dim OfficeFolder As String
	    Dim PathToFile As String
	    Dim TestStr As String
	    Dim wb As Workbook
	
	    OfficeFolder = MacScript("return POSIX path of (path to desktop folder) as string")
	    OfficeFolder = Replace(OfficeFolder, "/Desktop", "") & _
	        "Library/Group Containers/UBF8T346G9.Office/"
	    PathToFile = OfficeFolder & FolderName & Application.PathSeparator & FileName
	
	    On Error Resume Next
	    TestStr = Dir(PathToFile, vbDirectory)
	    On Error GoTo 0
	    If TestStr = vbNullString Then
	        MacOffice2016OpenFile = "Error"
	        MsgBox "Sorry the file not exists in this location : " & PathToFile
	    Else
	        MacOffice2016OpenFile = PathToFile
	    End If
	End Function

## How do I Save a file with VBA code in my folder ? ##

The first macro create a file of only the activesheet and save it in a folder named: ProjectName and the second macro save a copy of the file in a folder named Backup. Both are sub folders of your UBF8T346G9.Office folder.

**Note :** Both macros use the custum function CreateFolderinMacOffice2016 that you find in the first section of this page.

	Sub SaveAsExcel2016()
	    'Save only the activesheet with a Date/time stamp in a sub folder
	    'in the Microsoft Office Folder
	    'This macro use the custom function named : CreateFolderinMacOffice2016
	    Dim Folderstring As String
	    Dim Sourcewb As Workbook
	    Dim Destwb As Workbook
	    Dim FileExtStr As String
	    Dim FileFormatNum As Long
	    Dim FileName As String
	
	    'Create folder if it not exists in the Microsoft Office Folder
	    Folderstring = CreateFolderinMacOffice2016(NameFolder:="ProjectName")
	
	    'set reference to the Active Workbook
	    Set Sourcewb = ActiveWorkbook
	
	    'Copy the ActiveSheet to a new workbook
	    'You can also use Sheets("MySheetName").Copy
	    ActiveSheet.Copy
	    Set Destwb = ActiveWorkbook
	
	    'Determine file extension/format
	    With Destwb
	        Select Case Sourcewb.FileFormat
	            Case 52: FileExtStr = ".xlsx": FileFormatNum = 52
	            Case 53:
	                If .HasVBProject Then
	                    FileExtStr = ".xlsm": FileFormatNum = 53
	                Else
	                    FileExtStr = ".xlsx": FileFormatNum = 52
	                End If
	            Case 57: FileExtStr = ".xls": FileFormatNum = 57
	            Case Else: FileExtStr = ".xlsb": FileFormatNum = 51
	        End Select
	    End With
	
	    '    'Change all cells in the worksheet to values if you want
	    '    With Destwb.Sheets(1).UsedRange
	    '        .Cells.Copy
	    '        .Cells.PasteSpecial xlPasteValues
	    '        .Cells(1).Select
	    '    End With
	    '    Application.CutCopyMode = False
	
	
	    'Name the file and Save it
	    FileName = "Part of " & Sourcewb.Name & " " & Format(Now, "dd-mmm-yy h-mm-ss")
	    With Destwb
	        .SaveAs Folderstring & Application.PathSeparator & FileName & _
	            FileExtStr, FileFormat:=FileFormatNum
	    End With
	
	    'Close the file
	    Destwb.Close False
	    MsgBox "You find a workbook with the active sheet in this folder :" & Folderstring
	End Sub
	
	
	Sub SaveCopyAsExcel2016()
	    'Save a copy of the file with a Date/time stamp in a sub folder
	    'in the Microsoft Office Folder
	    'This macro use the custom function named : CreateFolderinMacOffice2016
	    Dim Folderstring As String
	    Dim wb As Workbook
	    Dim StrFilePath As String
	    Dim StrFileName As String
	    Dim FileExtStr As String
	
	    'Create folder if it not exists in the Microsoft Office Folder
	    Folderstring = CreateFolderinMacOffice2016(NameFolder:="Backup")
	
	    Set wb = ActiveWorkbook
	
	    StrFilePath = Folderstring & Application.PathSeparator
	    StrFileName = "Copy of " & wb.Name & " " & Format(Now, "dd-mmm-yy h-mm-ss")
	    FileExtStr = "." & LCase(Right(wb.Name, Len(wb.Name) - InStrRev(wb.Name, ".", , 1)))
	
	    With wb
	        .SaveCopyAs StrFilePath & StrFileName & FileExtStr
	    End With
	
	    MsgBox "You find a copy of the workbook in this folder :" & StrFilePath
	End Sub
	 

## Browse to a file or files in a sub folder of the Office folder ##

In the example below it opens a browse dialog with a folder folder named : MyExcelFolder from your UBF8T346G9.Office folder and you are only able to select xlsx files. Below the macro you find a list of format names and you can read how you can change it. Note: Do not forget to copy the bIsBookOpen function in your module, you find it below the macro.

	Sub Select_File_Or_Files_Mac_Office_2016()
	    'Select files in sub folder of the Mac Office folder
	    'Working only in Mac Office 2016
	    'http://www.rondebruin.nl/mac/mac034.htm
	    'Ron de Bruin, 20 March 2016
	    Dim NameFolder As String
	    Dim OfficeFolder As String
	    Dim MyPath As String
	    Dim MyScript As String
	    Dim MyFiles As String
	    Dim MySplit As Variant
	    Dim N As Long
	    Dim Fname As String
	    Dim mybook As Workbook
	    Dim OneFile As Boolean
	    Dim FileFormat As String
	
	    'Fill in the name of the folder where the files are that you want to select
	    'Note: this must be a subfolder of your Office folder
	    NameFolder = "MyExcelFolder"
	
	    'In this example you can only select xlsx files
	    'See my webpage how to use other and more formats.
	    FileFormat = "{""org.openxmlformats.spreadsheetml.sheet""}"
	
	    ' Set to True if you only want to be able to select one file
	    ' And to False to be able to select one or more files
	    OneFile = True
	
	    On Error Resume Next
	    OfficeFolder = MacScript("return POSIX path of (path to desktop folder) as string")
	    OfficeFolder = Replace(OfficeFolder, "/Desktop", "") & "Library/Group Containers/UBF8T346G9.Office/"
	
	    MyPath = MacScript("return POSIX file(" & _
	        Chr(34) & OfficeFolder & NameFolder & Chr(34) & ")")
	
	    'Building the applescript string, do not change this
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


