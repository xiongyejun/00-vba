Attribute VB_Name = "Module1"
Option Explicit

'Important: this Dim line must be at the top of your module
Dim MyFiles As String


Sub TestMacroForThisfileWithCellReferences()
    Dim MySplit As Variant
    Dim FileInMyFiles As Long
    Dim Fstr As String
    Dim LastSep As String

    'Note: I use cell references in this macro to make it easy to test the code
    'Normally you will use it like this :
    'Call GetFilesOnMacWithOrWithoutSubfolders(Level:=1, ExtChoice:=0, FileFilterOption:=0, FileNameFilterStr:="SearchString")

    'Clear MyFiles to be sure that it not return old info if no files are found
    MyFiles = ""

    'Fill the MyFiles string with the files if they match your criteria
    Call GetFilesOnMacWithOrWithoutSubfolders(Level:=Range("F9").Value, ExtChoice:=Range("G9").Value, FileFilterOption:=Range("H9").Value, FileNameFilterStr:=Range("I9").Text)
    'Level                     : 1= Only the files in the folder, 2 to ? levels of subfolders
    'ExtChoice             :  0=(xls|xlsx|xlsm|xlsb), 1=xls , 2=xlsx, 3=xlsm, 4=xlsb, 5=csv, 6=txt, 7=all files, 8=(xlsx|xlsm|xlsb), 9=(csv|txt)
    'FileFilterOption     :  0=No Filter, 1=Begins, 2=Ends, 3=Contains
    'FileNameFilterStr   : Search string used when FileFilterOption = 1, 2 or 3


    'This code below will list all files on the first sheet of this workbook
    'In column A :B the path/name, C the file date/time and D the size
    'You can browse to the folder you want when the code Run

    'In this example we list the file names but you can also use MySplit(FileInMyFiles)
    'in the loop to for example to open the files with Workbooks.Open(MySplit(FileInMyFiles))

    If MyFiles <> "" Then
        With Application
            .ScreenUpdating = False
        End With

        'Delete all cells in columns A:C in the first worksheet of this workbook
        Sheets(1).Columns("A:D").Cells.Clear

        With Sheets(1).Range("A1:D1")
            .Value = Array("Directory", "File Name", "Date/Time", "Size")
            .Font.Bold = True
        End With

        'Split MyFiles and loop through all the files
        MySplit = Split(MyFiles, Chr(13))
        For FileInMyFiles = LBound(MySplit) To UBound(MySplit)
            On Error Resume Next
            Fstr = MySplit(FileInMyFiles)
            LastSep = InStrRev(Fstr, Application.PathSeparator, , 1)
            Sheets(1).Cells(FileInMyFiles + 2, 1).Value = Left(Fstr, LastSep - 1)    'Column A
            Sheets(1).Cells(FileInMyFiles + 2, 2).Value = Mid(Fstr, LastSep + 1, Len(Fstr) - LastSep)    'Column B
            Sheets(1).Cells(FileInMyFiles + 2, 3).Value = FileDateTime(MySplit(FileInMyFiles))    'Column C
            Sheets(1).Cells(FileInMyFiles + 2, 4).Value = FileLen(MySplit(FileInMyFiles))    'Column D
            On Error GoTo 0
        Next FileInMyFiles
        Sheets(1).Columns("A:D").AutoFit
        With Application
            .ScreenUpdating = True
        End With
    Else
        MsgBox "Sorry no files that match your criteria, A 0 files result can be due to Apple sandboxing: Try using the Browse button to re-select the folder."
        'Delete all cells in columns A:D in the first worksheet of this workbook
        Sheets(1).Columns("A:D").Cells.Clear
        'ScreenUpdating is still True but we set it to true again to refresh the screen,
        With Application
            .ScreenUpdating = True
        End With
    End If

End Sub


'*******Function that do all the work that will be called by the macro*********

Function GetFilesOnMacWithOrWithoutSubfolders(Level As Long, ExtChoice As Long, _
                                              FileFilterOption As Long, FileNameFilterStr As String)
'Ron de Bruin,Version 4.0: 27 Sept 2015
'http://www.rondebruin.nl/mac.htm
'Thanks to DJ Bazzie Wazzie and Nigel Garvey(posters on MacScripter)
    Dim ScriptToRun As String
    Dim folderPath As String
    Dim FileNameFilter As String
    Dim Extensions As String

    On Error Resume Next
    folderPath = MacScript("choose folder as string")
    If folderPath = "" Then Exit Function
    On Error GoTo 0

    Select Case ExtChoice
    Case 0: Extensions = "(xls|xlsx|xlsm|xlsb)"  'xls, xlsx , xlsm, xlsb
    Case 1: Extensions = "xls"    'Only  xls
    Case 2: Extensions = "xlsx"    'Only xlsx
    Case 3: Extensions = "xlsm"    'Only xlsm
    Case 4: Extensions = "xlsb"    'Only xlsb
    Case 5: Extensions = "csv"    'Only csv
    Case 6: Extensions = "txt"    'Only txt
    Case 7: Extensions = ".*"    'All files with extension, use *.* for everything
    Case 8: Extensions = "(xlsx|xlsm|xlsb)"  'xlsx, xlsm , xlsb
    Case 9: Extensions = "(csv|txt)"   'csv and txt files
        'You can add more filter options if you want,
    End Select

    Select Case FileFilterOption
    Case 0: FileNameFilter = "'.*/[^~][^/]*\\." & Extensions & "$' "  'No Filter
    Case 1: FileNameFilter = "'.*/" & FileNameFilterStr & "[^~][^/]*\\." & Extensions & "$' "    'Begins with
    Case 2: FileNameFilter = "'.*/[^~][^/]*" & FileNameFilterStr & "\\." & Extensions & "$' "    ' Ends With
    Case 3: FileNameFilter = "'.*/([^~][^/]*" & FileNameFilterStr & "[^/]*|" & FileNameFilterStr & "[^/]*)\\." & Extensions & "$' "   'Contains
    End Select

    folderPath = MacScript("tell text 1 thru -2 of " & Chr(34) & folderPath & _
                           Chr(34) & " to return quoted form of it's POSIX Path")
    folderPath = Replace(folderPath, "'\''", "'\\''")

    If Val(Application.Version) < 15 Then
        ScriptToRun = ScriptToRun & "set foundPaths to paragraphs of (do shell script """ & "find -E " & _
                      folderPath & " -iregex " & FileNameFilter & "-maxdepth " & _
                      Level & """)" & Chr(13)
        ScriptToRun = ScriptToRun & "repeat with thisPath in foundPaths" & Chr(13)
        ScriptToRun = ScriptToRun & "set thisPath's contents to (POSIX file thisPath) as text" & Chr(13)
        ScriptToRun = ScriptToRun & "end repeat" & Chr(13)
        ScriptToRun = ScriptToRun & "set astid to AppleScript's text item delimiters" & Chr(13)
        ScriptToRun = ScriptToRun & "set AppleScript's text item delimiters to return" & Chr(13)
        ScriptToRun = ScriptToRun & "set foundPaths to foundPaths as text" & Chr(13)
        ScriptToRun = ScriptToRun & "set AppleScript's text item delimiters to astid" & Chr(13)
        ScriptToRun = ScriptToRun & "foundPaths"
    Else
        ScriptToRun = ScriptToRun & "do shell script """ & "find -E " & _
                      folderPath & " -iregex " & FileNameFilter & "-maxdepth " & _
                      Level & """ "
    End If
    On Error Resume Next
    MyFiles = MacScript(ScriptToRun)
    On Error GoTo 0
End Function


Sub SortData()
    Dim rng As Range
    On Error Resume Next
    Set rng = Range("A1").CurrentRegion
    rng.Sort key1:=rng.Cells(1, 1), _
             order1:=xlAscending, _
             Header:=xlNo
    Application.ScreenUpdating = True
End Sub

