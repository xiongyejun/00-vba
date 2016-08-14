'http://club.excelhome.net/thread-1251530-1-1.html

Attribute VB_Name = "模块1"
Private Type EOCD  'CD=Central Directory
    EOCDSignature As Long
    NumberOfThisDisk As Integer
    DiskDirectoryStarts As Integer
    NumberOfCDRecordsOnThisDisk As Integer
    TotalNumberOfCDRecords As Integer
    SizeOfCD As Long
    OffsetOfCD As Long
    CommentLength As Integer
End Type

Private Type CDFheader
    CDFHeaderSignature As Long
    PlaceHolder(0 To 23) As Byte
    FileNameLength As Integer
    ExtraFieldLength As Integer
    FileCommentLength As Integer
    PlaceHolder1(0 To 11) As Byte
End Type


'函数定义：GetZipFileList(sZipFile as String)
'参数：sZipFile，字符串，表示有效的.zip文件路径
'返回值：字符串数组，sZipFile压缩文件中所有文件名称的列表
'参考：https://en.wikipedia.org/wiki/Zip_(file_format)#cite_ref-appnote_25-1
Function GetZipFileList(sZipFile As String) As String()
    Dim iFreefile As Integer
    Dim bytBuffer() As Byte
    Dim i As Integer
    Dim lOffsetEOCD As Long
    Dim lLOF As Long
    Dim oEOCD As EOCD
    Dim oCDFH As CDFheader
    Dim lOffset As Long
    Dim sOutput() As String
    
    iFreefile = FreeFile
    ReDim bytBuffer(255) As Byte
    Open sZipFile For Binary As iFreefile
        lLOF = LOF(iFreefile)
        Get iFreefile, lLOF - 256, bytBuffer
        For i = 0 To 252
            If bytBuffer(i) = &H50 And bytBuffer(i + 1) = &H4B And bytBuffer(i + 2) = &H5 And bytBuffer(i + 3) = &H6 Then
                lOffsetEOCD = lLOF - 256 + i
                Exit For
            End If
        Next
        If lOffsetEOCD = 0 Then
            Err.Raise 1, , "zip文件格式可能有误，请检查"
            Exit Function
        End If
        Get iFreefile, lOffsetEOCD, oEOCD
        lOffset = oEOCD.OffsetOfCD + 1
        
        For i = 0 To oEOCD.TotalNumberOfCDRecords - 1
            Get iFreefile, lOffset, oCDFH
            ReDim bytBuffer(0 To oCDFH.FileNameLength - 1) As Byte
            Get iFreefile, lOffset + Len(oCDFH), bytBuffer
            ReDim Preserve sOutput(0 To i) As String
            sOutput(i) = StrConv(bytBuffer, vbUnicode)
            lOffset = lOffset + Len(oCDFH) + oCDFH.FileCommentLength + oCDFH.FileNameLength + oCDFH.ExtraFieldLength
        Next
    Close iFreefile
    GetZipFileList = sOutput
End Function

Sub Test()
    Dim sFileName As String
    Dim i As Integer
    Dim s
    sFileName = Application.GetOpenFilename("所有,*.*", , "请选择要查看的zip文件")
    If sFileName = "False" Then Exit Sub
    s = GetZipFileList(sFileName)
    If Not IsEmpty(s) Then
        Sheet1.Range("a:b").Clear
        Sheet1.Cells(1, 1) = sFileName
        For i = 0 To UBound(s)
            Sheet1.Cells(i + 2, 2) = s(i)
        Next
    End If
End Sub
