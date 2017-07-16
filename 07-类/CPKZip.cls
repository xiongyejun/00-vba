VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPKZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------ZIP文件结构-------------------------------------------
'ZIP文件结构的说明，下面这个帖子介绍的挺详细
'http://club.excelhome.net/thread-1251530-1-1.html
'把a1.txt、a2.txt、a3.txt压缩到a.zip中以后
'其内容在磁盘上的摆放顺序为h1、a1[]、h2、a2[]、h3、a3[]、c1、c2、c3、EOCD
'其中a1[]、a2[]、a3[]是三个文本文件压缩后的数据块
'h1、h2、h3和c1、c2、c3分别是三个数据块对应的Local File Header和Central Directory FileHeader结构
'EOCD是文件中唯一的EndOfCentralDirectory结构。
'-----------------------------------------ZIP文件结构-------------------------------------------


'-----------------------------------------ZIP文件解压缩-------------------------------------------
'http://www.cnblogs.com/esingchan/p/3958962.html
'-----------------------------------------ZIP文件解压缩-------------------------------------------


'local file header+file data+data descriptor这是一段ZIP压缩数据
'Local file header
Private Type LocalFileHeader
    Signature As Long           '文件头标识 0x04034b50
    VersionExtract As Integer   '解压文件所需最低版本
    GeneralBit As Integer       '通用位标记
    CompressionMethod As Integer '压缩方法
    FileModiTime As Integer     '文件最后修改时间
    FileModiDate As Integer     '文件最后修改日期
    CRC_32  As Long             '说明采用的算法
    CompressedSize As Long      '压缩后的大小
    UncompressedSize As Long    '压缩前的大小
    FileNameLength As Integer      '文件名长度 (n)
    ExtraFieldLength As Integer '附加信息长度 (m)

'    FileName() As Byte          '文件名
'    ExtraField() As Byte        '扩展区
End Type

Private Type CentralDirectoryHeader
    Signature As Long               'HEX 50 4B 01 02
    VersionMadeBy As Integer
    VersionNeeded As Integer
    GeneralBitFlag As Integer
    CompressionMethod As Integer
    LastModifyTime As Integer
    LastModifyDate As Integer
    CRC32 As Long
    CompressedSize As Long
    UncompressedSize As Long
    FileNameLength As Integer       '文件名长度(n)
    ExtraFieldLength As Integer     '附加信息长度 (m)
    FileCommentLength As Integer    '文件附注长度 (k)
    StartDiskNumber As Integer      '文件起始位置的磁盘编号【3】
    InteralFileAttrib As Integer    '内部文件属性
    ExternalFileAttrib As Long      '外部文件属性
    LocalFileHeaderOffset As Long   '对应的Local File  Header在文件中的起始位置。
'                                   46  n 文件名
'                                   46+n    m   附加信息
'                                   46+n+m  k   文件附注
End Type

Private Type EndOfCentralDirectory
    Signature As Long                     '核心目录结束标记 0x06054b50
    NumberOfThisDisk As Integer              '当前磁盘编号
    DiskDirectoryStarts As Integer           '第一条Central  Directory起始位置所在的磁盘编号
    NumberOfCDRecordsOnThisDisk As Integer  '当前磁盘上的Central  Directory数量
    TotalNumberOfCDRecords As Integer       'Zip文件中全部Central  Directory的总数量
    SizeOfCD As Long                        '全部Central  Directory的合计字节长度
    OffsetOfCD As Long                      '第一条Central  directory的起始位置在zip文件中的位置
    CommentLength As Integer                '注释长度
'    Comment() as Byte                       '注释内容
End Type

'记录ZIP文件结构，文件名，文件数据，和解压后的文件大小
Private Type DataType
    FileName As String
    bytData() As Byte
    UncompressedSize As Long
End Type

'分区间的数据
Private Type CodeInterval
    Bits As Byte
    Start As Long
    End As Long
End Type

'扩展用
Private Type MyType
    i_rows As Long
    arr_data(1 To 2, 1 To 1) As Byte
End Type

'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As Long)

Private ZipFileName As String
Private ArrZipFileList() As String      'ZIP文件名数组
Private ArrData() As DataType           '获取数据块

'Private FileByte() As Byte
'Private FileAddress As LongPtr

Private dic_1 As Object 'literal/length Huffman码表1
Private dic_2 As Object 'distance Huffman码表2
'
Property Let FileName(value As String)
    ZipFileName = value
End Property


Property Get ZipFileList() As String()
    ZipFileList = ArrZipFileList
End Property

Property Get FileName() As String
    FileName = ZipFileName
End Property

Function GetData()
    Dim iFreefile As Integer
    Dim bytBuffer() As Byte
    Dim i As Long
    Dim lOffsetEOCD As Long
    Dim lLOF As Long
    Dim oEOCD As EndOfCentralDirectory
    Dim oCDFH As CentralDirectoryHeader
    Dim oLFH As LocalFileHeader, lOffsetLFH As Long, lDataOffset As Long
    Dim lOffset As Long
    
    iFreefile = FreeFile
    
    Open ZipFileName For Binary As iFreefile
        lLOF = LOF(iFreefile)
        ReDim bytBuffer(lLOF - 1) As Byte
        Get iFreefile, 1, bytBuffer
        
        For i = lLOF - 1 - 4 To 0 Step -1
            '查找EndOfCentralDirectory的Signature标识
            If bytBuffer(i) = &H50 And bytBuffer(i + 1) = &H4B And bytBuffer(i + 2) = &H5 And bytBuffer(i + 3) = &H6 Then
                lOffsetEOCD = i + 1
                Exit For
            End If
        Next
        
        If lOffsetEOCD = 0 Then
            Err.Raise 1, , "zip文件格式可能有误，请检查"
            Exit Function
        End If
        '读取EndOfCentralDirectory
        Get iFreefile, lOffsetEOCD, oEOCD
        lOffset = oEOCD.OffsetOfCD + 1
        
        ReDim ArrData(oEOCD.TotalNumberOfCDRecords - 1) As DataType
        ReDim ArrZipFileList(0 To oEOCD.TotalNumberOfCDRecords - 1) As String
        For i = 0 To oEOCD.TotalNumberOfCDRecords - 1
            '读取CentralDirectoryHeader的数据
            Get iFreefile, lOffset, oCDFH
            '根据CentralDirectoryHeader的FileNameLength长度，往后读取文件名
            ReDim bytBuffer(0 To oCDFH.FileNameLength - 1) As Byte
            Get iFreefile, lOffset + Len(oCDFH), bytBuffer
            
            ArrZipFileList(i) = StrConv(bytBuffer, vbUnicode)
            ArrData(i).FileName = ArrZipFileList(i)
            lOffsetLFH = oCDFH.LocalFileHeaderOffset + 1        'VBA里open打开的是从1开始的
            '获取LocalFileHeader
            Get iFreefile, lOffsetLFH, oLFH
'            Debug.Print oLFH.Signature
            ReDim ArrData(i).bytData(oLFH.CompressedSize - 1) As Byte
            '找到数据的开始位置
            lDataOffset = lOffsetLFH + oLFH.FileNameLength + oLFH.ExtraFieldLength + Len(oLFH)
            Get iFreefile, lDataOffset, ArrData(i).bytData
            
            ArrData(i).UncompressedSize = oLFH.UncompressedSize
            
            '下1个CentralDirectoryHeader的位置
            lOffset = lOffset + Len(oCDFH) + oCDFH.FileCommentLength + oCDFH.FileNameLength + oCDFH.ExtraFieldLength
        Next
    Close iFreefile

End Function

'
'    '-----------------------------------------------------------------------------------
'    '用CopyMemory获取的数据不正确
'    'lOffsetLFH = oCDFH.LocalFileHeaderOffset 一直是0
'    'oLFH也是错的
'    '不知道为什么啊
'
'    '用open读取的又是对的，哪错了？
'    '------------------------------------------------------------------------------------
'获取ZIP文件数据结构
'Private Function GetData()
'    Dim iFreefile As Integer
'    Dim i As Long
'    Dim lOffsetEOCD As Long
'    Dim lLOF As Long
'    Dim oEOCD As EndOfCentralDirectory
'    Dim oCDFH As CentralDirectoryHeader
'    Dim oLFH As LocalFileHeader, lOffsetLFH As Long, lDataOffset As Long
'    Dim lOffset As Long
'    Dim FileNameByte() As Byte
'
'    iFreefile = FreeFile
'    '读取ZIP文件数据到FileByte
'    Open ZipFileName For Binary As iFreefile
'    lLOF = LOF(iFreefile)
'    ReDim FileByte(lLOF - 1) As Byte
'    Get iFreefile, 1, FileByte
'    Close iFreefile
'    FileAddress = VarPtr(FileByte(0))
'
'    '查找EndOfCentralDirectory的Signature标识
'    For i = lLOF - 1 - Len(oEOCD) + 4 To 0 Step -1
'        If FileByte(i) = &H50 And FileByte(i + 1) = &H4B And FileByte(i + 2) = &H5 And FileByte(i + 3) = &H6 Then
'            lOffsetEOCD = i
'            Exit For
'        End If
'    Next
'
'    If lOffsetEOCD = 0 Then
'        Err.Raise 1, , "zip文件格式可能有误，请检查"
'        Exit Function
'    End If
'    '读取EndOfCentralDirectory
'    CopyMemory VarPtr(oEOCD.Signature), FileAddress + lOffsetEOCD, Len(oEOCD)
'    'CentralDirectoryHeader的起始位置
'    lOffset = oEOCD.OffsetOfCD
'    lOffsetLFH = 0

'
'    ReDim ArrData(oEOCD.TotalNumberOfCDRecords - 1) As DataType
'    ReDim ArrZipFileList(0 To oEOCD.TotalNumberOfCDRecords - 1) As String
'    For i = 0 To oEOCD.TotalNumberOfCDRecords - 1
'        '读取CentralDirectoryHeader的数据
'        CopyMemory VarPtr(oCDFH.Signature), FileAddress + lOffset, Len(oCDFH)
'
'         '获取LocalFileHeader
''        lOffsetLFH = oCDFH.LocalFileHeaderOffset         '如果用open读取，要注意VBA里open打开的是从1开始的   ,这里第2个也是0，不知道为什么了
''        CopyMemory VarPtr(oLFH.Signature), FileAddress + lOffsetLFH, Len(oLFH)     '这里总是获取不到正确的数据，不知道为什么！下面的oLFH.ExtraFieldLength只能先用0替代下
'
'        ReDim ArrData(i).bytData(oCDFH.CompressedSize - 1) As Byte
'        '找到数据的开始位置
'        lDataOffset = lOffsetLFH + oCDFH.FileNameLength + Len(oLFH) + 0 ' oLFH.ExtraFieldLength
'        CopyMemory VarPtr(ArrData(i).bytData(0)), FileAddress + lDataOffset, oCDFH.CompressedSize
'
'        ArrData(i).UncompressedSize = oCDFH.UncompressedSize
'
'         '根据CentralDirectoryHeader的FileNameLength长度，往后读取文件名
'        ReDim FileNameByte(0 To oCDFH.FileNameLength - 1) As Byte
'        CopyMemory VarPtr(FileNameByte(0)), FileAddress + lOffset + Len(oCDFH), oCDFH.FileNameLength
'
'
'        ArrZipFileList(i) = StrConv(FileNameByte, vbUnicode)
'        ArrData(i).FileName = ArrZipFileList(i)
'
'        '下1个CentralDirectoryHeader的位置
'        lOffset = lOffset + Len(oCDFH) + oCDFH.FileCommentLength + oCDFH.FileNameLength + oCDFH.ExtraFieldLength
'        lOffsetLFH = lOffsetLFH + Len(oLFH) + oCDFH.FileNameLength + oCDFH.CompressedSize
'    Next
'
'End Function

Function InflateAll()
    
End Function

Function InflateIndex(i_index As Long) As Byte()
    On Error GoTo err_handle
'    Dim i As Long
'
'    For i = 0 To UBound(ArrData(i_index).bytData)
'        Cells(i + 1, 1).value = ArrData(i_index).bytData(i)
'    Next
    
    InflateIndex = Inflate(ArrData(i_index).bytData, ArrData(i_index).UncompressedSize)
    Exit Function
    
err_handle:
    InflateIndex = VBA.StrConv("解压出错。", vbFromUnicode)
End Function


Function Inflate(b() As Byte, i_UncompressedSize As Long) As Byte()
    Dim str_bit As String
    Dim i_CL1 As Long, i_CL2 As Long, i_CCL As Long
    Dim i_mid As Long, i As Long
    Dim arr_CCL() As Long, arr_CodeLen() As Long, tmp
    Const STR_SWAP As String = "16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15"
    Dim p_arr As Long
    Dim cls_huffman As CHuffman
    Dim dic As Object
    Dim arr_CL1() As Long, arr_CL2() As Long
    
    Set dic = CreateObject("Scripting.Dictionary")
    
    ReDim arr_CodeLen(18) As Long
    ReDim arr_CCL(18) As Long
    
    str_bit = GetBit(b)
    
    i_mid = 1
'       Header：3个比特，第一个比特如果是1，表示此部分为最后一个压缩数据块；否则表示这是.ZIP文件的某个中间压缩数据块，但后面还有其他数据块。
'   第2、3比特表示3个选择：
'    00 - no compression                                没有使用Huffman
'    01 - compressed with fixed Huffman codes           静态Huffman
'    10 - compressed with dynamic Huffman codes         动态Huffman
'    11 - reserved (error)
    If "01" <> VBA.Mid$(str_bit, 2, 2) Then
        '低比特在先
        Inflate = VBA.StrConv("不是动态Huffman，暂不支持", vbFromUnicode)
        Exit Function
    End If
    
    i_mid = i_mid + 3
    
    '   HLIT：5比特，记录literal/length码树中码长序列（CL1）个数的一个变量。后面CL1个数等于HLIT+257（因为至少有0-255总共256个literal，还有一个256表示解码结束，但length的个数不定）
    i_CL1 = BitToNum(VBA.Mid$(str_bit, i_mid, 5)) + 257
    ReDim arr_CL1(i_CL1 - 1) As Long
    i_mid = i_mid + 5
'       HDIST：5比特，记录distance码树中码长序列（CL2）个数的一个变量。后面CL2个数等于HDIST+1。哪怕没有1个重复字符串，distance都为0也是一个CL。
    i_CL2 = BitToNum(VBA.Mid$(str_bit, i_mid, 5)) + 1
    ReDim arr_CL2(i_CL2 - 1) As Long
    i_mid = i_mid + 5
    
'       HCLEN：4比特，记录Huffman码表3中码长序列（CCL）个数的一个变量。后面CCL个数等于HCLEN+4。PK认为CCL个数不会低于4个，即使对于整个文件只有1个字符的情况。
    i_CCL = BitToNum(VBA.Mid$(str_bit, i_mid, 4)) + 4
    i_mid = i_mid + 4
'       接下来是3比特编码的CCL，一共HCLEN+4个，用以构造Huffman码表3
    For i = 0 To i_CCL - 1
        arr_CodeLen(i) = BitToNum(VBA.Mid$(str_bit, i_mid, 3))
        i_mid = i_mid + 3
    Next i
    
    'PK认为CL序列里面CL范围为0-15，特殊的几个值是16、17、18
    '如果把CCL序列位置置换一下，把16、17、18这些放前面'
    '那么这个CCL序列就很可能最后面跟着一串0（因为CL=14,15这些很可能没有），所以最后还引入了一个置换
    tmp = VBA.Split(STR_SWAP, ",")
    
    For i = 0 To 18
        arr_CCL(i) = tmp(i)  '实际的数据
    Next i

'   清除0
    ArrClearZero arr_CodeLen, arr_CCL
     
    Set cls_huffman = New CHuffman
    cls_huffman.SetArrData arr_CCL, arr_CodeLen
    cls_huffman.PrintBit dic
    Set cls_huffman = Nothing

'   接下来是对CL1（码长）序列经过游程编码（SQ1：缩短的整数序列）后，并对SQ1继续用Huffman编码后的比特流。包含HLIT+257个CL1，其解码码表为Huffman码表3，用以构造Huffman码表1；
    RunLength dic, str_bit, i_mid, arr_CL1
    GetHuffmanOne arr_CL1, i_CL1
'   接下来是对CL2（码长）序列经过游程编码（SQ2：缩短的整数序列）后，并对SQ2继续用Huffman编码后的比特流。包含HDIST+1个CL2，其解码码表为Huffman码表3，用于构造Huffman码表2；
    RunLength dic, str_bit, i_mid, arr_CL2
    GetHuffmanTwo arr_CL2, i_CL2
    
'   先按照Huffman码表1解码，如果解码结果是长度（>256），则接下来按照Huffman码表2解码
    Inflate = UnCompressed(str_bit, i_mid, i_UncompressedSize)
    
End Function

Private Function UnCompressed(str_bit As String, ByRef i_mid As Long, i_UncompressedSize As Long) As Byte()
    Dim str As String
    Dim arr_un_compressed() As Byte
    Dim i_distanse As Long, i_len As Long
    Dim p_arr As Long
    Dim i As Long
    Dim num As Long
    
    ReDim arr_un_compressed(i_UncompressedSize - 1) As Byte
    
    p_arr = -1
    Do Until 0 'i_mid = ArrData(0).UncompressedSize
        Do Until dic_1.Exists(str)
            str = str & VBA.Mid$(str_bit, i_mid, 1)
            i_mid = i_mid + 1
        Loop
        
        num = dic_1(str)
        str = ""
        
        If num < 256 Then
            p_arr = p_arr + 1
            arr_un_compressed(p_arr) = num
        ElseIf num = 256 Then
            Exit Do
        Else
            '接下来按照Huffman码表2解码
            i_len = num - 257 + 3
            Do Until dic_2.Exists(str)
                str = str & VBA.Mid$(str_bit, i_mid, 1)
                i_mid = i_mid + 1
            Loop
            
            i_distanse = dic_2(str)
'            If i_distanse > p_arr Then Stop
            
            str = ""
            '找到了distance、length
            For i = p_arr - i_distanse + 1 To p_arr - i_distanse + i_len
                p_arr = p_arr + 1
                arr_un_compressed(p_arr) = arr_un_compressed(i)
            Next
            
        End If
    Loop
        
    UnCompressed = arr_un_compressed
    
End Function


'Huffman编码后的码字长度不会特别长，PK认为最长不会超过15，也就是树的深度不会超过15
'literal/length码树
Private Function GetHuffmanOne(arr_CL1() As Long, i_CL)
    Dim ArrCode(285 - 257) As CodeInterval
    Dim cls_huffman As CHuffman
    Dim arr_num() As Long
    Dim i  As Long
    
    ReDim arr_num(i_CL) As Long
    For i = 0 To i_CL
        arr_num(i) = i
    Next i
    ArrClearZero arr_CL1, arr_num
    
    Set cls_huffman = New CHuffman
    cls_huffman.SetArrData arr_num, arr_CL1
    dic_1.RemoveAll
    cls_huffman.PrintBit dic_1
    
    ArrCode(0).Bits = 0: ArrCode(0).Start = 3: ArrCode(0).End = 3
    ArrCode(1).Bits = 0: ArrCode(1).Start = 4: ArrCode(1).End = 4
    ArrCode(2).Bits = 0: ArrCode(2).Start = 5: ArrCode(2).End = 5
    ArrCode(3).Bits = 0: ArrCode(3).Start = 6: ArrCode(3).End = 6
    ArrCode(4).Bits = 0: ArrCode(4).Start = 7: ArrCode(4).End = 7
    ArrCode(5).Bits = 0: ArrCode(5).Start = 8: ArrCode(5).End = 8
    ArrCode(6).Bits = 0: ArrCode(6).Start = 9: ArrCode(6).End = 9
    ArrCode(7).Bits = 0: ArrCode(7).Start = 10: ArrCode(7).End = 10
    ArrCode(8).Bits = 1: ArrCode(8).Start = 11: ArrCode(8).End = 12
    ArrCode(9).Bits = 1: ArrCode(9).Start = 13: ArrCode(9).End = 14
    ArrCode(10).Bits = 1: ArrCode(10).Start = 15: ArrCode(10).End = 16
    ArrCode(11).Bits = 1: ArrCode(11).Start = 17: ArrCode(11).End = 18
    ArrCode(12).Bits = 2: ArrCode(12).Start = 19: ArrCode(12).End = 22
    ArrCode(13).Bits = 2: ArrCode(13).Start = 23: ArrCode(13).End = 26
    ArrCode(14).Bits = 2: ArrCode(14).Start = 27: ArrCode(14).End = 30
    ArrCode(15).Bits = 2: ArrCode(15).Start = 31: ArrCode(15).End = 34
    ArrCode(16).Bits = 3: ArrCode(16).Start = 35: ArrCode(16).End = 42
    ArrCode(17).Bits = 3: ArrCode(17).Start = 43: ArrCode(17).End = 50
    ArrCode(18).Bits = 3: ArrCode(18).Start = 51: ArrCode(18).End = 58
    ArrCode(19).Bits = 3: ArrCode(19).Start = 59: ArrCode(19).End = 66
    ArrCode(20).Bits = 4: ArrCode(20).Start = 67: ArrCode(20).End = 82
    ArrCode(21).Bits = 4: ArrCode(21).Start = 83: ArrCode(21).End = 98
    ArrCode(22).Bits = 4: ArrCode(22).Start = 99: ArrCode(22).End = 114
    ArrCode(23).Bits = 4: ArrCode(23).Start = 115: ArrCode(23).End = 130
    ArrCode(24).Bits = 5: ArrCode(24).Start = 131: ArrCode(24).End = 162
    ArrCode(25).Bits = 5: ArrCode(25).Start = 163: ArrCode(25).End = 194
    ArrCode(26).Bits = 5: ArrCode(26).Start = 195: ArrCode(26).End = 226
    ArrCode(27).Bits = 5: ArrCode(27).Start = 227: ArrCode(27).End = 257
    ArrCode(28).Bits = 0: ArrCode(28).Start = 258: ArrCode(28).End = 258
    
    
    '扩展
    Dim arr_key(), arr_code()
    Dim str_add() As String
    Dim i_row As Long, i_col As Long
    Dim str_key As String
    
    arr_key = dic_1.keys()
    arr_code = dic_1.items()
    dic_1.RemoveAll
    For i = 0 To UBound(arr_key)
        If arr_code(i) >= 257 + 8 And arr_code(i) - 257 < 28 Then
                        
            Combin ArrCode(arr_code(i) - 257).Bits, str_add
            
            For i_row = ArrCode(arr_code(i) - 257).Start To ArrCode(arr_code(i) - 257).End
            
                str_key = VBA.CStr(arr_key(i))
                For i_col = UBound(str_add, 2) To 1 Step -1
                    str_key = str_key & str_add(i_row - ArrCode(arr_code(i) - 257).Start + 1, i_col)
                Next
                '表示长度，dic_1里有记录257 = 0 的，所以要+257-3
                dic_1(str_key) = i_row + 257 - 3
            Next
            
        Else
            dic_1(VBA.CStr(arr_key(i))) = arr_code(i)
        End If
        
    Next i

    Set cls_huffman = Nothing
End Function
'distance码树
Private Function GetHuffmanTwo(arr_CL2() As Long, i_CL)
    Dim ArrCode(29) As CodeInterval
    Dim cls_huffman As CHuffman
    Dim arr_num() As Long
    Dim i  As Long, j As Long
    
    ReDim arr_num(i_CL) As Long
    For i = 0 To i_CL
        arr_num(i) = i
    Next i
    ArrClearZero arr_CL2, arr_num
    
    Set cls_huffman = New CHuffman
    cls_huffman.SetArrData arr_num, arr_CL2
    dic_2.RemoveAll
    cls_huffman.PrintBit dic_2
    
    ArrCode(0).Bits = 0: ArrCode(0).Start = 1: ArrCode(0).End = 1
    ArrCode(1).Bits = 0: ArrCode(1).Start = 2: ArrCode(1).End = 2
    ArrCode(2).Bits = 0: ArrCode(2).Start = 3: ArrCode(2).End = 3
    ArrCode(3).Bits = 0: ArrCode(3).Start = 4: ArrCode(3).End = 4
    ArrCode(4).Bits = 1: ArrCode(4).Start = 5: ArrCode(4).End = 6
    ArrCode(5).Bits = 1: ArrCode(5).Start = 7: ArrCode(5).End = 8
    ArrCode(6).Bits = 2: ArrCode(6).Start = 9: ArrCode(6).End = 12
    ArrCode(7).Bits = 2: ArrCode(7).Start = 13: ArrCode(7).End = 16
    ArrCode(8).Bits = 3: ArrCode(8).Start = 17: ArrCode(8).End = 24
    ArrCode(9).Bits = 3: ArrCode(9).Start = 25: ArrCode(9).End = 32
    ArrCode(10).Bits = 4: ArrCode(10).Start = 33: ArrCode(10).End = 48
    ArrCode(11).Bits = 4: ArrCode(11).Start = 49: ArrCode(11).End = 64
    ArrCode(12).Bits = 5: ArrCode(12).Start = 65: ArrCode(12).End = 96
    ArrCode(13).Bits = 5: ArrCode(13).Start = 97: ArrCode(13).End = 128
    ArrCode(14).Bits = 6: ArrCode(14).Start = 129: ArrCode(14).End = 192
    ArrCode(15).Bits = 6: ArrCode(15).Start = 193: ArrCode(15).End = 256
    ArrCode(16).Bits = 7: ArrCode(16).Start = 257: ArrCode(16).End = 384
    ArrCode(17).Bits = 7: ArrCode(17).Start = 385: ArrCode(17).End = 512
    ArrCode(18).Bits = 8: ArrCode(18).Start = 513: ArrCode(18).End = 768
    ArrCode(19).Bits = 8: ArrCode(19).Start = 769: ArrCode(19).End = 1024
    ArrCode(20).Bits = 9: ArrCode(20).Start = 1025: ArrCode(20).End = 1536
    ArrCode(21).Bits = 9: ArrCode(21).Start = 1537: ArrCode(21).End = 2048
    ArrCode(22).Bits = 10: ArrCode(22).Start = 2049: ArrCode(22).End = 3072
    ArrCode(23).Bits = 10: ArrCode(23).Start = 3073: ArrCode(23).End = 4096
    ArrCode(24).Bits = 11: ArrCode(24).Start = 4097: ArrCode(24).End = 6144
    ArrCode(25).Bits = 11: ArrCode(25).Start = 6145: ArrCode(25).End = 8192
    ArrCode(26).Bits = 12: ArrCode(26).Start = 8193: ArrCode(26).End = 12288
    ArrCode(27).Bits = 12: ArrCode(27).Start = 12289: ArrCode(27).End = 16384
    ArrCode(28).Bits = 13: ArrCode(28).Start = 16385: ArrCode(28).End = 24576
    ArrCode(29).Bits = 13: ArrCode(29).Start = 24577: ArrCode(29).End = 32768
    
    '扩展
    Dim arr_key(), arr_code()
    Dim str_add() As String
    Dim i_row As Long, i_col As Long
    Dim str_key As String
    
    arr_key = dic_2.keys()
    arr_code = dic_2.items()
    dic_2.RemoveAll
    For i = 0 To UBound(arr_key)
        If arr_code(i) > 3 Then
            
            Combin ArrCode(arr_code(i)).Bits, str_add
            
            For i_row = ArrCode(arr_code(i)).Start To ArrCode(arr_code(i)).End
            
                str_key = VBA.CStr(arr_key(i))
                For i_col = UBound(str_add, 2) To 1 Step -1
                    str_key = str_key & str_add(i_row - ArrCode(arr_code(i)).Start + 1, i_col)
                Next
                
                dic_2(str_key) = i_row
            Next
            
        Else
            dic_2(VBA.CStr(arr_key(i))) = VBA.Val(ArrCode(arr_code(i)).Start)
        End If
        
    Next i
    
'    For i = 0 To dic_2.Count - 1
'        Debug.Print dic_2.keys()(i), dic_2.items()(i)
'    Next
    
    Set cls_huffman = Nothing
End Function

'游程编码
'0-15是CL1可能出现的值
'16表示除了0以外的其它游程  2比特，记录连续的3-6个
'17表示0游程                3比特，记录连续的3-10个0
'18表示0游程                7比特，记录连续的11-138个0
Private Function RunLength(dic As Object, str_bit As String, ByRef i_mid As Long, arr() As Long)
    Dim str As String
    Dim num As Long, pre_num As Long
    Dim i_rept As Long '重复次数
    Dim num_rept As Long '重复的数字
    Dim i As Long, p_arr As Long
    Dim k_arr As Long
    
    k_arr = UBound(arr) + 1
    p_arr = 0
    Do Until p_arr = k_arr
        Do Until dic.Exists(str)
            str = str & VBA.Mid$(str_bit, i_mid, 1)
            i_mid = i_mid + 1
        Loop
        
        num = dic(str)
        str = ""
        
        Select Case num
            Case 16
                num_rept = pre_num
                i_rept = BitToNum(VBA.Mid$(str_bit, i_mid, 2)) + 3
                i_mid = i_mid + 2
            Case 17
                num_rept = 0
                i_rept = BitToNum(VBA.Mid$(str_bit, i_mid, 3)) + 3
                i_mid = i_mid + 3
            Case 18
                num_rept = 0
                i_rept = BitToNum(VBA.Mid$(str_bit, i_mid, 7)) + 11
                i_mid = i_mid + 7
            Case Else
                num_rept = num
                i_rept = 1
                pre_num = num '出现16的时候，循环的是前面的0-15的数字
        End Select
        
        For i = 1 To i_rept
            arr(p_arr) = num_rept
            p_arr = p_arr + 1
        Next i
        
    Loop
    
End Function

Private Function BitToNum(str_bit As String) As Long
'    按照低比特在先的方式去看
    Dim i As Long
    
    For i = 1 To VBA.Len(str_bit)
        BitToNum = BitToNum + (2 ^ (i - 1)) * VBA.Val(VBA.Mid$(str_bit, i, 1))
    Next i
End Function

'数组中=0的往后排，最后REDIM清除
'arr_num 实际的数字
Private Function ArrClearZero(arr() As Long, arr_num() As Long)
    Dim p_arr As Long
    Dim i As Long
    
    p_arr = -1
    For i = 0 To UBound(arr)
        If arr(i) > 0 Then
            p_arr = p_arr + 1
            arr(p_arr) = arr(i)
            arr_num(p_arr) = arr_num(i)
        End If
    Next i
    
    ReDim Preserve arr(p_arr) As Long
    ReDim Preserve arr_num(p_arr) As Long
End Function


Private Sub Class_Initialize()
    Set dic_1 = CreateObject("Scripting.Dictionary")
    Set dic_2 = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Erase ArrZipFileList
    Set dic_1 = Nothing
    Set dic_2 = Nothing
End Sub


'扩展用的函数
Private Function Combin(i_col As Byte, arr_result() As String) As Long
    Dim arr_MyType() As MyType
    Dim i As Long
    Dim i_row As Long
    Dim max_count As Long
    
    ReDim arr_MyType(1 To i_col) As MyType
    
    max_count = 1
    For i = 1 To i_col
        arr_MyType(i).arr_data(1, 1) = 0
        arr_MyType(i).arr_data(2, 1) = 1
        
        arr_MyType(i).i_rows = 2
        max_count = max_count * 2
    Next i
    
    CombinCols arr_MyType, arr_result
    
'    With Range("A1").Resize(UBound(arr_result, 1), UBound(arr_result, 2))
'        .Resize(1, 10).EntireColumn.Clear
'        .Value = arr_result
'    End With
    
    Combin = UBound(arr_result, 1)
End Function
    
'选择n列的数据进行排列组合
'每1列的数据：
'1、自身重复的次数=其后各列相乘
'2、整列需要循环的次数=其前各列相乘
Private Function CombinCols(arr_MyType() As MyType, arr_result() As String)
    Dim i_cols As Long
    Dim arr_product() As Long   '1自身重复的次数,2整列需要循环的次数
    Dim i As Long, j As Long
    Dim k As Long, m As Long
    
    i_cols = UBound(arr_MyType)
    ReDim arr_product(1 To i_cols, 1 To 2) As Long
    
    For i = 1 To i_cols
        arr_product(i, 1) = 1
        arr_product(i, 2) = 1
        
        '自身重复的次数
        For j = 1 To i - 1
            arr_product(i, 1) = arr_product(i, 1) * arr_MyType(j).i_rows
        Next j
        '整列需要循环的次数
        For j = i + 1 To i_cols
            arr_product(i, 2) = arr_product(i, 2) * arr_MyType(j).i_rows
        Next j
    Next i
    
    ReDim arr_result(1 To (arr_product(1, 1) * arr_product(1, 2) * arr_MyType(1).i_rows), 1 To i_cols) As String
    
    For i = 1 To i_cols
        '整列循环
        For j = 1 To arr_product(i, 1)
            For m = 1 To arr_MyType(i).i_rows
                 '自身重复
                For k = 1 To arr_product(i, 2)
                    arr_result((m - 1) * arr_product(i, 2) + (j - 1) * arr_MyType(i).i_rows * arr_product(i, 2) + k, i) = VBA.CStr(arr_MyType(i).arr_data(m, 1))
                Next k
            Next m
           
        Next j
    Next i
    
End Function

Private Function GetBit(b() As Byte) As String
    Dim arr_bit() As String
    Dim i_len As Long
    Dim i As Long
    Dim j As Long
    Dim tmp As Byte

    i_len = UBound(b)

    ReDim arr_bit(i_len) As String

    For i = 0 To i_len
        tmp = b(i)
        For j = 1 To 8

            arr_bit(i) = arr_bit(i) & VBA.CStr(tmp Mod 2)
            tmp = tmp \ 2
        Next j
    Next i

    GetBit = VBA.Join(arr_bit, "")

    Erase arr_bit
End Function

