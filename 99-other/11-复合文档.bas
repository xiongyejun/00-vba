Attribute VB_Name = "MFileStruct"
Option Explicit

Type CompdocFileFormat '复合文档
    fh_identifier(7) As Byte '是复合文档的标识,依次是D0、CF、11、E0、A1、B1、1A、E1
    file_identifier(15) As Byte '文件唯一标识
    file_format_revision As Integer '文件格式的修订号
    file_format_version As Integer '文件格式的版本号
    memory_endian As Integer 'FFFEh表示“Little-Endian”
    sector_size As Integer '通常为9，扇区的大小,2的幂
    short_sector_size As Integer '通常为6，短扇区的大小，2的幂
    not_used_1(9) As Byte
    SAT_count As Long '分区表扇区的总数,“sector allocation table”（缩写SAT）
    first_sector_id As Long '目录流第一个扇区的ID
    not_used_2(3) As Byte '
    min_stream_size As Long '最小标准流尺寸，通常为1000h
    SSAT_first_id As Long '短分区表（SSAT）的第一个扇区的ID
    SSAT_count As Long '短分区表扇区总数
    MSAT_first_id As Long '主分区表（缩写MSAT）的第一个扇区的ID
    MSAT_count As Long '分区表的扇区总数

    '109个32位整数，为主分区表的开头109个记录
    '当分区表扇区SAT_count的个数大于109个时，就需要另外的扇区来存储，否则为-1
    arr_id(108) As Long
End Type

Sub test()
    Dim fh As CompdocFileFormat
    Dim num_file As Integer
    Dim file_name As String
    
    file_name = "xx.xls"
    
    num_file = FreeFile
    
    Open file_name For Binary Access Read As #num_file
    
    Get #num_file, 1, fh
    
    Close num_file
  
End Sub
