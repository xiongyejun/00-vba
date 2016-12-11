Attribute VB_Name = "MFileStruct"
Option Explicit

Type CompdocFileFormat '�����ĵ�
    fh_identifier(7) As Byte '�Ǹ����ĵ��ı�ʶ,������D0��CF��11��E0��A1��B1��1A��E1
    file_identifier(15) As Byte '�ļ�Ψһ��ʶ
    file_format_revision As Integer '�ļ���ʽ���޶���
    file_format_version As Integer '�ļ���ʽ�İ汾��
    memory_endian As Integer 'FFFEh��ʾ��Little-Endian��
    sector_size As Integer 'ͨ��Ϊ9�������Ĵ�С,2����
    short_sector_size As Integer 'ͨ��Ϊ6���������Ĵ�С��2����
    not_used_1(9) As Byte
    SAT_count As Long '����������������,��sector allocation table������дSAT��
    first_sector_id As Long 'Ŀ¼����һ��������ID
    not_used_2(3) As Byte '
    min_stream_size As Long '��С��׼���ߴ磬ͨ��Ϊ1000h
    SSAT_first_id As Long '�̷�����SSAT���ĵ�һ��������ID
    SSAT_count As Long '�̷�������������
    MSAT_first_id As Long '����������дMSAT���ĵ�һ��������ID
    MSAT_count As Long '���������������

    '109��32λ������Ϊ��������Ŀ�ͷ109����¼
    '������������SAT_count�ĸ�������109��ʱ������Ҫ������������洢������Ϊ-1
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
