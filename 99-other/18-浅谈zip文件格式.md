[http://club.excelhome.net/thread-1251530-1-1.html](http://club.excelhome.net/thread-1251530-1-1.html)

# 浅谈zip文件格式 #

在众多压缩文件格式中，Zip格式由于其规范、算法公开、无专利限制的特点得到了广泛的应用。自2007版本开始，Office套件中的文件格式就转向了以zip格式为外壳打包的.docx、.xlsx等。
本示例主要使用了以下方法，实现了以下功能：

1. 通过open for binary方法打开zip文件，分析其中的Local File Header，Central Directory File Header, End of Central Directory Record结构，获得zip压缩包中包含的文件列表，以及各文件压缩后的字节数组；
2. 通过.net Framework中自带的System.IO.Compression.DeflateStream在内存中将压缩后的字节数组还原成原文件的字节数组，可使用代码进一步分析或输出成磁盘文件。
通过上述功能，可以实现以下目的：
1. 实现zip文件的自动化处理，比如搜索包含特定文件名称的zip压缩包，搜索压缩文件中含有特征字符串、字节串的zip压缩包；
2. 在不打开office程序的情况下读取docx、xlsx文件中的特定文本，提取其中嵌入的OLE对象、图片、包等内容。

要辨别一个文件的类型，通常是看它的后缀名，但这种方法不是100%靠得住。比如我们把一个.txt文件修改成.docx，双击的时候会看到word程序会提示一个错误信息。这是因为，不同的文件类型通常有自己特有的二进制结构，要正确读取出文件蕴含的信息，需要按照这种二进制结构的规定到特定的位置找到特定的字节或字节组合，并转换成它要表达的信息。

通常来说，一个文件中会储存一类或多类有意义的信息（比如文本、图像、音频视频等），在这里，我们把存储这些信息的字节称为“数据块”，这些数据块中的字节值是0-255之间的任意值，通常有相当大的随机性，程序无法通过这些字节判断出它们是否代表合法的文字、图像或其他类型，也无法判断这些数据是否完整、是否被人为篡改或错误输出。所以，要想让其他程序可以正确地解读这些信息，需要在数据块的外部（前部、后部或前后）添加一些额外的固定格式的结构，其中最常见的就是各种各样的header（头信息、标头）结构。通常header会含有以下内容：

标识符，通常位于header结构的最前部，比如exe的前两个字节是十六进制的4D 5A （“MZ”），bmp的前两个字节是十六进制的42 4D（“BM”），而jpg文件的前4个字节是十六进制的FF D8 FF E0，同时第9-12个字节是十六进制的4A 46 49 46（“JFIF”）；

文件、结构、数据块的字节长度；

下一个需要处理的结构的相对或绝对偏移量、字节长度；

各文件类型特有的信息，比如图像的宽高值、颜色深度，文字的编码格式，音频的采样率等。

有些Header和数据块是一体的，即一个header对应一个数据块，我们合称之为一条“记录”，而有些Header在文件中只出现一次，相当于是这些记录的索引。如果把一个文件看作一本书，把文件中的各个数据块看作是各章节的内容，前面一类header可以看作是各章节的摘要，后面一类Header则可以看作是书的总目录。后面分析zip文件时，我们会反复用到这个类比。


回到Zip文件的话题，如果把a1.txt、a2.txt、a3.txt压缩到一个名为a.zip的文件中，压缩软件首先要做的一个工作就是读取这三个文件（“原文件”）中的全部字节，然后用某种压缩算法使字节数降低，形成三个数据块（假设为a1[]、a2[]、a3[]三个字节数组），这时它当然不能一股脑地把所有字节依次写到一个文件当中，那样，各数据块长度、原始文件名、压缩算法等信息会全部丢失，文件解压也就无从入手了。所以，压缩软件需要分别在这三个数据块前加个称为Local File Header的header结构，这个结构的长度为30 + n + m个字节，其中m，n是不确定的值，分别代表原始文件名和附加信息的长度，结构具体如下：

	偏移量	字节长度  	说明
		VBA数据类型
	
	0	4（Long）		LocalFileHeader标识符，依次为50 4B 03 04，对应长整数&H04034B50，其中前两个字节对应AscII字符为PK，是Zip算法发明者Phil Katz的首字母缩写。
	
	4	2（Integer）	解压文件所需的最低zip版本
	
	6	2（Integer）	General purpose bit flag
	
	8	2（Integer）	压缩方法，通常为8
	
	10	2（Integer）	文件最后修改时间【1】
	12	2（Integer）	文件最后修改日期【1】
	14	4（Long）		CRC-32
	18	4（Long）		压缩后的字节长度
	22	4（Long）		压缩前的字节长度
	26	2（Integer）	文件名长度 (n)
	28	2（Integer）	附加信息长度 (m)
	30	n（byte(n)）	文件名【2】
	30+n m (byte(m))	附加信息【2】


对应的VBA数据类型为：

	Private Type LocalFileHeader
	    LocalHeaderSignature AsLong     'HEX 50 4B 03 04
	    VersionNeeded As Integer
	    GeneralPurposeBitFlag As Integer
	    CompressionMethod As Integer
	    LastModifyTime As Integer
	    LastModifyDate As Integer
	    CRC32 As Long
	    CompressedSize As Long
	    UncompressedSize As Long
	    FileNameLength As Integer
	    ExtraFieldLength As Integer
	End Type

注：
【1】这里的文件最后修改日期/时间是DOS日期/时间，可通过API函数DosDateTimeToFileTime和FileTimeToSystemTime转换为我们平时使用的日期数字。

【2】由于VBA的自定义类型中不能包含变长字节数组，所以上面的LocalFileHeader结构中未包含最后两个字段（文件名和附加信息）。

这样一来，通过在a1[]、a2[]、a3[]三个数据块前面分别添加h1、h2、h3三个header，压缩后的文件就有了初步的条理，相当于一本书中划分出了章节，并在每章的头部添加了摘要信息。

这时，如果创建一个a.zip文件，把h1、a1[]、h2、a2[]、h3、a3[]的字节内容依次写入a.zip中，之后是可以通过解析a.zip还原出a1.txt、a2.txt和a3.txt三个文件的（方法是先把a.zip的前30个字节读入到一个LocalFileHeader型的变量h1中，这时30 + h1. FileNameLength + h1. ExtraFieldLength即是数据块a1[]在压缩文件a.zip中的位置， h1.CompressedSize即是数据块a1[]的长度，从而获取到a1[]的全部字节，即a1.txt文件压缩后的全部字节。同时，30 + h1. FileNameLength + h1. ExtraFieldLength + h1.CompressedSize即是第二个LocalFileHeader在a.zip文件中的位置，依此类推，即可得到三个压缩后的数据块a1[]、a2[]、a3[]，进而通过解压算法还原出a1.txt、a2.txt、a3.txt）。

然而，上面的方法虽然可行但并不完善，因为照上面的方法，如果用户只想解压出a3.txt，也要按照顺序从头分析a.zip文件，压缩包中文件不多时尚无大碍，文件一多就会影响到解压的效率了。想像一下你在阅读一本划分了章节但没有目录的大部头书籍，你就能体会这种文件格式设计的缺陷了。

不过不用担心，zip格式的发明者想到了这个问题，所以他设计了一种叫做Central Directory File Header的header结构，结构具体如下：

	偏移量	字节长度  	说明
	0		4		Central Directory Header标识符，依次为50 4B 01 02，对应长整数&H02014B50
	
	4		2		压缩所用的zip版本
	
	6		2		解压缩所需的最低zip版本
	
	8		2		General purpose bit flag
	
	10		2		压缩方法
	
	12		2		文件最后修改时间
	
	14		2		文件最后修改日期
	
	16		4		CRC-32
	
	20		4		压缩后的字节长度
	
	24		4		压缩前的字节长度
	
	28		2		文件名长度 (n)
	
	30		2		附加信息长度 (m)
	
	32		2		文件附注长度 (k)
	
	34		2		文件起始位置的磁盘编号【3】
	
	36		2		内部文件属性
	
	38		4		外部文件属性
	
	42		4		对应的Local File  Header在文件中的起始位置。
	
	46		n		文件名
	
	46+n	m		附加信息
	
	46+n+m	k		文件附注

对应的VBA数据格式如下：
	
	Private Type CentralDirectoryHeader
	    CDFHeaderSignature As Long      'HEX 50 4B 01 02
	    VersionMadeBy As Integer
	    VersionNeeded As Integer
	    GeneralBitFlag As Integer
	    CompressionMethod As Integer
	    LastModifyTime As Integer
	    LastModifyDate As Integer
	    CRC32 As Long
	    CompressedSize As Long
	    UncompressedSize As Long
	    FileNameLength As Integer
	    ExtraFieldLength As Integer
	    FileCommentLength As Integer
	    StartDiskNumber As Integer【3】
	    InteralFileAttrib As Integer
	    ExternalFileAttrib As Long
	    LocalFileHeaderOffset As Long
	End Type

注：
【3】zip出现的时候，软盘还是主流的存储介质，所以经常会有分卷压缩的情况，即一个大文件需要分割到好几张磁盘上，要解压zip压缩包中的一个大文件，需要知道该文件起始位置的数据存储在哪张磁盘上，所以才有了这个字段，对于非分卷压缩的文件，这个字段为0（即第一张磁盘）。简单起见，附件代码中未考虑分卷压缩的情况。

可以看到Central Directory File Header的大部分字段和LocalFile Header是重合的，这种冗余提供了一定的纠错能力。在zip文件中，Central Directory File Header同样与各个数据块是一一对应的，但它们统一放在了zip文件的尾部（即a.zip文件中先是h1、a1[]、h2、a2[]、h3、a3[]，然后是三个Central Directory File Header结构c1、c2、c3），而非像大多数文件那样把总目录放在文件的头部，这种摆放是有其道理的。

以上面的a.zip为例，想象一下，如果a.zip的结构是c1、c2、c3、h1、a1[]、h2、a2[]、h3、a3[]，这时若向a.zip中添加一个新的文件a4.txt会出现什么情况，文件头部会增加一条Central Directory File Headr，相应的h1、a1[]、….h3、a3[]都要向后移动，如果向一个500Mb的压缩包中添加一个1k的文件，整个文件的500Mb内容都要重写一遍！而如果a.zip的结构是h1、a1[]、h2、a2[]、h3、a3[]、c1、c2、c3，这时向其中添加a4.txt时，只需在a3[]后追加h4、a4[]、c1、c2、c3、c4即可，除a4[]数据块外，只需重写几百个字节即可，特别是在软盘时代，这代表着向分卷的zip压缩包中添加文件时，只需要插入最后一张磁盘即可，而需要删除zip压缩包中的文件时，只需要把该数据块对应的Central Directory File Header删除即可，虽然有些浪费空间，但节省了重新打包的时间.

事情总有其两面性，把所有的Central Directory File Header统一写到zip文件尾部虽然可以提高追加文件的效率，但是不利于查找第一条Central Directory File Header的起始位置。仍然以书籍做类比，如果目录在正文前面，不论目录有多长，都很容易查找到，封面、封二、前言、目录，翻不了几页就看到了；而如果目录在正文后面，而且目录内容较长的话，找起目录第一页所在的位置就不那么容易了。这时我们有个简单的办法：在封底上写明“本书目录始于第560页”。Zip文件也是这样做的，在所有的Central Directory File Header都写完之后，后面又追加了一个Endof Central Directory（EOCD）结构，具体如下：

	偏移量	字节长度  	说明
	
	0		4		EOCD标识符，依次为50 4B 05 06，对应长整数&H06054B50
	
	4		2		当前磁盘的序号
	
	6		2		第一条Central  Directory起始位置所在的磁盘编号
	
	8		2		当前磁盘上的Central  Directory数量
	
	10		2		Zip文件中全部Central  Directory的总数量
	
	12		4		全部Central  Directory的合计字节长度
	
	16		4		第一条Central  directory的起始位置在zip文件中的位置
	
	20		2		附注信息长度 (n)
	
	22		n		附注信息

对应的VBA数据格式如下：

	Private Type EndOfCentralDirectory  ‘EOCD
	    EOCDSignature AsLong           'HEX 504B 05 06
	    NumberOfThisDisk As Integer
	    DiskDirectoryStarts As Integer
	    NumberOfCDRecordsOnThisDisk AsInteger
	    TotalNumberOfCDRecords As Integer
	    SizeOfCD As Long
	    OffsetOfCD As Long
	    CommentLength As Integer
	End Type

总结起来，把a1.txt、a2.txt、a3.txt压缩到a.zip中以后，其内容在磁盘上的摆放顺序为h1、a1[]、h2、a2[]、h3、a3[]、c1、c2、c3、EOCD，其中a1[]、a2[]、a3[]是三个文本文件压缩后的数据块，h1、h2、h3和c1、c2、c3分别是三个数据块对应的Local File Header和Central Directory FileHeader结构，EOCD是文件中唯一的EndOfCentralDirectory结构。原来是这样？！嗯，就是这样！

了解了zip文件的格式之后，我们就可以反向解析获取其中全部的文件名和压缩数据块了。步骤如下：

1. 通过open [文件名] for binary for [文件编号]语句以二进制方式打开zip文件；
2. 通过get [文件编号]，起始位置，[字节数组]把文件尾部一定数量的字节读取到字节数组中，用遍历的方法查找十六进制的50 4B 0506，找到EOCD的起始位置；
3. 通过get [文件编号]，EOCD起始位置，[EOCD类型变量]把获取EOCD全部内容；
4. 用EOCD.OffsetOfCD获取c1的起始位置，同样用get语句将c1读取到一个CentralDirectoryFileHeader类型的变量中；
5. 46 + c1. FileNameLength + c1.ExtraFieldLength +c1.FileCommentLength为c2的起始位置，依此类推。
6. c1. LocalFileHeaderOffset为h1的起始位置，用get语句将h1读取到一个LocalFileHeader类型的变量中；
7. 30 + h1. FileNameLength + h1.ExtraFieldLength为数据块a1[]的起始位置，h1.CompressedSize为a1[]的字节长度，用get语句将a1[]读取到字节数组中，a2[]、a3[]同理。

通过上述步骤，我们可以获得压缩文件中所有文件压缩后的数据。接下来只要用想办法解压缩就可以获取到各文件的原始数据了(未完待续)。

UnzipLibrary.DeflateClass.Decompress函数定义如下：

Function Decompress(ptrCompressArray As Long, lenCompressArray As Long, ptrDecompressArray As Long, lenDecompressArray As Long) As Long

- ptrCompressArray表示压缩数据块字节数组的指针，即varPtr(arrIn(0)); 
- lenCompressArray表示压缩字节的字节长度;
- ptrDecompressArray表示接受解压缩数据块字节数组的指针，即varPtr(arrOut(0)); 
- lenDecompressArray表示解压缩后字节的字节长度;

上面的arrIn和arrOut分别是压缩字节和解压缩字节的数组，调用Decompress前需要在arrIn中填充压缩数据块的字节，并用Redim为arrOut设置正确的字节长度，即LocalFileHeader结构中的UncompressedSize字段。

理想状况下，这个函数定义为Decompress(ArrayIn() as byte, ArrayOut() as byte) as Long会比较易于使用，但因为.net是托管代码，不能直接传入和传出非托管的Byte数组（或者我没找到正确的方法），所以传入两个数组指针和数组长度作为变通