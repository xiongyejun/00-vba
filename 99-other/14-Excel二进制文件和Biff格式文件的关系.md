[http://club.excelhome.net/thread-600816-1-1.html](http://club.excelhome.net/thread-600816-1-1.html)

# 关于Excel二进制文件和Biff格式文件的关系 #

通常我们把excel的文件格式叫做biff，即Binary Interchange File Format的缩写，随着excel的不断升级，其文件格式biff也在同步的更新与完善，因此对于不同版本的excel有着不同的biff版本。

其中buff8x为biff8的扩展版本，它在biff8的基础上改动了部分属性值。

	excel版本		扩展名	biff版本
	4.0			.xlw	无
	5.0 95		.xls	5
	8.0 97		.xls	8
	9.0 2000	.xls	8
	10.0 XP		.xls	8x
	11.0 2003	.xls	8x
	12.0 2007	.xlsb	12
	14.0 2010	.xlsb	12

大家都知道2007能选择.xlsb保存为二进制文件，减少容量的同时提高安全性。有人问excel2003能不能另存为二进制文件？

其实2003中的.xls本身就是二进制文件，.doc和.ppt都是，只不过对应的biff版本低一些。

认为2003默认文件格式不是二进制，是一个误区。

Microsoft Office 2007中，微软引入了一种全新的文档格式：Open XML。由于Open XML是一种开放的文档格式（基于两种开放技术：XML、Zip），所以解决了过去Microsoft Office所使用的二进制文档难以交互、难以被第三方应用程序访问的问题。但是自从微软决定将Open XML提交给ISO之后，从业界的反馈来看，很多人仍然非常关心过去的二进制文档格式（.doc, .xls, .ppt），并希望能得到其相关的技术细节文档。经过慎重的考虑，Microsoft决定将Microsoft Office所使用的二进制文档格式公开。任何人和企业，在不违反相关协议的前提下，都可以免费得到其技术规范文件。

BIFF8基于微软的复合文档格式。Excel文档的内容存放在一个Stream里。一个新建的空白Excel文件，一般包含Root Storge，Workbook Stream，<05H>SummaryInformation Stream，<05H>DocumentSummaryInformation Stream。Workbook Stream就存储了Excel的内容。把整个Excel文件叫Workbook，Workbook里至少有一个Worksheet。Excel文件不仅存放Excel的文本内容，还存放一些格式，比如字体信息，颜色。

Workbook globals里就记录了整个Workbook公用的东西，比如格式，字符串常量，文档保护等。worksheets，workbook globals都是存储在Workbook Stream里，可以把它们看做是substream。各个substream按一定顺序出现在Workbook Stream里，Globals Substream是必须的，而且至少保护一个Sheet Substream。那么这些Substream是如何区分开的，如何知道一个Substream从文件的哪个位置开始，哪里结束呢？这得从BIFF8描述信息所采用的格式来说。这里说的"信息"，包含很多，比如格式信息，位置信息，文本信息，图片信息，宏代码信息，等等。这么多信息要记录，BIFF8是采用Record的方式来记录的。一个Record描述一个信息。因此有很多种Record。因此Record Type肯定是Record数据结构的组成部分。Record是不定长的，种类非常多，各个种类的内容格式千差万别。

结构包含两部分，Record Header和Record Data。Record Header部分的前2字节是Record Type，然后是2字节描述Record Data的长度。文档给出了几乎所有的Record Type，针对每种Record，描述了Record Data里各字节的含义。Substream的开始和结束，就是采用BOF和EOF这两个Record Type。也就是说在Workbook Stream是Record序列，一个Record接着一个Record，有些BOF的record表示接下来要描述一个Substream了。因此处理程序可以一个Record一个Record的读取，先检查读入的Record的Type，如果不认识，那么就不处理Record data部分，如果认识就解析Record data部分的含义。Excel的版本是不断升级的，补丁是不断更新的，新的版本，新的build就可能有新的Record Type。因此旧版本的Excel打开高版本的xls文件，有些Record就无法识别，但是不影响基本数据的读取。FastExcel就是根据这个原 理，只处理一些和文本内容相关的必需的Record Type，像格式之类的Record就不处理了，因此就Fast了。

当一个Cell的内容为字符串常量时，BIFF8一般把这个字符串常量用SST这种Record记录在Global Substream，然后在Sheet Substream里使用LabelSST这种Record引用这个字符串。LabelSST这种Record的Record Data部分就包含了行号，列号，字符串索引号（index）等信息。查看文档，可知LabelSST类型的Record Header的Identifier是00FD(hex)，size是10，FastExcel 的项目主页也说了它们只实现了如下的Parser来处理一些Record，仅仅是针对文本内容，BIFF8的Record类型可是几百个。这里的 BOFParser就是处理BOF类型的Record，LabelSSTParser处理LabelSST类型的Record。

* BOFParser
* EOFParser
* BoundSheetParser
* SSTParser
* IndexParser
* DimensionParser
* RowParser
* LabelSSTParser
* RKParser
* MulRKParser
* LabelParser
* BoolerrParser
* XFParser
* FormatParser
* DateModeParser
* StyleParser
* MulBlankParser
* NumberParser
* RStringParser

BIFF8里多字节数据，比如4字节整数，2字节整数，Unicode都是采用Little-Endian的存储方式