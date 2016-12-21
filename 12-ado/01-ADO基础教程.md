
[enter link description here](http://club.excelhome.net/forum.php?mod=viewthread&tid=540915&extra=page=1)

ActiveX Data Object(ADO)基础教程（一） by qee用



#ADO的概念#

这一部分是不得不讲，却又很难讲清的部分，当你以后能熟练使用ADO的时候，你可能会把这些“概念”全部忘掉了，但如果你从未了解过ADO的这些概念，它会始终困扰你，甚至影响你继续学习的信心。

但是要想完全真正理解这些概念，对我们几乎是不可能的。我的理论水平也非常有限，下面只就ADO涉及的最常用的概念给出一些尽可能“易于理解”的说明，首先声明：这些概念不是官方的严谨叙述，更象是“演义”，目的只是让初学的朋友知道“是那么回事”或者产生一个“朦朦胧胧的印象”，如果有对ADO相关理论感性趣的朋友，请参考MSDN（Microsoft Developer Network）的文档。

闲言少叙。

##概念1：什么是ADO？ActiveX Data Objects：ActiveX 数据对象
我从未见过有人给出“ActiveX”的汉语翻译，不过仅从后面的两个英文单词，我们已经可以知道ADO是一种数据对象。

数据对象嘛，其作用就是用来管理数据的。当然管理数据的不一定非得是数据对象，数据对象也不是可以管理所有的数据。（绕口令呵）
对数据的管理我们可以不使用任何对象，而只使用普通的代码来完成；也可以使用数据对象来做，至于选用何种方式，主要取决于哪种方式更适合（有时也取决于写代码者的偏好）！

问题是，怎么知道哪种方式更适合呢，当然你必要要了解各种方式，今天我们要了解的是ADO！
在这个概念中，我已经初步回答了ADO的作用。更多的进一步的回答我放在后面的实战中：-）。

##概念2：什么是ActiveX？
在很早以前，我曾经问过我的一个朋友：ActiveX是什么意思？他回答是：一种商标的名字。
当时我确实注意到，ActiveX后面有一个®（R），我是学经济类专业的，知道®(R)是注册商标的意思。所以在很长时间我不再去追究它的具体含义，商标的名字有什么好研究的。
再后来，无意中看到了些关于ActiveX更多的介绍，现在，我还是觉得我朋友给我的解释最好，大道至简！科学的东西从来都不复杂。
但我还是要给大家介绍一下我所看的相关介绍，先要来了解另一个概念。
##概念3：什么是OLE?
OLE是Object Linking and Embedding，对象链接与嵌入技术
OLE是封装了一些软件（对象）的库文件，这个库文件通常称为“部件”，它有几个特征：
（1）它是可运行代码
（2）它是可被其它外部应用程序调用的代码
（3）外部程序可以重复调用库中的代码，通常称为代码重用
大家可以看出，上面的三个特征都与“类”有关，这就是为什么说“类”是部件的基础的原因。
扯远了，赶紧回来。
那么OLE和ActiveX有什么关系呢？
当发展到网络时代的时候，OLE需要能够与Web浏览器交互，嵌入到网页中，随网页传送到客户的浏览器上，并在客户端执行。这个时候，OLE的基础技术也有了发展，就是我们常听说的COM（Component Object Model，部件对象模型），我们不再去讨论COM了，不然就越说越远了。按照一般的升级命名原则，这时应该叫OLE 2.0，但微软给OLE改名了，它就是ActiveX。
所以可以说，ActiveX其实就是OLE 2.0，或者是支持网页技术的OLE。
大家知道，由于互联网本身具有安全问题，访问速度远低于本地访问速度等一些特殊性，ActiveX部件通常还有如下特征：
（1）一般都提供“代码签名”或要求注册使用，以保证其安全性。
（2）占用内存尽可能小，效率（速度）尽可能高。但这也不是绝对的，随着网速的提升，很多ActiveX部件的制作要求也在下降。
到这儿，大家再统起来看看ActiveX Data Objects，是不是对这几个词有了一个是“朦胧”的印象了~~~
##概念4：什么是关系数据库？
ADO管理的是数据，其实这里的数据通常情况下是“关系数据”，这些“关系数据”的集合称为关系数据库。
何谓“关系”，简而言之，即“表格”。
这样，关系数据库的含义就是由“表格”组成的数据库。
这样解释可能出乎很多朋友的意料，但这个解释肯定错不了。我不再去细说这个“表格”，说的多了，只会让人糊涂。只说一些我们后面有用的：
表格的列一般称为字段，每一列（字段）都具有相同的类型
表格的行一般称为记录。一行称为一条记录。
大家记住一点：当我们打算使用ADO来管理EXCEL数据时，这个数据区域一定要可以被看做“表格”，它的每一列要保证相同的类型，举个例子说，不能有些是日期，而另外一些是文本或数字类型。
关系数据库的概念解释到此为止。
##概念5：什么是SQL？
SQL：Stuctured Query Language 结构化查询语言
ADO管理数据，是通过连接OLE DB驱动来完成的（OLE Database这个词不用解释了吧，大家看名字就知道是干什么活的），真正的数据管理者是OLE DB，管理嘛，当然要使用语言了，OLE DB使用的语言就是SQL。所以，SQL对我们来说，是使用OLE DB的核心，也就成为使用ADO的核心内容，你要发布管理数据的“命令”必须使用SQL语言。不会SQL就无法管理数据，也就谈不上使用ADO。
这里我们知道了ADO和SQL的关系了。
简单介绍SQL的历史。
SQL是关系数据库研究的产物，他是美国的一位博士于上世纪70年代最先提出，80年代美国国家标准局（ANSI）制订发布了SQL美国国家标准，并被国际标准化组织（ISO）所接受。这样，随着SQL标准地位的确定，很多数据库厂家都纷纷采用，SQL也就成了最流行的数据库语言。但各家在采用SQL，都对“标准”SQL进行了扩充和改动，形成了很多“方言”，OLE DB采用的SQL也是方言之一。

其它概念我们将在后面遇到时再讲。
请大家看3遍。以后就可以放下这些概念问题，而把更多的注意力放在ADO的实际应用上。15分钟后，我们进入ADO的 
#实战
二、ADO代码步骤从现在起，我们需要同步互动。
请打开你下载的《模拟数据.xls》，进入VBE，插入一个模块，先写下这样一个框架
Sub Ado0()

End Sub
我们下面以“查询”为例介绍ADO的工作步骤。使用ADO工作共有五个步骤：
步骤1：创建ADO对象。 
我们只介绍最常用两个ADO对象Connection和Recordset，Record（记录）对象表示Recordset（记录集）对象中的一条记录，我们也会提到。
Connection 对象代表打开的、与数据源的连接。
Recordset 对象表示的是来自基本表或命令执行结果的记录全集。
上面的概念来自ADO的帮助文档，现在觉得抽象不要紧，关键是后面学会怎么用它们就行。

是用ADO，必须先创建ADO对象。
创建ADO对象方法1：使用VBA的CreateObject函数。

	Dim cnn As Object, rst As Object
	
	Set cnn = CreateObject("ADODB.Connection")
	Set rst = CreateObject("ADODB.Recordset")
  上面语句为我们创建了两个ADO对象。

创建ADO对象方法2：添加工程引用这个方法首先通过VBE“工具”菜单-引用，在“引用”列表中找到
Microsoft ActiveX Data Objects 2.x Library
勾选后确定。应尽量选择高一点版本。
然后就可以使用下面的代码创建ADO对象：

	Dim cnn As ADODB.Connection
	Set cnn = New ADODB.Connection

也可以在声明是直接创建，上面代码写为：

	Dim cnn As New ADODB.Connection

创建ADO对象的方法使用上面的两种方法之一即可，第二种方法的好处是可以在编辑代码时“自动列出对象成员”，后面的代码我们将采用这种方法。
现在，请在你的Sub Ado0中写入如下代码（后三行我们后面会用到）：

    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    Dim Sql As String
    Dim i As Integer
    Dim j As Integer

步骤2：建立连接创建了ADO对象后，首先要做的就是为Connection对象指定连接的数据源。
ADO建立连接是通过OLE DB进行的，OLE DB的驱动种类有很多，对EXCEL而言，支持的OLE DB连接方式有两种：
ODBC（Open Database Connectivety）开放数据库连接
JET（Joint Engine Technology）连接引擎技术
ODBC是早期的OLE DB驱动，它对系统的底层依赖和限制过多，且以“效率最低”著称，相对而言，JET更为灵活高效，所以我们只介绍JET连接。不是ODBC没用，对早期的某些数据源，因为JET没有提供支持，或者你的机器上没有JET驱动（可能性不大），还是必须使用ODBC的。
建立连接是通过使用Connection对象的Open方法来完成的，在打开连接前，需要先设置Connection对象的Provider和ConnectionString属性。下面代码为前面创建的cnn对象建立连接：

	With cnn
		.Provider = "Microsoft.Jet.OleDb.4.0"
		.C & ThisWorkbook.FullName
		.Open
	End With

上面with语句体中，第一句为cnn指定OLE DB驱动提供者为Microsoft.Jet.OleDb.4.0，下面说明一下第二句：
Extended Properties='Excel 8.0;Hdr=Yes; 表示要连接的数据源是EXCEL文件，Hdr=Yes表示后面对数据源进行查询时，将要查询的“表”区域的第一行做为表头区，即每一列的第一行作为“字段名”，第二行起为数据区。如果Hdr=No，则表示全部为数据区，这时需要“字段名”第一列系统默认为f1,第二列为f2，依此类推。
后面的ThisWorkbook.FullName表示数据源文件的全路径，这里是连接自身文件，如果是其它EXCEL文件，只需要替换一下这儿的ThisWorkbook.FullName即可。
上面连接Hdr=Yes是系统的默认设置，所以一般不需要写出来，ConnectionString属性的设置可以简化为：
.C & ThisWorkbook.FullName
注意如果存在两个以上的“Extended Properties”，等号后面必须用引号引起来，而且上面连接表达式中的分号是不能省略的。
如果不想先设置Connection对象的Provider和ConnectionString属性在打开连接，也可以在使用Open方法打开连接的时候完成这些设置。上面的连接代码可以写成这样：

	cnn.Open "Provider=Microsoft.Jet.OleDb.4.0;Extended Properties=Excel 8.0;Data Source=" & ThisWorkbook.FullName
这种写法看起来更简洁，我们后面将采用这种写法。请将上行代码写到你Sub Ado0中Open 方法 (ADO Connection)打开到数据源的连接。
语法

	connection.Open ConnectionString, UserID, Password, Options
参数
ConnectionString
可选，字符串，包含连接信息。
UserID
可选，字符串，包含建立连接时所使用用户名。
Password
可选，字符串，包含建立连接时所使用密码。
Options
可选，决定该方法是连接是异步还是同步返回。
说明一下：其实ADO对不同数据源如ACCESS，VFP的操作，步骤都是相同的，甚至后面要讲的SQL语句的使用也是相同的，唯一差别就表现在连接的方式上，或者说连接字符串的表达上
步骤3：构造并执行SQL语句，得到结果集 
现在我们先举一个例子：
请将下行代码写到你Sub Ado0中：
Sql = "Select 班级,姓名 From [一年级$]"
上面Sql语句的意思是从“一年级”表中查询（提取）所有记录的班级和姓名两个字段。SQL语法我们下一部分会详细讲。
构造了Sql语句后，就是执行查询，得到结果集。也有两种方法：
方法1：使用Connection对象的Execute方法
Set rst = cnn.Execute(Sql)
通过上面语句，我们就可以执行查询，并将结果保存到rst对象中。
Execute 方法 (ADO Connection)
执行指定的查询、SQL 语句、存储过程或特定提供者的文本等内容。
语法
connection.Execute CommandText, RecordsAffected, Options 
返回值
返回 Recordset 对象引用。
参数
CommandText
字符串，通常为要执行的 SQL 语句、表名。
RecordsAffected
可选，长整型变量，提供者向其返回操作所影响的记录数目。
Options
可选，长整型值，指示提供者应如何计算 CommandText 参数。

后面两个可选参数我们一般用不到去设置，这里不做介绍。

使用Connection对象的Execute方法返回的结果集，始终为只读、仅向前的游标。也无法取得返回结果集合中的记录数。一般在只需将结果一次性写入工作表中（CopyFromRecordset）时使用，它的好处是写法简洁。如果需要处理返回结果的更多操作，应使用下面的方法。
方法2：使用Recordset对象的Open方法
rst.Open Sql, cnn
本句的功能效果同前面的Set rst = cnn.Execute(Sql)一样。

Open 方法 (ADO Recordset)打开游标。
语法
recordset.Open Source, ActiveConnection, CursorType, LockType, Options
参数
Source
可选，变体型，通常为SQL 语句、表名。
ActiveConnection
可选。变体型，一般为有效 Connection 对象变量名。
CursorType
可选，CursorTypeEnum 值，打开 Recordset 时使用游标类型。 
LockType
可选。打开 Recordset 时使用的锁定（并发）类型。
Options
可选，长整型值，用于指示提供者如何计算 Source 参数。

应该说这五个参数都比较有用，但我们最常用的就是前面三个参数，后面两个参数不用管它们就可以了。我只贴一部分帮助内容，大家实际看帮助时，不要被这么多帮助内容吓到，通过实际使用就容易理解它们了。学习ADO和其它的知识一样，也是需要理论和实践交互的过程，实际应用后再回头去看帮助中的一些理论内容，可以理解的更深，提高也会更快。

这里我们遇到了一个词：游标（Cursor）。游标是数据库的组件，在数据库中，对数据的操作我们直观的感觉是对“表”或者记录（集）进行的，但在系统内部记录的留览和更新都是通过游标来进行的。通俗点讲，游标就是“数据指针”。

下面说明一下第三个参数CursorType，游标可以并且一般也需要在打开前确定其类型。打开游标是可以指定的类型有四种：

	AdOpenForwardOnly （默认值）打开仅向前类型游标。 
	AdOpenKeyset 打开键集类型游标。 
	AdOpenDynamic 打开动态类型游标。 
	AdOpenStatic 打开静态类型游标。 

如果需要计算返回记录集的记录数（RecordCount），需要将游标指定为adOpenStatic或adOpenKeyset类型，如果需要对游标进行更新，则需要指定为adOpenKeyset或AdOpenDynamic类型。

请将下行代码写到你Sub Ado0中：
rst.Open Sql, cnn, adOpenKeyset
步骤4：处理查询结果处理查询结果通常是将查询结果写入工作表中或控件（比如Listview）中。
处理1：CopyFromRecordset方法简便处理如果只需要将查询的结果写入工作表中，可以使用Range对象的CopyFromRecordset方法简便处理：
Sheet7.Range("A2").CopyFromRecordset rst
上面A2是要写入工作表区域的左上角单元格。
CopyFromRecordset 方法（Range 对象）
将一个ADO或 DAO Recordset 对象的内容复制到工作表中，复制的起始位置在指定区域的左上角。
句法
Rng.CopyFromRecordset(Data, MaxRows, MaxColumns)
参数
Data：Void 类型，必选。复制到指定区域的 Recordset 对象。
MaxRows：Variant 类型，可选。复制到工作表的记录个数上限。如果省略该参数，将复制 Recordset 对象的所有记录。
MaxColumns：Variant 类型，可选。复制到工作表的字段个数上限。如果省略该参数，将复制 Recordset 对象的所有字段。
处理2：更为细致的处理当查询结果不是写入工作表中，或者虽是写入工作表中但不是按查询结果的方式时。这时需要对更为细致的处理，比如逐条记录、逐个字段进行处理。

	For i = 1 To rst.RecordCount
		For j = 0 To rst.Fields.Count - 1
			Sheet7.Cells(i + 1, j + 1) = rst.Fields(j)
		Next j
		rst.MoveNext
	Next i
大家仔细看一下上面的代码，应该不难理解的。简单地解释一下：
rst.RecordCount 是记录结果集中的记录数，前面我们已经提过。
rst.Fields.Count 是记录结果集中的字段数，Fields是字段集对象，由单个的Field字段组成，表示Recordset对象的列的集合。Fields成员的下标从0开始，0表示第一个字段。
上面代码我们都假定第一行为预先设定好的表头，代码中不再考虑。有时需要将字段名写入表头，请看下面的代码：

	For i = 1 To rst.Fields.Count
		Sheet7.Cells(1, i) = rst.Fields(i - 1).Name
	Next i
请将上面代码写入Sub Ado0中。
 处理3：记录定位（1）Move系列方法上面已经用到了Recordset对象的MoveNext方法。
rst.MoveNext 移动游标到下一记录。在使用MoveNext移动游标时，一般需要通过Recordset对象的EOF属性先进行判断游标是否到了记录尾。当游标到了记录尾时，EOF属性会被设置为True。上面的代码可以这样写：

	i = 1
	Do While Not rst.EOF
		For j = 0 To rst.Fields.Count - 1
			Sheet7.Cells(i + 1, j + 1) = rst.Fields(j)
		Next j
		i = i + 1
	Loop
请将上面代码写入Sub Ado0中。
与EOF对应的是BOF，用来判断游标是否到了记录首。
与MoveNext类似的还有MoveFirst、MoveLast和 MovePrevious 方法，在指定 Recordset 对象中移动到第一个、最后一个或前一个记录并使该记录成为当前记录。
此外，移动记录还可以使用Move方法。
Move 方法（Recordset 对象）
移动Recordset 对象中当前记录的位置。
语法
recordset.Move NumRecords, Start
参数
NumRecords
长整型，指定当前记录位置移动的记录数。
Start
可选，字符串或变体型，指定从哪儿开始移动。也可为下值之一：
AdBookmarkCurrent（0） 默认。从当前记录开始。 
AdBookmarkFirst（1） 从首记录开始。 
AdBookmarkLast（2） 从尾记录开始。 
在Recordset对象中定位游标位置，除了上面的几个Move方法外，常用的还有：
（2）使用Recordset 对象的AbsolutePosition 属性。AbsolutePosition属性可以设置或返回游标当前的记录位置。下面代码将游标当前位置保存在变量c中，然后设置为第10条记录：
c = rst.AbsolutePosition
rst.AbsolutePosition = 10
（3）使用Recordset 对象的Bookmark属性。Bookmark属性可以设置或返回游标当前当前记录的书签。Recordset 对象的每一条记录都有唯一的“书签”值。下面代码先将游标当前位置设置为第10条记录，然后将当前记录的书签保存到变量c中，然后移动到下一条记录（实际使用时一般是进行其它的处理操作），最后在通过设置Bookmark属性将记录定位到原来的第10条记录。

	rst.AbsolutePosition = 10
	c = rst.Bookmark
	rst.MoveNext
	rst.Bookmark = c

与使用AbsolutePosition 属性的区别是，使用Bookmark属性时，往往不知道或不关心记录所处的实际位置。
4）Find方法Find 方法（Recordset 对象）
搜索 Recordset 中满足指定标准的记录。如果满足标准，则记录集位置设置在找到的记录上，否则位置将设置在记录集的末尾。
语法
Find (criteria, SkipRows, searchDirection, start)
参数
criteria
字符串，包含指定用于搜索的列名、比较操作符和值的语句。
SkipRows
可选，长整型值，默认值为零，指定当前行或 start 书签的位移以开始搜索。
searchDirection
可选的 SearchDirectionEnum 值，指定搜索应从当前行还是下一个有效行开始。其值可为 adSearchForward（1）或 adSearchBackward（-1）。搜索是在记录集的开始还是末尾结束由 searchDirection 值决定。
start
可选，变体型书签，用作搜索的开始位置。
下面代码搜索所有记录，将姓陈的同学名单写入Sheet7的第3列：

	i = 2
	rst.MoveFirst
	rst.Find "姓名 Like '陈*'"
	Do While Not rst.EOF
		Sheet7.Cells(i, 3) = rst.Fields("姓名")
		rst.Find "姓名 Like '陈*'", 1, adSearchForward
		i = i + 1
	Loop
请将上面代码写入Sub Ado0中。
步骤5：关闭并释放ADO对象使用ADO完成了全部工作后，应该关闭并释放创建的ADO对象。
请将下面代码写到你Sub Ado0中：

	rst.Close
	cnn.Close
	Set rst = Nothing
	Set cnn = Nothing

至此，我们完成了一个实例，也介绍完了ADO代码的全部步骤。大家休息10分钟。

上面的代码并没有完全写进我们的Sub Ado0中，大家可以自己试验一下运行结果。
构建SQL语句
从上面我们对ADO工作步骤的了解，已经知道要让ADO有效工作，关键是我们给它发出什么样的SQL指令。
在概念部分，我们已经简单介绍了SQL的有关情况。现在我们来详细探讨它。
SQL语句从功能上可以分为两大类：数据定义语言（DDL）和数据操纵语言（DML）。前者主要用于对数据库中表及字段，还有我们没有提到的索引的创建、删除、修改；后者用于对记录的查询、更新、插入、删除等操作。就EXCEL而言，我们通常使用的是DML部分语句。下面将对常用的语句进行介绍。
简单查询句法1：
Select 查询表达式 From 数据区域前面我们使用的SQL语句就属于此类。
查询表达式请粘贴下面的过程：
	
	Sub Ado1()
	Dim cnn As New ADODB.Connection
	Dim Sql As String
	
	cnn.Open "Provider=Microsoft.Jet.OleDb.4.0;Extended Properties=Excel 8.0;Data Source=" & ThisWorkbook.FullName
	
	Sql = "Select * from [一年级$]"
	Sheet7.Cells.Clear
	Sheet7.[a2].CopyFromRecordset cnn.Execute(Sql)
	cnn.Close
	Set cnn = Nothing
	End Sub
查询表达式可以是下列之一或其组合，对多种方式的组合，用逗号搁开：
（1）星号（*） 表示“数据区域”的所有字段。
（2）字段名
（3）常量表达式
（4）任何有效的计算表达方式

下面是一些SQL语句，请分别替换Sub Ado1的Sql句，并查看运行结果。

	Sql = "Select '一年级',* from [一年级$]"
	Sql = "Select 姓名,语文+数学+英语 from [一年级$]"
	Sql = "Select 姓名,iif(语文>=60,'及格','不及格') from [一年级$]" 

使用AS重新命名列名称当查询表达式使用（2）字段名时，字段名就是其本身，使用（3）常量表达式和（4）任何有效的计算表达方式时，系统将为该字段重新命名一个字段名，这个字段名通常没有意义，这时可以在表达式中使用AS为字段重新命名，当然对字段名也可以通过使用AS为其重新命名。使用AS通常在需要使用字段名的场合（在对HDR=NO的EXCEL数据源更为常见），如我们前面提过的将字段名写入第一行，也可用在多表查询时简化构造语句或者因特殊处理需要。后面我们或许会看到有关的例子。AS并不对查询结果造成实质影响。下面是使用AS的一个例子：
Sql = "Select 班级,姓名 AS 名字,语文+数学+英语 AS 总成绩 from [一年级$]"
数据区域请粘贴下面的过程：
	
	Sub Ado2()
	Dim cnn As New ADODB.Connection
	Dim Sql As String
	
	cnn.Open "Provider=Microsoft.Jet.OleDb.4.0;Extended Properties=Excel 8.0;Data Source=" & ThisWorkbook.FullName
	
	Sql = "Select * from [一班$]"
	Sheet7.Cells.Clear
	Sheet7.[a2].CopyFromRecordset cnn.Execute(Sql)
	cnn.Close
	
	Set cnn = Nothing
	End Sub
数据区域可以是下列之一：
（1）当要查询的数据区域是从工作表的第一行、第一列开始的整个表格时，可以使用[工作表名$]的形式
（2）如果不是（1）的情形，则需要使用[工作表名$区域范围]的形式'A:C列
Sql = "Select * from [一班$A:C]"
'《不规范表》的A2:H19

Sql = "Select * from [不规范表$A2:H19]"
上面两中方式中的方括号和美元符号不能省略。
（3）如果工作表中定义了名称，则可以直接使用名称。
'《不规范表》的A2:H19已经定义名称为DATA
Sql = "Select * from DATA"
（4）数据区域是多个区域的情况我们后面再讲
使用DISTINCT删除重复记录《不规范表》的K2:R24区域有重复的记录，如果希望重复的记录只显示一条，可以使用DISTINCT进行限定。
Sql = "Select distinct * from [不规范表$K2:R24]"
使用Top限制返回行数如果记录返回的行数比较多，可以使用Top限制返回的行数，通常和后面介绍的Order by排续配合使用。
下面语句返回前20条记录。

Sql = "Select top 20 * from [一班$]"
下面语句返回全部符合条件记录的1%。
Sql = "Select top 1 percent * from [一班$]"
句法2：Select 查询表达式 From 数据区域 Where 条件表达式在句法1的基础上，通过使用Where可以设置查询条件。
请粘贴下面的过程：
	
	Sub Ado3()
	Dim cnn As New ADODB.Connection
	Dim Sql As String
	
	cnn.Open "Provider=Microsoft.Jet.OleDb.4.0;Extended Properties=Excel 8.0;Data Source=" & ThisWorkbook.FullName
	
	Sql = "Select * from [一班$] where 性别='男'"
	Sheet7.Cells.Clear
	Sheet7.[a2].CopyFromRecordset cnn.Execute(Sql)
	cnn.Close
	
	Set cnn = Nothing
	End Sub
查询的条件表达式可以是：
（1）任何逻辑表达式
'语文+数学成绩大于120的男生
Sql = "Select * from [一班$] where 性别='男'and 语文+数学>120"
'语文或数学成绩大于80
Sql = "Select * from [一班$] where 语文>80 or 数学>80"
（2）IN/NOT IN ( 表达式1,表达式2,…. )
注意上面的括号不可少，各表达式用逗号搁开。
'查询姓名在括号中列出名单范围内的人
Sql = "Select * from [一班$] where 姓名 in ('梁少娟','袁泳霞')"
将上面的IN 换为 NOT IN查询范围正好相反。
IN/NOT IN的另一种表达后面再讲。 
句法3：Select 查询表达式 From 数据区域 [Where 条件表达式] Order by 排序字段通过使用Order by可以对查询结果按一列或多列进行排序。
请粘贴下面的过程：
	
	Sub Ado4()
	Dim cnn As New ADODB.Connection
	Dim Sql As String
	
	cnn.Open "Provider=Microsoft.Jet.OleDb.4.0;Extended Properties=Excel 8.0;Data Source=" & ThisWorkbook.FullName
	
	Sql = "Select * from [一班$] Order by 语文"
	Sheet7.Cells.Clear
	Sheet7.[a2].CopyFromRecordset cnn.Execute(Sql)
	cnn.Close
	
	Set cnn = Nothing
	End Sub
再看几个例子：
'首先按语文成绩降序排列,语文成绩相同的按数学成绩升序排列
Sql = "Select * from [一班$] Order by 语文 desc,数学 asc"

'按三科成绩之和排序

Sql = "Select * from [一班$] Order by 语文+数学+英语 "
ASC是升序排列，在不指定排序方式的情况下是默认的，因此可以省略。
（二）统计查询句法4：Select 查询表达式 From 数据区域 [Where 条件表达式] Group by 分组表达式[Order by  
排序字段]请粘贴下面的过程：
	
	Sub Ado5()
	
	Dim cnn As New ADODB.Connection
	
	Dim Sql As String
	
	cnn.Open "Provider=Microsoft.Jet.OleDb.4.0;Extended Properties=Excel 8.0;Data Source=" &  
	ThisWorkbook.FullName
	'统计各班的人数和平均分
	
	Sql = "Select 班级,count(*),avg(语文) from [一年级$] group by 班级"
	
	Sheet7.Cells.Clear
	
	Sheet7.[a2].CopyFromRecordset cnn.Execute(Sql)
	
	cnn.Close
	
	Set cnn = Nothing
	End Sub
上面的统计查询语句使用了Group by子句。在统计查询中，查询表达式中没有使用聚合函数的字段必须在分组表达式 
中出现，但分组表达式中的字段可以不在查询表达式中出现。以上面的查询语句：
Select 班级,count(*),avg(语文) from [一年级$] group by 班级
在查询表达式中，“班级”不是以聚合函数的方式出现的，则Group by后面的分组表达式必须出现它。下面语句是错 
误的，因为分组字段没有“性别”：
Select 班级,性别,count(*),avg(语文) from [一年级$] group by 班级
但下面语句是允许的：
Select count(*),avg(语文) from [一年级$] group by 班级
Select 班级,性别,count(*),avg(语文) from [一年级$] group by 班级,性别
聚合函数聚合函数常与group by子句一起使用，用于对查询结果集合中的多个值进行统计计算，并返回单个计算结果 
，但聚合函数不能用在Where子句中。在聚合函数中允许使用DISTINCT关键字。常用的聚合函数有：
COUNT
统计项数，可以是任何类型的表达式，允许使用星号表达
MIN 最小值，可用于数字型、文本型、日期型
MAX最大值，可用于数字型、文本型、日期型

SUM 和值，可用于数字型
AVG平均值，可用于数字型
句法5：Select 查询表达式 From 数据区域 [Where 条件表达式] Group by 分组表达式Having [Order by 排序字段 
]在使用group by子句时，可用Having子句为分组统计进一步设置统计条件。
Having子句和Group by子句的关系相当于Where子句和Select子句的关系。

Sql = "Select 班级,count(*),avg(语文) from [一年级$] group by 班级 having avg(语文)>75"
上面语句只统计语文平均分大于75分的班级。
