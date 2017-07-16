#从未打开的工作簿中读取数据#

	ExecuteExcel4Macro(arg)
	arg = "'" & path & "[" & file & "]" & sheet & "'!" & Range(ref).Range("A1").Address(, , xlR1C1)

说明: 如果工作簿处于隐藏状态,或者工作表是图表工作表,将会报错


# 返回当前窗口左边界到活动单元格的左边之间的水平距离#
	GET.CELL(42)


GET.CELL

GET是得到的意思CELL是单元格的意思 

那么它的意思就是你想得到单元格的什么东西(信息) 

函数定义: 
	GET.CELL(类型号,单元格(或范围)) 

其中类型号,即你想要得到的信息的类型号,可以在1-66(表示可以返回一个单元格里66种你要的信息)

	1        参照储存格的绝对地址
	2        参照储存格的列号
	3        参照储存格的栏号
	4        类似 TYPE 函数
	5        参照地址的内容
	6        文字显示参照位址的公式
	7        参照位址的格式，文字显示
	8        文字显示参照位址的格式
	9        传回储存格外框左方样式，数字显示
	10        传回储存格外框右方样式，数字显示
	11        传回储存格外框方上样式，数字显示
	12        传回储存格外框方下样式，数字显示
	13        传回内部图样，数字显示
	14        如果储存格被设定 locked传回 True
	15        如果公式处于隐藏状态传回 True
	16        传回储存格宽度
	17        以点为单位传回储存格高度
	18        字型名称
	19        以点为单位元传回字号
	20        如果储存格所有或第一个字符为加粗传回 True
	21        如果储存格所有或第一个字符为斜体传回 True
	22        如果储存格所有或第一个字符为单底线传回True
	23        如果储存格所有或第一个字符字型中间加了一条水平线传回 True
	24        传回储存格第一个字符色彩数字， 1 至 56。如果设定为自动，传回 0
	25        MS Excel不支持大纲格式
	26        MS Excel不支持阴影格式
	27        数字显示手动插入的分页线设定
	28        大纲的列层次
	29        大纲的栏层次
	30        如果范围为大纲的摘要列则为 True
	31        如果范围为大纲的摘要栏则为 True
	32        显示活页簿和工作表名称
	33        如果储存格格式为多行文字则为 True
	34        传回储存格外框左方色彩，数字显示。如果设定为自动，传回 0
	35        传回储存格外框右方色彩，数字显示。如果设定为自动，传回 0
	36        传回储存格外框上方色彩，数字显示。如果设定为自动，传回 0
	37        传回储存格外框下方色彩，数字显示。如果设定为自动，传回 0
	38        传回储存格前景阴影色彩，数字显示。如果设定为自动，传回 0
	39        传回储存格背影阴影色彩，数字显示。如果设定为自动，传回 0
	40        文字显示储存格样式
	41        传回参照地址的原始公式
	42        以点为单位传回使用中窗口左方至储存格左方水平距离
	43        以点为单位传回使用中窗口上方至储存格上方垂直距离
	44        以点为单位传回使用中窗口左方至储存格右方水平距离
	45        以点为单位传回使用中窗口上方至储存格下方垂直距离
	46        如果储存格有插入批注传回 True
	47        如果储存格有插入声音提示传回 True
	48        如果储存格有插入公式传回 True
	49        如果储存格是数组公式的范围传回 True
	50        传回储存格垂直对齐，数字显示
	51        传回储存格垂直方向，数字显示
	52        传回储存格前缀字符
	53        文字显示传回储存格显示内容
	54        传回储存格数据透视表名称
	55        传回储存格在数据透视表的位置
	56        枢纽分析
	57        如果储存格所有或第一个字符为上标传回True
	58        文字显示传回储存格所有或第一个字符字型样式
	59        传回储存格底线样式，数字显示
	60        如果储存格所有或第一个字符为下标传回True
	61        枢纽分析
	62        显示活页簿和工作表名称
	63        传回储存格的填满色彩
	64        传回图样前景色彩
	65        枢纽分析
	66        显示活页簿名称


#GET.DOCUMENT(type_num, name_text)#

Type_num    指明信息类型的数。

下表列出 type_num 的可能值与对应结果。

	Type_num        返回
	
	1        如果工作簿中不只一张表，用文字形式以“［book1］sheet1”的格式返回工作表的文件名。否则，只返回工作簿的文件名。工作簿文件名不包括驱动器，目录或窗口编号。通常最好使用 GET. DOCUMENT(76) 
	和 GET. DOCUMENT(88) 来返回活动工作表和活动工作簿的文件名。
	2        作为文字，包括 name_text 的目录的路经。如果工作簿name_text 未被保存，返回错误值 #N/A
	3        指明文件类型的数
			1 = 工作表
			2 = 图表 
			3 = 宏表 
			4 = 活动的信息窗口
			5 = 保留文件
			6 = 模块表
			7 = 对话框编辑表
	
	4        如果最后一次存储文件后表发生了变化，返回TRUE；否则，返回FALSE。
	5        如果表为只读，返回TRUE；否则，返回FALSE。
	6        如果表设置了口令加以保护，返回TRUE；否则， 返回FALSE。
	7        如果表中的单元格，表中的内容或图表中的系列被保护，返回TRUE；否则，返回FALSE。
	8        如果工作簿窗口被保护，返回TRUE；否则，返回FALSE。
	
	下面四个 type_num 的数值只用于图表。
	
	Type_num        返回
	
	9        指示主图表的类型的数。
			1 = 面积图
			2 = 条形图
			3 = 柱形图
			4 = 折线形
			5 = 饼形
			6 = XY (散点图)
			7 = 三维面积图
			8 =三维柱形图
			9 = 三维折线图
			10 = 三维饼图
			11 = 雷达图
			12 = 三维等形图
			13 = 三维曲面图?
			14 = 圆环图
	10        指示覆盖图表类型的数，同以上主图表的 1，2，3，4，5，6,11 和 14。没有覆盖图表的情况下返回错误值 #N/A?
	
	11        主图表系列的数
	12        覆盖图表系列的数
	
	下列 Type_num 的值用于工作表，宏表，在适当的时候用于图表。
	
	Type_num        返回
	
	9        第一个使用行的编号。如文件是空的，返回零。
	10        最后一个使用行的偏号。如文件是空的，返回零。
	11        第一个使用列的编号。如文件是空的，返回零。
	12        最后一个使用列的编号。如文件是空的 ，返回零。
	13        窗口的编号。
	14        指明计算方式的数。
			1 = 自动生成?有
			2 = 除表格外自动生成
			3 = 手动
	15        如果在［选项］对话框的［重新计算设置］标签下选择［迭代］选择框，返回TRUE；否则，返回FALSE。
	
	16        迭代间的最大数值。
	17        迭代间的最大改变
	18        如果在［选项］对话框的［重新计算设置］标签下选择［更新过程引用］选择框，返回TRUE；否则，返回
	FALSE。
	19        如果在［选项］对话框的［重新计算设置］标签下选［以显示值为准］选择框，返回TRUE；否则，返回
	FALSE。
	20        如果在 Options 对话框的［重新计算设置］标签下选择［1904 日期系统选择框，返回TRUE；否则，返回
	FALSE。
	
	Type_num 是21-29之间的数， 对应于 Microsoft Excel 先前版本的四种默认字体。提供这些值是为了宏的兼容性。
	下列 Type_num 数值应用于工作表，宏表和指定的图表。
	
	Type_num        返回
	
	30        以文字形式返回当前表合并引用的水平数组. 如果列表是空的，返回错误值 #N/A
	31        1至11 之间的一个数，指明用于当前合并的函数。对应于每个数的函数列于下面 CONSOLIDATE 函数中，默认函数为SUM
	32        三项水平数组，用于指明 Data Consolidate 对话框中选择框的状态。如果此项为TRUE，选择选择框. 
	如果此项为FALSE，清除选择框. 第一项指明［顶端行］选择框，第二项指［最左列］选择框,第三项指［与源数据链接］选择框。
	
	33        如果选择了［选项］对话框的［重新计算设置］标签下的［保存前重新计算］选择框，返回TRUE；否则，返回FALSE。
	34        如工作簿定义为只读，返回TRUE；否则，返回FALSE。
	35        工作簿为写保护，返回TRUE；否则，返回FALSE。
	36        如文件设置了写保护口令，并以可读/可写方式打开，返回最初使用写保护口令存文件的用户的名字。如文件以只读形式打开，或文件未设置口令，返回当前用户的名字。
	37         对应于显示在［另存为］对话框中的文档的文件类型。所有  Microsoft Excel 可识别的文件类型列于
	SAVE.AS函数中。
	
	38         如选择了［分级显示］对话框中的［明细数据的下方选择框，返回TRUE；否则，返回FALSE。
	39        如果选择了［分级显示］对话框中的［明细数据的右侧］选择框，返回TRUE；否则，返回FALSE。
	40        如果选择了［另存为］对话框中的［建立备份文件］选择框，返回TRUE；否则，返回FALSE。
	41        1至3中的一个数字，指明是否显示对象：
			1 = 显示所有对象
			2 = 图和表的位置标志符
			3 = 所有对象被隐藏
	
	42        包括表中所有对象的水平数组，如无对象，返回错误值 #N/A
	43        如果在［选项］对话框的［重新计算设置］标签下选择了［保存外部链接值］选择框，返回TRUE；否则，返回FALSE。
	44        如文件中的对象被保护，返回TRUE；否则，返回FALSE。
	45        0至3中的一个数，指明窗口同步化方式。
	0 = 不同步
	1 = 水平方向上同步
	2 = 垂直方向上同步
	3 = 水平方向，垂直方向上均同步
	46        七项水平数组，用于打印设置，可由 LINE. PRINT 宏函数完成。
	
			-        建立文字
	        -        左边距
	        -        右边距
	        -        顶边距
	        -        底边距
	        -        页长
	        -        用于指明打印时输出是否格式化的逻辑值，格式化为TRUE,                        非格式化为FALSE。
	47        如果在［选项］对话框的［转换］标签中选择了［转换表达式求值］选择框，返回TRUE；否则，返回FALSE
	。
	48        标准栏宽度设置
	
	下列 type_num 值对应于打印与页的设置。
	
	Type_num        返回
	
	49        开始页的页码，如未指明或在［页面设置］对话框的［页］标签下的［起始页号］文字框输入了“自动”，返回错误值#N/A
	50        当前设置下欲打印的总页数，其中包括注释，如果文件为图表，值为1
	51        如只打印注释时的总页数。如文件为图表类型，返回错误值 #N/A?
	52        在当前指定的单位中，指明边距设置(左，右，顶，底)的四项水平数组。
	53        指明方向的数字:
				1 = 纵向
				2 = 横向
	54        文本串的页眉，包括格式化代码。
	
	55        文本串的脚注，包括格式化代码。
	56        包括两个逻辑值的水平数组，对应于水平垂直方向置中。
	57        如打印行或列的上标题，返回TRUE；否则，返回FALSE。
	58        如打印网格线，返回TRUE；否则，返回FALSE。
	59        如表以黑白方式打印，返回TRUE；否则，返回FALSE。
	60        1至3中的一个数,指明打印时定义图表大小的方式。
			1 = 屏幕大小
			2 = 调整到
			3 = 使用整页
	61        指明重排页命令的数:
			1 = 先列后行
			2 = 先行后列
			如文件为图表类型,返回错误值#N/A
	
	62        扩缩比,未指定时为100%。如当前打印机不支持此项操作或文件为图表类型时，返回错误值#N/A。
	63        一个两项水平数组,指明其报表需按比例换算，以适合的页数印出 ,第一项等于宽度(如未指明宽度按比例缩放,返回#N/A)第二项等于高度(如未指明高度按比例缩放,返回#N/A)。如文件为图表类型,返回#N/A
	64        行数的数组,相应于手动或自动生成页中断下面的行。
	65         列数的数组。相应于手动或自动生成的页中断右边的列。
	
	附注        GET.DOCUMENT(62)和GET.DOCUMENT(63)互相排斥,如果其中一个返回一个数值,另外一个返回错误值#N/A。
	
	下列type_num数值对应不同文件设置。
	
	Type_num        返回
	
	66        Microsoft Excel for  Windows 中，如果在［选项］对话框的［转换］标签中选择了［转换公式项］选择框,返回TRUE;否则,返回FALSE。
	67        Microsoft Excel 5.0版本下,通常返回TRUE。
	68        Microsoft Excel 5.0版本下,通常返回簿的文件名。
	69        如果在［选项］对话框的［查看］标志中选择了［自动分页线］，返回TRUE;否则,返回FALSE。
	70        返回文件中所有数据透视表的文件名
	71         返回表示文件中所有类型的水平数组。
	
	72        返回表示当前表显示的所有图表类型的水平数组。
	73        返回表示当前工作表每一个图表中系列数的水平数组。
	74        返回控制的对象标识符，控制当前执行中的由用户定义的对话框编辑表中获得焦点的控制(以对话框编辑表为基础)。
	75         返回对象的对象标识符,对象正在执行中的由用户定义的对话框编辑表中的默认按枢(以对话框编辑表为基础)。
	76        以［Book1］sheel的形式返回活动表或宏表的文件名。
	77         以整数的形式返回页的大小: 
	
				1=Letter 8.5x11 in
		        2 = Letter Small 8.5 x 11 in
				5 = Legal 8.5 x 14 in
				9 = A4 210 x 297 mm
				10 = A4 Small 210 x 297 mm
				13 = B5 182 x 257 mm
				18 = Note 8.5 x 11 in
	78         返回打印分辨率,为一个二项水平数组。
	79        如在［页面设置］对话框的［工作表］标签中选择［草稿质量］选择框返回TRUE;否则,返回FALSE。
	80        如在［页面设置］对话框的［工作表］标签下选择了［附注］选择框，返回TRUE;否则,返回FALSE。
	
	81        做为一个单元格的引用,从［页面设置］对话框的［工作表］标签返回打印区域。
	82        做为一个单元格引用从［页面设置］对话框的［工作表］标签回打印标题。
	83        如果工作表为方案而被保护起来,返回TRUE;否则,返回FALSE。
	84        返回表中第一个循环引用的值,如无循环引用,返回错误值#N/A。
	85        返回表的高级筛选方式状态。这种方式顶部设有向下的箭头,如数据精单通过选择［筛选］,再从［数据］菜单选择［高级筛选］被筛选,返回TRUE;否则,返回FALSE。
	
	86        返回表的自动筛选方式状态。这种方式顶部有向下的箭头,如选择了［筛选］,再从［数据］菜单选择［自动筛选］,筛选向下的箭头被显示出来,返回TRUE;否则,返回FALSE。
	87        返回指示表的位置的数字,第一张表位置为1。计算中包含隐藏起来的表。
	88        以“book1”的形式返回活动工作簿的文件名。