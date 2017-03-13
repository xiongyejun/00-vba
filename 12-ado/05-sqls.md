
##生成排序##

	str_sql = "Select 项目,时间 From [Sheet1$] Group By 项目,时间"

	str_sql = "Select A.项目,A.时间,(Select Count(*)+1 From (" & str_sql & ") B Where B.项目=A.项目 And B.时间<A.时间) As 排序 From(" & str_sql & ") A"
