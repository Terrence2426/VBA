' Author: Chen Tian
' Date: 2020/10/12
' Based on ISBN-978-7-301-27202-2


' 快捷键
'   打开 VBA 编辑器：Alt + F11
'   运行 VBA 脚本： F5


' 数据类型
'   分为文本、数值、日期值（特殊数值）、逻辑值、错误值五类
'     Integer   整数型   %
'     Currency  货币型   @
'     Date      日期型   
'     String    字符串   $

' 对象类型
'

' 声明变量
'   作用域：Dim 变量 / Private 私有变量 / Public 公有变量 / Static 静态变量
'   Dim/Private/Public/Static 变量名 As 变量类型

    Dim sht As Worksheet
    Set sht = ActiveSheet ' 活动工作表


' 变量赋值
'		Let 变量名 = 变量值

		Dim IntCount As Interger
		Let IntCount = 3000

