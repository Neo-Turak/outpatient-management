Attribute VB_Name = "Module1"
Public Function st() As String
Dim cn As New ADODB.Connection '声明一个连接对象
    Dim rst As New ADODB.Recordset '声明一个记录集对象
    Dim SqlStr As String '声明一个字符串变量
    cn.Open "Provider=SQLOLEDB.1;Password=Q123456;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
    '上面是连接到数据库
    SqlStr = "Select * From Scjd"
    rst.CursorLocation = adUseClient '设置游标位置
    rst.Open SqlStr, cn, adOpenDynamic, adLockOptimistic, adCmdText '打开记录集
    rst.Fields("合作医疗号").Value = Text1.Text '
    rst.Fields("病人姓名").Value = Text2.Text '
    rst.UpdateBatch '提交，就是写到硬盘的数据库文件
    rst.Close '关闭记录集
    Set rst = Nothing '释放
    cn.Close '关闭连接
    Set cn = Nothing '释放
    End Function
Public Function 打印() As String

Printer.FontSize = 16
Printer.Print Space(20)
Printer.Print Space(20)
Printer.Print Space(20)
Printer.CurrentX = 5
Printer.CurrentY = 5
Printer.Print "荒地镇卫生院挂号单"
Printer.FontSize = 12
Printer.Print ""
Printer.Print "合作医疗号：" + Space(5) + Text1.Text
Printer.Print "姓名：" + vbTab + Text2.Text
Printer.Print "身份证号：" + vbtab5 + Text3.Text
End Function
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'*定义一个连接
Public Function 连接() As String


Dim conn As ADODB.Connection
'*定义一个记录集
Dim Mrc As ADODB.Recordset
'*分别实例化
Set conn = New ADODB.Connection
Set Mrc = New ADODB.Recordset
'*定义一个连接字符串
Dim ConnectString As String
ConnectString = "provider=microsoft.jet.oledb.4.0;data source=C：\参合数据库\db2016.mdb;jet oledb:database"
'*打开连接
conn.Open ConnectString
'*定义游标位置
conn.CursorLocation = adUseClient
'*查询记录集(从student表中找出名子为"张三"的记录)
Mrc.Open "select * from 2016 where 身份证号='" & Text1.Text & "'", conn, adOpenKeyset, adLockOptimistic

'*现在你已经得到了你想要查询的记录集了，那就是mrc
'*你可以把此记录集与DataGrid榜定，用datagrid显示你查询的记录
Set Adodc1.DataGrid1.DataSource = Mrc

End Function
