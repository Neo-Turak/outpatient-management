Attribute VB_Name = "Module1"
Public Function st() As String
Dim cn As New ADODB.Connection '����һ�����Ӷ���
    Dim rst As New ADODB.Recordset '����һ����¼������
    Dim SqlStr As String '����һ���ַ�������
    cn.Open "Provider=SQLOLEDB.1;Password=Q123456;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=NURA\SQLEXPRESS"
    '���������ӵ����ݿ�
    SqlStr = "Select * From Scjd"
    rst.CursorLocation = adUseClient '�����α�λ��
    rst.Open SqlStr, cn, adOpenDynamic, adLockOptimistic, adCmdText '�򿪼�¼��
    rst.Fields("����ҽ�ƺ�").Value = Text1.Text '
    rst.Fields("��������").Value = Text2.Text '
    rst.UpdateBatch '�ύ������д��Ӳ�̵����ݿ��ļ�
    rst.Close '�رռ�¼��
    Set rst = Nothing '�ͷ�
    cn.Close '�ر�����
    Set cn = Nothing '�ͷ�
    End Function
Public Function ��ӡ() As String

Printer.FontSize = 16
Printer.Print Space(20)
Printer.Print Space(20)
Printer.Print Space(20)
Printer.CurrentX = 5
Printer.CurrentY = 5
Printer.Print "�ĵ�������Ժ�Һŵ�"
Printer.FontSize = 12
Printer.Print ""
Printer.Print "����ҽ�ƺţ�" + Space(5) + Text1.Text
Printer.Print "������" + vbTab + Text2.Text
Printer.Print "���֤�ţ�" + vbtab5 + Text3.Text
End Function
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'*����һ������
Public Function ����() As String


Dim conn As ADODB.Connection
'*����һ����¼��
Dim Mrc As ADODB.Recordset
'*�ֱ�ʵ����
Set conn = New ADODB.Connection
Set Mrc = New ADODB.Recordset
'*����һ�������ַ���
Dim ConnectString As String
ConnectString = "provider=microsoft.jet.oledb.4.0;data source=C��\�κ����ݿ�\db2016.mdb;jet oledb:database"
'*������
conn.Open ConnectString
'*�����α�λ��
conn.CursorLocation = adUseClient
'*��ѯ��¼��(��student�����ҳ�����Ϊ"����"�ļ�¼)
Mrc.Open "select * from 2016 where ���֤��='" & Text1.Text & "'", conn, adOpenKeyset, adLockOptimistic

'*�������Ѿ��õ�������Ҫ��ѯ�ļ�¼���ˣ��Ǿ���mrc
'*����԰Ѵ˼�¼����DataGrid�񶨣���datagrid��ʾ���ѯ�ļ�¼
Set Adodc1.DataGrid1.DataSource = Mrc

End Function
