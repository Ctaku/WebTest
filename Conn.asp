<%
starttime=timer()

DbType = "ACCESS" '����ACCESS���ݿ�

 DbType="ACCESS" 
	DB = "DataBases/###EC0703.V6013#.asp"
	ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(db)

On Error Resume Next
Set conn = Server.CreateObject("ADODB.Connection")
conn.open ConnStr

If Err Then
	Response.Write Err.Description
	
	err.Clear
	Set Conn = Nothing
	Response.Write "���ݿ����ӳ���[���룺01]���������ݿ������ļ��е������ִ���"
	Response.End
End If
%>