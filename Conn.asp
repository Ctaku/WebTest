<%
starttime=timer()

DbType = "ACCESS" '链接ACCESS数据库

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
	Response.Write "数据库连接出错[代码：01]，请检查数据库链接文件中的连接字串。"
	Response.End
End If
%>