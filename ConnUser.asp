<%
DbUserType = "ACCESS" '链接ACCESS数据库

DbUserType="ACCESS"
	UserDB = "DataBases/###EC0703.V6013#.asp"
	ConnUserStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(UserDB)

On Error Resume Next
Set ConnUser = Server.CreateObject("ADODB.Connection")
ConnUser.open ConnUserStr

If Err Then
	err.Clear
	Set ConnUser = Nothing
	Response.Write "数据库连接出错[代码：02]，请检查数据库链接文件中的连接字串。"
	Response.End
End If
%>