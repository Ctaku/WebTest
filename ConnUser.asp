<%
DbUserType = "ACCESS" '����ACCESS���ݿ�

DbUserType="ACCESS"
	UserDB = "DataBases/###EC0703.V6013#.asp"
	ConnUserStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(UserDB)

On Error Resume Next
Set ConnUser = Server.CreateObject("ADODB.Connection")
ConnUser.open ConnUserStr

If Err Then
	err.Clear
	Set ConnUser = Nothing
	Response.Write "���ݿ����ӳ���[���룺02]���������ݿ������ļ��е������ִ���"
	Response.End
End If
%>