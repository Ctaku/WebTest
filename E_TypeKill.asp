<%@ Language=VBScript %>
<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<%
IF request.cookies(eChsys)("KEY")<>"super" THEN
	Show_Err("对不起，您的后台管理权限不够！")
	response.end
else
	dim E_typeid
	E_typeid=ChkRequest(request("E_typeid"),1)	'防注入
	Set rs6 = Server.CreateObject("ADODB.Recordset")
	sql6 ="SELECT  * From "& db_EC_Type_Table &" where E_typeid=" & E_typeid & " order by E_typeid"
	RS6.open sql6,Conn,3,3
	%>

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 删除总栏</title>
</head>
<body topmargin="0">

<div align="center">
<table border="1" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="<%=border%>">
<form action=E_TypeKillOK.asp method=post>
<tr class="TDtop">
	<td height=25 class="TDtop" align="center">┊ <B>删除总栏确认</B> ┊ </td>
</tr>
<input type="hidden" name="E_typeid" value="<%=E_typeid%>">
<tr>
	<td width="100%" align=center bgcolor="#FFFFFF"> 
		<p></p>
		<p>删除总栏：[ <%=rs6("E_typename")%> ]？
			<font color=red><br><br>
			（此操作将一起删除相应的大类、小类和文章！并且删除后将无法恢复！）</font>
		</p>
		<p></p>
	</td>
</tr>
<tr>
	<td width="100%" align=center class="TDtop" height="55" > 
		<input type=submit value="  是  " name="alert_button" >&nbsp;&nbsp;&nbsp;&nbsp;
		<input type=submit value="  否  " name="alert_button" >
	</td>
</tr>
</form>
</table>
</div>
</body>
</html>
<!--#include file="Admin_Bottom.asp"-->
	<%rs6.close
	set rs6=nothing
end if
%>