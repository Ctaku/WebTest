<%@ Language=VBScript %>
<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp"-->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp"-->
<%
IF request.cookies(eChsys)("ManageKEY")<>"super" then
	Show_Err("对不起，您的后台管理权限不够！")
	response.end
else
	if depmanage="1" or (request.cookies(eChsys)("ManagePurview")="99999" and request.cookies(eChsys)("purview")="99999") then
	dim ID
	dim E_DepName
	ID = ChkRequest(Request("ID"),1)	'防注入
	E_DepName=request("E_DepName")
	%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 单位删除</title>
</head>
<body topmargin="0">

<table border="1" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="<%=border%>" align=center>
<form action=E_DepKillOK.asp method=post>
<tr>
	<td width="100%" align="center" height="55" bgcolor="<%=m_top%>"><b>删 除 单 位 确 认</b></td>
</tr>
<input type="hidden" name="id" value="<%=id%>">
<tr>
	<td width="100%" align=center bgcolor="#FFFFFF"> 
		<p>　</p>
		<p>删除单位：[ <%=E_DepName%> ]？<font color=red><br><br>（此操作将无法恢复！）</font></p>
		<p>　</p>
	</td>
</tr>
<tr>
	<td width="100%"align=center bgcolor="<%=m_top%>" height="55" > 
		<input type=submit value="确定" name="alert_button" >&nbsp;&nbsp;&nbsp;&nbsp;
		<input type=button value="取消" name="alert_button" style="font-size: 9pt;  color: #000000; background-color: #FEEBE4; solid #FEEBE4"  onClick="javascript:history.go(-1)" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#FEEBE4'">

	</td>
</tr>
</form>
</table>
</body>
</html>
<!--#include file="Admin_Bottom.asp"-->
	<%else
		Show_Err("对不起，该功能已经被超级系统管理员关闭，您没有权限操作！")
		response.end
	end if
end if%>