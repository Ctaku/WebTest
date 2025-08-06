<%@ Language=VBScript %>
<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="Config.ASP"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="chkManage.asp" -->
<!--#include file="ChkURL.asp"-->
<%
if not(request.cookies(eChsys)("ManageKEY")="super" or request.cookies(eChsys)("ManageKEY")="typemaster" or request.cookies(eChsys)("ManageKEY")="bigmaster") then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	dim SpecialID
	SpecialID = ChkRequest(Request("SpecialID"),1)	'防注入
	set rs=server.CreateObject("ADODB.RecordSet")
	rs.Source="select * from "& db_EC_Special_Table &" where SpecialID=" & SpecialID
	rs.Open rs.Source,conn,1,1
	
	if rs("specialmaster")=request.cookies(eChsys)("ManageUserName") or request.cookies(eChsys)("ManageKEY")="super" or request.cookies(eChsys)("ManageKEY")="check" or request.cookies(eChsys)("ManageKEY")="typemaster" or request.cookies(eChsys)("ManageKEY")="bigmaster" then
	%>

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 专题删除</title>
</head>
<body topmargin="0">
<div align="center">
<table border="1" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="<%=border%>">
<form action=SpecialKillok.asp method=post>
<tr>
	<td width="100%" align="center" height="55" bgcolor="<%=m_top%>"><b>删 除 专 题 确 认</b></td>
</tr>
<input type="hidden" name="Specialid" value="<%=Specialid%>">
<tr>
	<td width="100%" align=center bgcolor="#FFFFFF"> 
	<p>　</p>
	<p>删除专题：[ <%=rs("E_SpecialName")%> ]？<font color=red><br><br>
	一起删除相应的文章
		<%if request.cookies(eChsys)("ManageKEY")="super" then%>
		<input type="radio" name="killnews" value="1">是
		<%end if%>
		<input type="radio" name="killnews" value="0" checked>否</font>
	</p>
	<p></p>
	</td>
</tr>
<tr>
	<td width="100%"align=center bgcolor="<%=m_top%>" height="55" > 
		<input type=submit value="确定" name="alert_button" >&nbsp;&nbsp;&nbsp;&nbsp;
		<input type=button value="取消" name="alert_button" onClick="javascript:history.go(-1)" >
	</td>
</tr>
</form>
</table>
</div>
</body>
</html>
<!--#include file="Admin_Bottom.asp"-->
		<%rs.close
		set rs=nothing
		conn.close
		set conn=nothing
	else
		Show_Err("因为这不是您添加的专题，所以您无权删除该专题！<br><br>－－－<a href='javascript:history.back()'>返回</a>－－－")
		response.end
	end if
end if%>