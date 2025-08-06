<%@ Language=VBScript %>
<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkURL.asp"-->
<%
IF not(request.cookies(eChsys)("KEY")="super" or request.cookies(eChsys)("KEY")="typemaster") THEN
	Show_Err("对不起，您的后台管理权限不够！")
	response.end
else
	dim E_BigClassID
	E_BigClassID = ChkRequest(Request("E_BigClassID"),1)	'防注入
	Set rs6 = Server.CreateObject("ADODB.Recordset")
	sql6 ="SELECT * From "& db_EC_BigClass_Table &" where E_BigClassID="&E_BigClassID
	RS6.open sql6,Conn,3,3
	set rs2=server.createobject("adodb.recordset")
	sql2="select * from EC_type where E_typeid="&rs6("E_typeid")
	rs2.open sql2,conn,1,3
	master=rs2("typemaster")
	if master<>"" then
		tmaster=split(master,"|")
		for i=0 to ubound(tmaster)
			if request.cookies(eChsys)("username")=trim(tmaster(i)) then
				typemaster=true
				exit for
			else
				typemaster=false
			end if
	 	next
	end if
	rs2.close
	set rs2=nothing
	if request.cookies(eChsys)("key")="super" or (request.cookies(eChsys)("key")="typemaster" and typemaster=true) then
	%>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 删除大类</title>
</head>
<body topmargin="0">

<div align="center">
<table border="1" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="<%=border%>">
<form target=_self action=E_BigClassKillOK.asp method=post>
<tr class="TDtop">
	<td height=25 class="TDtop" align="center">┊ <B>删除文章大类确认</B> ┊ </td>
</tr>
<input type="hidden" name="E_BigClassID" value="<%=rs6("E_BigClassID")%>">
<tr>
	<td width="100%" align=center bgcolor="#FFFFFF"> 
		<p>　</p>
		<p>删除文章大类：[ <%=rs6("E_BigClassName")%> ]？
			<font color=red><br><br>（此操作将一起删除相应的小类和文章！并且删除后将无法恢复！）</font>
		</p>
		<p>　</p>
	</td>
</tr>
<tr>
	<td width="100%" align=center class="TDtop" height="55"> 
		<input type="hidden" name="delbigclass" value="1">
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
	<%else
		Show_Err("您无权删除！<br><br><a href='javascript:history.back()'>返回</a>")
		response.end
	end if
end if
rs6.close
set rs6=nothing%>