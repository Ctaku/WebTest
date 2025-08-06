<%@ LANGUAGE = VBScript %>
<!--#include file="Conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--'#include file="char.inc"-->

<%dim jingyong
set rs=server.createobject("adodb.recordset")
sql="select * from "& db_User_Table &" where  "& db_User_Name &"='"&Request.cookies(eChsys)("username")&"'"
rs.open sql,ConnUser,1,1
if rs.bof or rs.eof then
	response.redirect "login.asp"
	response.end
end if
jingyong=rs("jingyong")
rs.close
set rs=nothing

if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="check" or Request.cookies(eChsys)("KEY")="bigmaster" or Request.cookies(eChsys)("KEY")="smallmaster" or Request.cookies(eChsys)("KEY")="typemaster" or (Request.cookies(eChsys)("key")="selfreg" and jingyong=0) then
	list=request("action")
	%>
<html>
<head>
<title>后台管理</title>
<meta http-equiv=Content-Type content=text/html;charset=gb2312>
<script language="JavaScript">
<!--
if (self != top) top.location.href = window.location.href
//-->
</script>

<style type="text/css">
<!--
.STYLE2 {color: #FF6600}
-->
</style>
</head>
<frameset rows="35,*" frameborder="NO" border="0" framespacing="0">
	<frame src="admin_top.asp" noresize="noresize" frameborder="0" name="topFrame" scrolling="no" marginwidth="0" marginheight="0" />
<frameset rows="*" cols="185,*" id="frame">
	<frame src="admin_left.asp" name="leftFrame" noresize="noresize" marginwidth="0" marginheight="0" frameborder="0" scrolling="yes" />
	<frame src="admin_main.asp" name="main" marginwidth="0" marginheight="0" frameborder="0" scrolling="yes" />
</frameset>
<noframes>
	<body></body>
</noframes>
</frameset>
</html>
<%else
	Show_Err("对不起，管理员尚未开通您的帐号，请等待开通或联系管理员。<br><br><a href='index.asp'>返回首页</a>")
	response.end
end if%>