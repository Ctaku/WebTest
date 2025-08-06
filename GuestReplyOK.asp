<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="ChkUser.asp"-->
<!--#include file="ChkURL.asp"-->
<%if request.cookies(eChsys)("key")="super" then
	reviewid=checkstr(request.form("reviewid"))
	guestreply=checkstr(request.form("guestreply"))
	if guestreply="" then
		Show_Err("回复内容为空！<br><br><a href='javascript:history.back()'>返回</a>")
		response.end
	end if
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_EC_Review_Table &" where reviewid="&reviewid
	rs.open sql,conn,2,3
	rs("reply")=guestreply
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.redirect "E_GuestBook.asp"
else
	Show_Err("您无权回复此留言！<br><br><a href='javascript:history.back()'>返回</a>")
	response.end
end if
%>