<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->

<%
if not(request.cookies(eChsys)("KEY")="super" or request.cookies(eChsys)("key")="bigmaster" or request.cookies(eChsys)("key")="check" or request.cookies(eChsys)("key")="typemaster") then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	newsID=trim(request.QueryString("newsid"))
	set rs=server.createobject("adodb.recordset")
	sql="Select * from "& db_EC_News_Table &" where newsID="&newsid
	rs.open sql,conn,1,3
	E_SmallClassID=rs("E_SmallClassID")
	E_BigClassID=rs("E_BigClassID")
	E_typeid=rs("E_typeid")
	if rs("goodnews")=0 then
		rs("goodnews")="1"
	else
		rs("goodnews")="0"
	end if
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	
Response.Redirect "CheckNews1.asp?"
	
end if
%>