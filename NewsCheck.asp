<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkURL.asp"-->
<%
if not(request.cookies(eChsys)("KEY")="super" or request.cookies(eChsys)("KEY")="check" or request.cookies(eChsys)("key")="typemaster" or request.cookies(eChsys)("key")="bigmaster") then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	newsid=trim(request.QueryString("newsid"))
	set rs=server.createobject("adodb.recordset")
	sql="Select * from "& db_EC_News_Table &" where newsID="&newsid
	rs.open sql,conn,1,3
	E_SmallClassID=rs("E_SmallClassID")
	E_BigClassID=rs("E_BigClassID")
	E_typeid=rs("E_typeid")
	if rs("checkked")=0 then
		rs("checkked")=1
	else
		rs("checkked")=0
	end if
	
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	
	if E_SmallClassID<>"" then
		Response.Redirect "ListNews.asp?E_SmallClassID="&E_SmallClassID
	else if E_BigClassID<>"" then
		Response.Redirect "E_SmallNO.asp?E_BigClassID="&E_BigClassID
		else
			Response.Redirect "E_bigclassNO.asp?E_typeid="&E_typeid
		end if
	end if
end if
%>