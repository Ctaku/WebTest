<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<%
if not(request.cookies(eChsys)("ManageKEY")="super" or request.cookies(eChsys)("ManageKEY")="typemaster" or request.cookies(eChsys)("ManageKEY")="bigmaster")  THEN
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	if request.cookies(eChsys)("ManagePurview")="99999" and request.cookies(eChsys)("purview")="99999" then
		dim rs,tsql
		dim rst
		dim id
		id=ChkRequest(request("id"),1)	'防注入

		set rst=server.CreateObject("ADODB.RecordSet")
		if request("name")="del" then
			rst.open "delete from "& db_EC_Manage_Table &" where "& db_ManageUser_ID &"="&id,Conn,3,3
		end if
		conn.close
		set conn=nothing
	else
		Show_Err("对不起，您的管理权限不够！")
		response.end
	end if
	response.redirect "AdminManage.asp"
end if%>