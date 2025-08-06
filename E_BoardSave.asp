<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="ChkUser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp" -->
<%
if not(request.cookies(eChsys)("ManageKEY")="super" or request.cookies(eChsys)("ManageKEY")="bigmaster" or request.cookies(eChsys)("ManageKEY")="check" or request.cookies(eChsys)("ManageKEY")="typemaster") then
	Show_Err("对不起，您的管理权限不够！")
	response.end
else
	if ggmanage="1" or (request.cookies(eChsys)("ManagePurview")="99999" and request.cookies(eChsys)("purview")="99999") then

	ID=ChkRequest(Request.QueryString("ID"),1)	'防注入
	Title=Request.Form("Title")

	if Title="" then
		Show_Err("请填写文章标题！<br><br><a href='javascript:history.back()'>回去重来</a>")
		response.end
	end if

	if len(title)>100 then
		Show_Err("公告标题过长！<br><br><a href='javascript:history.back()'>回去重来</a>")
		response.end
	end if

	Content=CheckStr(Request.Form("Content"))
	if Content="" then
		Show_Err("请输入公告内容！<br><br><a href='javascript:history.back()'>回去重来</a>")
		response.end
	end if

	upload=CheckStr(Request.Form("upload"))
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_EC_Board_Table
	rs.open sql,conn,1,3
	rs.addnew
	rs("content")=content
	rs("upload")=upload
	rs("title")=title
	rs("dateandtime")=now()
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "E_BoardManage.asp"
	
	else
		Show_Err("对不起，该功能已经被超级系统管理员关闭，您没有权限操作！")
		response.end
	end if
end if
%>