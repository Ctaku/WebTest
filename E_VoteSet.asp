<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="ChkUser.asp" -->
<!--#include file="ChkManage.asp" -->
<%
IF not(request.cookies(eChsys)("ManageKEY")="super" or request.cookies(eChsys)("ManageKEY")="bigmaster" or request.cookies(eChsys)("ManageKEY")="check" or request.cookies(eChsys)("ManageKEY")="typemaster") THEN
	Show_Err("对不起，您的管理权限不够！")
	response.end
else
	if votemana="1" or (request.cookies(eChsys)("ManagePurview")="99999" and request.cookies(eChsys)("purview")="99999") then
		Checked=ChkRequest(request.form("Checked"),1)	'防注入
		if Checked="" then
			Show_Err("对不起，您必须选择一个投票项！<br><br><a href='javascript:history.back()'>回去重来</a>")
			Response.End 
		end if
		conn.execute "update "& db_EC_Vote_Table &" set IsChecked=0"

		set rs=server.createobject("adodb.recordset")
		sql="Select * from "& db_EC_Vote_Table &" where ID="&Checked
		rs.open sql,conn,1,3
		if not rs.EOF then
			do while not rs.EOF 
				rs("isChecked")=1
			rs.MoveNext 
			loop
		end if
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		response.redirect "E_VoteManage.asp"
	else
		Show_Err("对不起，该功能已经被超级系统管理员关闭，您没有权限操作！")
		response.end
	end if
end if
%>