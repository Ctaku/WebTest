<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkURL.asp"-->
<%
if not(request.cookies(eChsys)("KEY")="super" or request.cookies(eChsys)("KEY")="check" or request.cookies(eChsys)("KEY")="typemaster" or request.cookies(eChsys)("KEY")="bigmaster") then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	dim ReViewID,passed,passed1
	ReViewID=ChkRequest(request("ReViewID"),1)	'防注入
	passed=request("passed")
	set rs=server.createobject("adodb.recordset")
	sql="Select ReViewID,NewsID,passed from "& db_EC_Review_Table &" where ReViewID="&ReViewID
	rs.open sql,conn,1,3
	if not (rs.bof or rs.eof) then
		if rs("passed")=0 then
			rs("passed")=1
		else
			rs("passed")=0
		end if
		rs.update
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
	
		dim ViewUrl
		ViewUrl=request.cookies(eChsys)("ViewUrl")
		if ViewUrl="" then
			ViewUrl="E_GuestBook.asp"
		end if
		if trim(request("Guest"))=1 then
			Response.Redirect "GuestReView.asp"
		else
			if trim(request("Guest"))=2 then
				Response.Redirect ViewUrl
			else
				Response.Redirect "ReViewManage.asp"
			end if
		end if
	else
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Show_Err("未找到相关记录，操作失败！<br><br><a href=history.go(-1)>返回</a>")
		response.end
	end if
end if%>