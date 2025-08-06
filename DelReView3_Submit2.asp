<%@ Language=VBScript %>
<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp"-->
<%
if not(request.cookies(eChsys)("ManageKEY")="super" or request.cookies(eChsys)("ManageKEY")="typemaster" or request.cookies(eChsys)("ManageKEY")="bigmaster") then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	ReViewID=ChkRequest(Request.Form("ReViewID"),1)	'防注入
	button_value=trim(Request.Form("alert_button"))
	if button_value="确定" then
		conn.execute("delete from "& db_EC_Review_Table &" where ReViewID=" & ReViewID)
	else
		Response.Redirect "GuestReView.asp"
	end if

	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=GuestReView.asp"">"
	Show_Message("留言 ID =["& ReViewID &"]删除成功！!<br><br>"&freetime&"秒钟后返回上页!</font>")
	response.end
end if
%>