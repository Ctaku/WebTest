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
	if mailmanage="1" or (request.cookies(eChsys)("ManagePurview")="99999" and request.cookies(eChsys)("purview")="99999") then

	LeadMailID=ChkRequest(Request.Form("LeadMailID"),1)	'防注入
	button_value=trim(Request.Form("alert_button"))
	if button_value="确定" then
		conn.execute("delete from "& db_EC_LeadMail_Table &" where LeadMailID=" & LeadMailID)
	else
		Response.Redirect "E_LeadMailList.asp"
	end if

	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=E_LeadMailList.asp"">"
	Show_Message("信件 ID =["& LeadMailID &"]删除成功！!<br><br>"&freetime&"秒钟后返回上页!</font>")
	response.end
	else
		Show_Err("对不起，该功能已经被超级系统管理员关闭，您没有权限操作！")
		response.end
	end if
end if
%>