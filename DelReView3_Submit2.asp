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
	Show_Err("�Բ������ĺ�̨����Ȩ�޲���������")
	response.end
else
	ReViewID=ChkRequest(Request.Form("ReViewID"),1)	'��ע��
	button_value=trim(Request.Form("alert_button"))
	if button_value="ȷ��" then
		conn.execute("delete from "& db_EC_Review_Table &" where ReViewID=" & ReViewID)
	else
		Response.Redirect "GuestReView.asp"
	end if

	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=GuestReView.asp"">"
	Show_Message("���� ID =["& ReViewID &"]ɾ���ɹ���!<br><br>"&freetime&"���Ӻ󷵻���ҳ!</font>")
	response.end
end if
%>