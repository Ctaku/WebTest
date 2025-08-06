<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="Chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<%
if request.cookies(eChsys)("ManageKEY")="super" or request.cookies(eChsys)("ManageKEY")="typemaster" or request.cookies(eChsys)("ManageKEY")="bigmaster" then
	OpinionID=CheckStr(request.Form("OpinionID"))
	ReplyContent=CheckStr(request.Form("ReplyContent"))
	
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_EC_Opinion_Table &" where OpinionID="&OpinionID
	rs.open sql,conn,2,3
	rs("ReplyTime")=now()
	rs("Replycontent")=ReplyContent
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=E_OpinionList.asp?OpinionID="&OpinionID&""">"
	Show_Message("<p align=center><font color=red>恭喜您，已经成功回复!<br><br>"&freetime&"秒钟后返回上页!</font>")
else
	Show_Err("您无权回复此发言！<br><br><a href='javascript:history.back()'>返回</a>")
	response.end
end if

%>