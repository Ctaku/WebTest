<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="Chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<%
if request.cookies(eChsys)("ManageKEY")="super" or request.cookies(eChsys)("ManageKEY")="typemaster" or request.cookies(eChsys)("ManageKEY")="bigmaster" then
	ComplainID=CheckStr(request.Form("ComplainID"))
	ReplyContent=CheckStr(request.Form("ReplyContent"))
	
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_EC_Complain_Table &" where ComplainID="&ComplainID
	rs.open sql,conn,2,3
	rs("ReplyTime")=now()
	rs("Replycontent")=ReplyContent
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=ComplainList.asp?ComplainID="&ComplainID&""">"
	Show_Message("<p align=center><font color=red>��ϲ�����Ѿ��ɹ��ظ�!<br><br>"&freetime&"���Ӻ󷵻���ҳ!</font>")
else
	Show_Err("����Ȩ�ظ���<br><br><a href='javascript:history.back()'>����</a>")
	response.end
end if

%>