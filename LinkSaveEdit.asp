<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="ChkUser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp"-->
<%
if not(request.cookies(eChsys)("ManageKEY")="super" or request.cookies(eChsys)("ManageKEY")="bigmaster" or request.cookies(eChsys)("ManageKEY")="check" or request.cookies(eChsys)("ManageKEY")="typemaster") then
	Show_Err("�Բ������Ĺ���Ȩ�޲�����")
	response.end
else
	if linkmana="1" or (request.cookies(eChsys)("ManagePurview")="99999" and request.cookies(eChsys)("purview")="99999") then
		dim sql
		dim rs
		dim webname
		dim weburl
		dim logo
		dim webmaster
		dim content
		dim email
		dim linktype
		dim pass
		dim id
		
		id=ChkRequest(request("id"),1)	'��ע��
		webname=checkstr(request.form("webname"))
		weburl=checkstr(request.form("weburl"))
		content=htmlencode(request.form("content"))
		logo=checkstr(request.form("logo"))
		webmaster=checkstr(request.form("webmaster"))
		email=checkstr(request.form("email"))
		linktype=checkstr(request.form("linktype"))
		pass=request.form("pass")
		
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_EC_Link_Table &" where id="&id
		rs.open sql,conn,3,3
		rs("linktype")=linktype
		rs("webname")=webname
		rs("content")=content
		rs("weburl")=weburl
		rs("logo")=logo
		rs("webmaster")=webmaster
		rs("email")=email
		rs("pass")=pass
		rs("dateandtime")=now()
		rs.update
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=LinkManage.asp"">"
		Show_Message("<p align=center><font color=red>���������޸ĳɹ���<br><br>"&freetime&"���Ӻ󷵻���ҳ!</font>")
		response.end
	else
		Show_Err("�Բ��𣬸ù����Ѿ�������ϵͳ����Ա�رգ���û��Ȩ�޲�����")
		response.end
	end if
end if%>