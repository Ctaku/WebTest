<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="ChkUser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp" -->
<%
if not(request.cookies(eChsys)("ManageKEY")="super" or request.cookies(eChsys)("ManageKEY")="bigmaster" or request.cookies(eChsys)("ManageKEY")="check" or request.cookies(eChsys)("ManageKEY")="typemaster") then
	Show_Err("�Բ������Ĺ���Ȩ�޲�����")
	response.end
else
	if ggmanage="1" or (request.cookies(eChsys)("ManagePurview")="99999" and request.cookies(eChsys)("purview")="99999") then

	ID=ChkRequest(Request.QueryString("ID"),1)	'��ע��
	Title=Request.Form("Title")
	if Title="" then
		Show_Err("����д���±��⣡<br><br><a href='javascript:history.back()'>��ȥ����</a>")
		response.end
	end if

	if len(title)>100 then
		Show_Err("��������������<br><br><a href='javascript:history.back()'>��ȥ����</a>")
		response.end
	end if
	upload=Request.Form("upload")
	content=Request.Form("content")
	if Content="" then
		Show_Err("�������������ݣ�<br><br><a href='javascript:history.back()'>��ȥ����</a>")
		response.end
	end if

	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_EC_Board_Table &" where id="&id
	rs.open sql,conn,3,3
	rs("title")=title
	rs("upload")=upload
	rs("content")=content
	rs("dateandtime")=now()
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing

	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=E_BoardManage.asp"">"
	Show_Message("<p align=center><font color=red>��ϲ��!����Ϊ��"&title&"���Ĺ����޸ĳɹ�!<br><br>"&freetime&"���Ӻ󷵻���ҳ!</font>")
	response.end
	else
		Show_Err("�Բ��𣬸ù����Ѿ�������ϵͳ����Ա�رգ���û��Ȩ�޲�����")
		response.end
	end if
end if%>