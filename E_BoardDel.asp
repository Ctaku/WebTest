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
	if ggmanage="1" or (request.cookies(eChsys)("ManagePurview")="99999" and request.cookies(eChsys)("purview")="99999") then
	id=ChkRequest(request.QueryString("id"),1)	'��ע��
	set rs=server.createobject("adodb.recordset")
	sql="delete from "& db_EC_Board_Table &" where ID="&id
	rs.open sql,conn,1,3
	conn.close
	set conn=nothing
	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=E_BoardManage.asp"">"
	Show_Message("<p align=center><font color=red>��ϲ��!����ɾ���ɹ�!<br><br>"&freetime&"���Ӻ󷵻���ҳ!</font>")
	response.end
	else
		Show_Err("�Բ��𣬸ù����Ѿ�������ϵͳ����Ա�رգ���û��Ȩ�޲�����")
		response.end
	end if
end if%>