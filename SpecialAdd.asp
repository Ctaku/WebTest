<!--#INCLUDE FILE="Conn.asp" -->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="chkManage.asp" -->
<!--#include file="ChkURL.asp"-->
<!--#include file="Function.asp"-->

<%
if not(request.cookies(eChsys)("ManageKEY")="super" or request.cookies(eChsys)("ManageKEY")="typemaster" or request.cookies(eChsys)("ManageKEY")="bigmaster") then
	Show_Err("�Բ������ĺ�̨����Ȩ�޲���������")
	response.end
else
	
	dim E_SpecialName
	specialzs=CheckStr(trim(request.form("specialzs")))
	E_SpecialName=CheckStr(trim(request("E_SpecialName")))
	
	If E_SpecialName="" Then
		Show_Err("ר�����Ʋ���Ϊ�գ�<br><br>������<a href='javascript:history.back()'>��ȥ����</a>������")
		response.end
	end if
	
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_EC_Special_Table &""
	rs.open sql,conn,3,3
	do while not rs.eof
		if rs("E_SpecialName")=E_SpecialName then
			Show_Err("�Ѿ��������ר�⣡<br><br>������<a href='javascript:history.back()'>����������д</a>������")
			response.end
		end if
		rs.movenext
	loop
	rs.close
	set rs=nothing

	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select * from "& db_EC_Special_Table &""
	rs.open sql,conn,3,3
	rs.addnew
	rs("E_SpecialName")=E_SpecialName
	rs("Specialzs")=Specialzs
	rs("Specialmaster")=request.cookies(eChsys)("ManageUserName")
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing

	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=SpecialManage.asp"">"
	Show_Message("<p align=center><font color=red>��ϲ��!�����ר�⡰"&E_SpecialName&"���ɹ�!<br><br>"&freetime&"���Ӻ󷵻���ҳ!</font>")
	response.end
end if
%>