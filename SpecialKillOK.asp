<!--#include file="Conn.ASP"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="chkManage.asp" -->
<!--#include file="ChkURL.asp"-->
<%
if not(request.cookies(eChsys)("ManageKEY")="super" or request.cookies(eChsys)("ManageKEY")="typemaster" or request.cookies(eChsys)("ManageKEY")="bigmaster") then
	Show_Err("�Բ������ĺ�̨����Ȩ�޲���������")
	response.end
else
	SpecialID = ChkRequest(Request("SpecialID"),1)	'��ע��
	dim E_SpecialName
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select * from "& db_EC_Special_Table &" where SpecialID="&SpecialID
	rs.open sql,conn,3,3
	E_SpecialName=rs("E_SpecialName")
	rs.close
	set rs=nothing
	
	button_value=trim(Request.Form("alert_button"))
	if button_value="ȷ��" then
		conn.execute("delete from "& db_EC_Special_Table &" where Specialid=" &Specialid)
		if killnews=1 then
			conn.execute("delete from "& db_EC_News_Table &" where Specialid='" &Specialid &"'")
		end if
		conn.close
		set conn=nothing
		response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=SpecialManage.asp"">"
		Show_Message("<p align=center><font color=red>��ϲ��!��ѡ���ר���Ѿ���ɾ��!<br><br>"&freetime&"���Ӻ󷵻���ҳ!</font>")
		response.end
	end if
	response.redirect "E_Special.asp"
end if%>