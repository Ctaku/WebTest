<!--#include file="Conn.ASP"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkURL.asp"-->
<%
if not(request.cookies(eChsys)("KEY")="super" or request.cookies(eChsys)("KEY")="bigmaster" or request.cookies(eChsys)("KEY")="typemaster") THEN
	Show_Err("�Բ������ĺ�̨����Ȩ�޲�����")
	response.end
else
	delsmallclass=request.form("delsmallclass")
	if delsmallclass="1" then	
		dim E_smallclassname
		E_SmallClassID=ChkRequest(request("E_SmallClassID"),1)	'��ע��
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_EC_SmallClass_Table &" where E_SmallClassID="&E_SmallClassID
		rs.open sql,conn,1,1
		E_smallclassname=rs("E_smallclassname")
		E_BigClassID=rs("E_BigClassID")
		rs.close
		set rs=nothing
	
		button_value=trim(Request.Form("alert_button"))
		if button_value="��" then
			conn.execute("delete from "& db_EC_News_Table &" where E_SmallClassID=" &E_SmallClassID)
			conn.execute("delete from "& db_EC_SmallClass_Table &" where E_SmallClassID=" &E_SmallClassID)
			conn.close
			set conn=nothing
			response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=E_SmallClassManage.asp?E_BigClassID="&E_BigClassID&""">"
			Show_Message("<p align=center><font color=red>��ϲ��!��ɾ��С�ࡰ"&E_smallclassname&"���ɹ�!!<br><br>"&freetime&"���Ӻ󷵻���ҳ!</font>")
			response.end
		else
			response.redirect "E_SmallClassManage.asp?E_BigClassID="&E_BigClassID
		end if
	else
		Show_Err("����Ȩɾ������Ŀ��<br><br><a href='javascript:history.back()'>����</a>")
		response.end
	end if
end if%>