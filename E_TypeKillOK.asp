<!--#include file="Conn.ASP"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="char.inc"-->
<!--#include file="config.asp"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkURL.asp"-->
<%
IF request.cookies(eChsys)("KEY")<>"super" THEN
	Show_Err("对不起，您的后台管理权限不够！")
	response.end
else
	E_typeid = ChkRequest(Request("E_typeid"),1)	'防注入
	dim E_typename
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_EC_Type_Table &" where E_typeid="&E_typeid
	rs.open sql,conn,3,3
	E_typename=rs("E_typename")
	rs.close
	set rs=nothing
	
	button_value=trim(Request.Form("alert_button"))
	if button_value="是" then
		conn.execute("delete from "& db_EC_News_Table &" where E_typeid=" &E_typeid)
		conn.execute("delete from "& db_EC_SmallClass_Table &" where E_typeid=" &E_typeid)
		conn.execute("delete from "& db_EC_BigClass_Table &" where E_typeid=" &E_typeid)
		conn.execute("delete from "& db_EC_Type_Table &" where E_typeid=" &E_typeid)
		conn.close
		set conn=nothing
		response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=E_TypeManage.asp"">"
		Show_Message("<p align=center><font color=red>恭喜您!您选择的总栏已经被删除!<br><br>"&freetime&"秒钟后返回上页!</font>")
		response.end
	else
		response.redirect "E_TypeManage.asp"
	end if
end if
%>