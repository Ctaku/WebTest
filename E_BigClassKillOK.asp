<!--#include file="Conn.ASP"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkURL.asp"-->
<%
if not(request.cookies(eChsys)("KEY")="super" or request.cookies(eChsys)("KEY")="typemaster") then
	Show_Err("对不起，您的后台管理权限不够！")
	response.end
else
	delbigclass=request.form("delbigclass")
	
	if delbigclass="1" then
		E_BigClassID=ChkRequest(request("E_BigClassID"),1)	'防注入
		dim E_BigClassName,E_typeid,E_typename
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_EC_BigClass_Table &" where E_BigClassID="&E_BigClassID
		rs.open sql,conn,3,3
		E_BigClassName=rs("E_BigClassName")
		E_typeid=rs("E_typeid")
		rs.close
		set rs=nothing
		set rs=server.createobject("adodb.recordset")
		sql="select * from EC_type where E_typeid="&E_typeid
		rs.open sql,conn,3,3
		E_typename=rs("E_typename")
		rs.close
		set rs=nothing
		button_value=trim(Request.Form("alert_button"))
		if button_value="是" then
			conn.execute("delete from EC_News where E_BigClassID=" &E_BigClassID)
			conn.execute("delete from E_Smallclass where E_BigClassID=" &E_BigClassID)
			conn.execute("delete from "& db_EC_BigClass_Table &" where E_BigClassID=" &E_BigClassID)
			conn.close
			set conn=nothing
			response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=E_BigClassManage.asp?E_typeid="&E_typeid&""">"
			Show_Message("<p align=center><font color=red>恭喜您!您选择的大类已经被删除!<br><br>"&freetime&"秒钟后返回上页!</font>")
			response.end
		else
			response.redirect "E_BigClassManage.asp?E_typeid="&E_typeid
		end if
	else
		Show_Err("您无权删除此项目！<br><br><a href='javascript:history.back()'>返回</a>")
		response.end
	end if
end if
%>