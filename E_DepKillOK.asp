<!--#include file="Conn.ASP"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp"-->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp"-->
<%
IF request.cookies(eChsys)("ManageKEY")<>"super" then
	Show_Err("对不起，您的后台管理权限不够！")
	response.end
else
	if depmanage="1" or (request.cookies(eChsys)("ManagePurview")="99999" and request.cookies(eChsys)("purview")="99999") then

	ID = ChkRequest(Request("ID"),1)	'防注入
	set rs=server.createobject("adodb.recordset")
	sql="Select * from "& db_EC_Dep_Table &" where ID="&id
	rs.open sql,conn,1,3
	E_DepName=rs("E_DepName")
	rs.close
	set rs=nothing
	
	button_value=trim(Request.Form("alert_button"))
	if button_value="确定" then
		conn.execute("delete from "& db_EC_Dep_Table &" where ID=" &ID)
		conn.close
		set conn=nothing
	end if
	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=E_DepManage.asp"">"
	Show_Message("<p align=center><font color=red>单位删除成功!<br><br>"&freetime&"秒钟后返回上页!</font>")
	response.end
	else
		Show_Err("对不起，该功能已经被超级系统管理员关闭，您没有权限操作！")
		response.end
	end if
end if%>