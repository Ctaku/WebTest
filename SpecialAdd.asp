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
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	
	dim E_SpecialName
	specialzs=CheckStr(trim(request.form("specialzs")))
	E_SpecialName=CheckStr(trim(request("E_SpecialName")))
	
	If E_SpecialName="" Then
		Show_Err("专题名称不能为空！<br><br>－－－<a href='javascript:history.back()'>回去重来</a>－－－")
		response.end
	end if
	
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_EC_Special_Table &""
	rs.open sql,conn,3,3
	do while not rs.eof
		if rs("E_SpecialName")=E_SpecialName then
			Show_Err("已经存在这个专题！<br><br>－－－<a href='javascript:history.back()'>返回重新填写</a>－－－")
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
	Show_Message("<p align=center><font color=red>恭喜您!您添加专题“"&E_SpecialName&"”成功!<br><br>"&freetime&"秒钟后返回上页!</font>")
	response.end
end if
%>