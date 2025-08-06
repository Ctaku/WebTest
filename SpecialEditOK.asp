<!--#INCLUDE FILE="Conn.asp" -->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="chkManage.asp" -->
<!--#include file="ChkURL.asp"-->
<%
if not(request.cookies(eChsys)("ManageKEY")="super" or request.cookies(eChsys)("ManageKEY")="typemaster" or request.cookies(eChsys)("ManageKEY")="bigmaster") then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	dim E_SpecialName
	dim SpecialID
	dim Specialzs
	dim name
	name=request.form("name")
	Specialzs=request.form("specialzs")
	SpecialID=ChkRequest(request("SpecialID"),1)	'防注入
	E_SpecialName=trim(request("E_SpecialName"))
	
	if E_SpecialName="" then
		Show_Err("专题名称不能为空！<br><br>－－－<a href='javascript:history.back()'>回去重来</a>－－－")
		response.end
	end if
	
	'修改专题库
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_EC_Special_Table &" where SpecialID="&SpecialID
	rs.open sql,conn,3,3
	rs("E_SpecialName")=E_SpecialName
	rs("Specialzs")=Specialzs
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	
	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=SpecialManage.asp"">"
	Show_Message("<p align=center><font color=red>恭喜您!您修改专题“"&name&"”为“"&E_SpecialName&"”成功!<br><br>"&freetime&"秒钟后返回上页!</font>")
	response.end
end if%>