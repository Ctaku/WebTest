<!--#INCLUDE FILE="Conn.asp" -->
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

	function changechr(str) 
		changechr=replace(replace(replace(replace(str,"<","&lt;"),">","&gt;"),chr(13),"<br>")," "," ") 
		changechr=replace(changechr,"'","&acute;")
		changechr=replace(changechr,mid(" "" ",2,1),"&quot;")
	end function

	dim E_DepName
	dim E_DepType
	dim ID
	ID=ChkRequest(request("ID"),1)	'防注入
	E_DepType=changechr(trim(request.form("E_DepType")))
	name=request.form("name")
	E_DepName=changechr(trim(request("E_DepName")))
	if E_DepName="" or E_DepType="" then
		Show_Err("单位部门名称不能为空<br><br><a href=javascript:history.go(-1)>返回重新填写</a>")
		response.end
	else
		'修改部门库
		dim depid
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_EC_Dep_Table &" where ID="&ID
		rs.open sql,conn,3,3
		depid=rs("id")
		rs("E_DepName")=E_DepName
		rs("E_DepType")=E_DepType
		rs.update
		rs.close
		set rs=nothing
	
		'修改用户库
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_User_Table &" where depid="&depid
		rs.open sql,ConnUser,3,3
		do while not rs.eof
			rs("E_DepName")=E_DepName
			rs.update
			rs.movenext
		loop
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=E_DepManage.asp"">"
		Show_Message("<p align=center><font color=red>单位部门修改成功!<br><br>"&freetime&"秒钟后返回上页!</font>")
		response.end
	end if
	else
		Show_Err("对不起，该功能已经被超级系统管理员关闭，您没有权限操作！")
		response.end
	end if
end if%>