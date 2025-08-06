<!--#INCLUDE FILE="Conn.asp" -->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkURL.asp"-->
<%
IF request.cookies(eChsys)("KEY")<>"super" and request.cookies(eChsys)("KEY")<>"typemaster" THEN
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	if request("action")="update" then
		dim E_bigclassorder,bigclassmaster,E_BigClassView,E_BigClassName,E_typeid,bigclasszs,classurl
		for i=1 to request.form("E_BigClassID").count
			E_BigClassID=CheckStr(request.form("E_BigClassID")(i))
			E_bigclassorder=CheckStr(request.form("E_bigclassorder")(i))
			classurl=CheckStr(request.form("url")(i))
			bigclassmaster=CheckStr(request.form("bigclassmaster")(i))
			E_BigClassView=CheckStr(request.form("E_BigClassView")(i))
			E_BigClassName=CheckStr(request.form("E_BigClassName")(i))
			bigclasszs=CheckStr(request.form("bigclasszs")(i))
			E_typeid=CheckStr(request.form("E_typeid")(i))
			if CheckStr(request.form("E_bigclassorder")(i))="" then
				Show_Err("请填写大类排序！<br><br><a href='javascript:history.back()'>返回</a>")
				response.end
			end if
			if replace(request.form("bigclassmaster")(i),"'","")="" then
				bigclassmaster="无"
			else
				master=split(CheckStr(request.form("bigclassmaster")(i)),"|")
		 		for k=0 to ubound(master)
					set rs=server.createobject("adodb.recordset")
					sql="Select * from "& db_User_Table &" where oskey='bigmaster' and "& db_User_Name &"='"&trim(master(k))&"'"
					rs.open sql,ConnUser,1,3
					if trim(master(k))<>"无" then
						if rs.eof then
							Show_Err("大类管理员中无"& trim(master(k)) &"用户！请重新选择该大类的管理员！<br><br><a href='javascript:history.back()'>返回</a>")
							Response.End
						end if
					else
						bigclassmaster="无"
					end if
					rs.close
					set rs=nothing
		 		next
			end if
		
			conn.execute("update "& db_EC_BigClass_Table &" set E_bigclassorder="&E_bigclassorder&",bigclassmaster='"&bigclassmaster&"',E_BigClassView="&E_BigClassView&",E_BigClassName='"&E_BigClassName&"',E_typeid="&E_typeid&",bigclasszs='"&bigclasszs&"',url='"&classurl&"' where E_BigClassID="&E_BigClassID)
		    Set nrs = Server.CreateObject("ADODB.Recordset")
		    sqln="select * from EC_News where E_BigClassID="&E_BigClassID
		    nrs.open sqln,conn,3,3
		    while not nrs.EOF
		    nrs("E_typeid")=E_typeid
		    nrs.MoveNext
		    wend
		    nrs.close
		    set nrs=nothing
		
		next
	end if
	
	if request("action")="add" then
	bigclasszs=request.form("bigclasszs")
	E_BigClassName=trim(request("E_BigClassName"))
	E_bigclassorder=request.form("E_bigclassorder")
	bigclassmaster=request.form("bigclassmaster")
	E_BigClassView=request.form("E_BigClassView")
	E_typeid=request.form("E_typeid")
	classurl=request.Form("url")
	
	If E_bigclassorder="" Then
		Show_Err("大类排序不能为空！请<a href=javascript:history.go(-1)>返回重新填写</a>！")
		response.end
	end if
	If E_BigClassName="" Then
		Show_Err("大类名称不能为空！请<a href=javascript:history.go(-1)>返回重新填写</a>！")
		response.end
	end if
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select * from "& db_EC_BigClass_Table &""
	rs.open sql,conn,3,3
	rs.addnew
	rs("E_BigClassName")=E_BigClassName
	rs("bigclasszs")=bigclasszs
	rs("E_typeid")=E_typeid
	rs("url")=classurl
	if E_bigclassorder="" then
		rs("E_bigclassorder")=0
	else
		rs("E_bigclassorder")=E_bigclassorder
	end if
	rs("E_BigClassView")=E_BigClassView
	if bigclassmaster="" then
		rs("bigclassmaster")="无"
	else
		rs("bigclassmaster")=bigclassmaster
	end if
	rs.update
	E_typeid=rs("E_typeid")
	rs.close
	set rs=nothing
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select * from EC_type where E_typeid="&E_typeid
	rs.open sql,conn,3,3
	dim E_typename
	E_typename=rs("E_typename")
	rs.close
	set rs=nothing
	end if
	
	conn.close
	set conn=nothing
	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=E_BigClassManage.asp?E_typeid="&E_typeid&""">"
	Show_Message("<p align=center><font color=red>恭喜您!设置成功!<br><br>"&freetime&"秒钟后返回上页!</font>")
	response.end
end if
%>