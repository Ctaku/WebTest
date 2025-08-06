<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="ChkUser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp"-->
<%
if not(request.cookies(eChsys)("ManageKEY")="super" or request.cookies(eChsys)("ManageKEY")="bigmaster" or request.cookies(eChsys)("ManageKEY")="check" or request.cookies(eChsys)("ManageKEY")="typemaster") then
	Show_Err("对不起，您的管理权限不够！")
	response.end
else
	if linkmana="1" or (request.cookies(eChsys)("ManagePurview")="99999" and request.cookies(eChsys)("purview")="99999") then
		dim sql
		dim rs
		dim webname
		dim weburl
		dim logo
		dim webmaster
		dim content
		dim email
		dim linktype
		dim pass
		
		webname=checkstr(request.form("webname"))

if Instr(request("webname"),"=")>0 or Instr(request("webname"),"%")>0 or Instr(request("webname"),chr(32))>0 or Instr(request("webname"),"?")>0 or Instr(request("webname"),"&")>0 or Instr(request("webname"),";")>0 or Instr(request("webname"),",")>0 or Instr(request("webname"),"'")>0 or Instr(request("webname"),",")>0 or Instr(request("webname"),chr(34))>0 or Instr(request("webname"),chr(9))>0 or Instr(request("webname"),"")>0 or Instr(request("webname"),"$")>0 then
	Response.Write "<script>alert('对不起，您输入的网站名称中包含有非法字符。');history.back();</Script>"
   	Response.End 
end if

		weburl=checkstr(request.form("weburl"))

if Instr(request("weburl"),"=")>0 or Instr(request("weburl"),"%")>0 or Instr(request("weburl"),chr(32))>0 or Instr(request("weburl"),"?")>0 or Instr(request("weburl"),"&")>0 or Instr(request("weburl"),";")>0 or Instr(request("weburl"),",")>0 or Instr(request("weburl"),"'")>0 or Instr(request("weburl"),",")>0 or Instr(request("weburl"),chr(34))>0 or Instr(request("weburl"),chr(9))>0 or Instr(request("weburl"),"")>0 or Instr(request("weburl"),"$")>0 then
	Response.Write "<script>alert('对不起，您输入的网站地址中包含有非法字符。');history.back();</Script>"
   	Response.End 
end if

		content=htmlencode(request.form("content"))

if Instr(request("content"),"=")>0 or Instr(request("content"),"%")>0 or Instr(request("content"),chr(32))>0 or Instr(request("content"),"?")>0 or Instr(request("content"),"&")>0 or Instr(request("content"),";")>0 or Instr(request("content"),",")>0 or Instr(request("content"),"'")>0 or Instr(request("content"),",")>0 or Instr(request("content"),chr(34))>0 or Instr(request("content"),chr(9))>0 or Instr(request("content"),"")>0 or Instr(request("content"),"$")>0 then
	Response.Write "<script>alert('对不起，您输入的网站简介中包含有非法字符。');history.back();</Script>"
   	Response.End 
end if

		logo=checkstr(request.form("logo"))

if Instr(request("logo"),"=")>0 or Instr(request("logo"),"%")>0 or Instr(request("logo"),chr(32))>0 or Instr(request("logo"),"?")>0 or Instr(request("logo"),"&")>0 or Instr(request("logo"),";")>0 or Instr(request("logo"),",")>0 or Instr(request("logo"),"'")>0 or Instr(request("logo"),",")>0 or Instr(request("logo"),chr(34))>0 or Instr(request("logo"),chr(9))>0 or Instr(request("logo"),"")>0 or Instr(request("logo"),"$")>0 then
	Response.Write "<script>alert('对不起，您输入的网站LOGO中包含有非法字符。');history.back();</Script>"
   	Response.End 
end if

		webmaster=checkstr(request.form("webmaster"))

if Instr(request("webmaster"),"=")>0 or Instr(request("webmaster"),"%")>0 or Instr(request("webmaster"),chr(32))>0 or Instr(request("webmaster"),"?")>0 or Instr(request("webmaster"),"&")>0 or Instr(request("webmaster"),";")>0 or Instr(request("webmaster"),",")>0 or Instr(request("webmaster"),"'")>0 or Instr(request("webmaster"),",")>0 or Instr(request("webmaster"),chr(34))>0 or Instr(request("webmaster"),chr(9))>0 or Instr(request("webmaster"),"")>0 or Instr(request("webmaster"),"$")>0 then
	Response.Write "<script>alert('对不起，您输入的网站站长名中包含有非法字符。');history.back();</Script>"
   	Response.End 
end if

		email=checkstr(request.form("email"))

if IsValidEmail(email)=0 then
	Response.Write "<script>alert('对不起，您输入的email非法。');history.back();</Script>"
   	Response.End 
end if

		linktype=checkstr(request.form("linktype"))
		pass=request.form("pass")
		
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_EC_Link_Table &" where (id is null)" 
		
		rs.open sql,conn,1,3
		rs.addnew
		rs("linktype")=linktype
		rs("webname")=webname
		rs("content")=content
		rs("weburl")=weburl
		rs("logo")=logo
		rs("webmaster")=webmaster
		rs("email")=email
		rs("pass")=pass
		rs("dateandtime")=now()
		rs.update
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=LinkManage.asp"">"
		Show_Message("<p align=center><font color=red>友情链接添加成功!<br><br>"&freetime&"秒钟后返回上页!</font>")
		response.end
	else
		Show_Err("对不起，该功能已经被超级系统管理员关闭，您没有权限操作！")
		response.end
	end if
end if%>