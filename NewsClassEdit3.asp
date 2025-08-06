<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp" -->
<!--#include file="chkuser.asp" -->
<!--#include file="char.inc"-->
<!--#include file="ChkURL.asp"-->
<%
NewsID=ChkRequest(Request.QueryString("NewsID"),1)	'防注入
E_typeid=trim(request.form("E_typeid"))
E_BigClassID=trim(Request.Form("E_BigClassID"))
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from "& db_EC_News_Table &" where newsid="&newsid
rs.open sql,conn,3,3
dim title
title=rs("title")
rs.close
set rs=nothing
if E_typeid="" then
%>
<script language=javascript>
history.back()
alert("请选择该文章所属的总栏，如果系统中还没有文章总栏的话，请在《总栏管理》中添加新总栏！" )
</script>
<%
Response.End
end if

E_SmallClassID=trim(Request.Form("E_SmallClassID"))

set rs13=server.createobject("adodb.recordset")
sql13="select * from "& db_EC_News_Table &" where NewsID="&NewsID
rs13.open sql13,conn,3,3

if E_BigClassID="" then
rs13("E_BigClassID")=null
else
rs13("E_BigClassID")=E_BigClassID
end if

if E_SmallClassID="" then
rs13("E_SmallClassID")=null
else
rs13("E_SmallClassID")=E_SmallClassID
end if

rs13("E_typeid")=E_typeid
rs13.update
rs13.close
set rs13=nothing


Response.Redirect "E_NewsEdit.asp?NewsID="&NewsID
%>