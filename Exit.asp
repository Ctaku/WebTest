<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<%
dim ViewUrl
ViewUrl=request.cookies(eChsys)("ViewUrl")
if ViewUrl="" then
	ViewUrl="index.asp"
end if
if UserTableType = "Dvbbs" then		'是否整合论坛
	Response.Write "<iframe width='0' height='0' frameborder=no scrolling=no src='"& BbsPath &"logout.asp'></iframe>"
end if
response.cookies(eChsys)("UserName")=""
response.cookies(eChsys)("KEY")=""
response.cookies(eChsys)("purview")=""
response.cookies(eChsys)("fullname")=""
response.cookies(eChsys)("passwd")=""
response.cookies(eChsys)("sex")=""
response.cookies(eChsys)("reglevel")=""
response.cookies(eChsys)("shenhe")=""
response.redirect ViewUrl
%>