<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="Char.inc"-->
<html>
<head>
<title></title>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%=request.cookies(eChsys)("ReplyContent")%>
<%
Dim action,newsID,rs,ReplyContent
dim sql
dim conn
Action=LCase(Request.QueryString("Action"))
LeadMailID=ChkRequest(Request.QueryString("LeadMailID"),1)	'防注入

if not IsNumeric(LeadMailID)then
	response.write "<script>alert('非法参数');history.back()</script>"
	response.end
end if

If request("action")="modify" Then
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_EC_Complain_Table &" where LeadMailID="&LeadMailID
	rs.open sql,conn,1,1
	If Not rs.Eof Then
		ReplyContent=rs("ReplyContent")
	End If
	Response.Write ReplyContent
End If
%>
</body>
</html>