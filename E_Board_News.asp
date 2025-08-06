<%@ Language=VBScript%>
<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--'#include file="char.inc"-->
<!--#include file="function.asp"-->
<%
request_BoardID=ChkRequest(Request.QueryString("ID"),1)	'防注入
if request_BoardID="" then
Response.Write "<script>alert('未指定参数');history.back()</script>"
response.end
else

if not IsNumeric(request_BoardID) then
 response.write "<script>alert('非法参数');history.back()</script>"
 response.end
else

set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from "& db_EC_Board_Table &" where id="&request_BoardID&" order by ID"
rs.Open rs.Source,conn,1,1
if rs.EOF then
Response.Write "<script>alert('该公告不存在');history.back()</script>"
else
title=rs("title")
rs.close
set rs=nothing
end if

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>公告详细内容_<%=title%></title>
<LINK href="news.css" rel=stylesheet>
</head>
<body>
<table width="641" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><img src="Images/tan1.gif" width="641" height="43"></td>
  </tr>
  <tr>
    <td height="144" valign="top" background="Images/tan2.gif">
<%
set rs2=server.CreateObject("ADODB.RecordSet")
rs2.Source="select * from "& db_EC_Board_Table &" where id="&request_BoardID&" order by ID"
rs2.Open rs2.Source,conn,1,1
content=CheckStr(rs2("Content"))
%>
<table width="95%" border="0" align="center" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF" style="WORD-WRAP: break-word;">
  <tr>
   <td height="20"><BR>
<div align="center"><font color="<%=titlecolor%>" size="+2" face="楷体_GB2312"><strong><%=RS2("title")%></strong></font></div>
								  <HR style="color: #FCD3B8" width="90%"></td>
  </tr>
  <tr>
   <td height="20" ><div align="center">发布人：<%=rs2("upload")%>&nbsp;&nbsp;发布时间：<%=rs2("dateandtime")%></div></td>
   </tr>
   <tr>
    <td>
     <%	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;" & Content%><br>
    </td>
   </tr>
</table>

<%
rs2.close
set rs2=nothing
%>
<%
end if
end if 
%>	</td>
  </tr>
  <tr>
    <td><img src="Images/tan3.gif" width="641" height="37"></td>
  </tr>
</table>


