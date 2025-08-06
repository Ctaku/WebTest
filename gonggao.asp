<%@ Language=VBScript%>
<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="function.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>公告信息</title>
<LINK href="news.css" rel="stylesheet">
</head>
<body>
<table width="300" height="400" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#990000">
  <tr>
    <td height="20"><img src="Images/tan01.gif" width="300" height="20"></td>
  </tr>
  <tr>
    <td height="340" valign="top" background="Images/tan02.gif"><div align="center">
	
      <table width="94%" border="0" align="center" cellpadding="0" cellspacing="4">
        <%
set rs=server.CreateObject("ADODB.RecordSet") 
rs.Source="select * from "& db_EC_Board_Table &" where inuse=1"
rs.Open rs.Source,conn,1,1
if not rs.EOF then
%>
  <tr>
   <td>
   <div align="center"><A HREF="E_Board_News.asp?ID=<%=rs("ID")%>" class="middle" target="_blank"><%=rs("title")%></A></div></td>   </tr>
		<tr>
          <td>
		  <HR align="center" width="95%" noshade style="color:#FCDBC1">
		    <div align="center">发布人：<%=rs("upload")%>&nbsp;发布时间：<%=year(rs("dateandtime"))%>-<%=month(rs("dateandtime"))%>-<%=day(rs("dateandtime"))%> 
            </div></td>
			</tr>
			 <tr><td><%=CutStr(rs("content"),600)%> </td></tr>
        <%else
     rs.close
      set rs=nothing
      end if%>
      </table>
    </div></td>
  </tr>
  <tr>
    <td height="17"><img src="Images/tan03.gif" width="300" height="17"></td>
  </tr>
</table>
 
</body>
</html>
