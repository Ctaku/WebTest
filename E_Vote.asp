<% Response.Buffer=True %>
<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<%
if request.QueryString("stype")="" then
	if Request.ServerVariables("REMOTE_ADDR")=request.cookies(eChsys)("IPAddress") then
		response.write"<SCRIPT language=JavaScript>alert('感谢您的支持，您已经投过票了，请勿重复投票，谢谢！');"
		response.write"javascript:window.close();</SCRIPT>"
	else
		options=request.form("options")
		response.cookies(eChsys)("IPAddress")=Request.ServerVariables("REMOTE_ADDR") 
		set rs=server.createobject("adodb.recordset")
		sql="update "& db_EC_Vote_Table &" set answer"&options&"=answer"&options&"+1 where IsChecked=1"
		rs.open sql,conn,1,3
		set rs=nothing
	end if
end if
%>
<head> 
<title><%=redcaff%>投票结果</title>
<LINK href=news.css rel=stylesheet>
</head>
<body>
	<table width="780" height="150" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor=<%=border%> bgcolor="#FFFFFF" style="border-collapse: collapse">
    		<br>
    		<br><br><br><%
		total=0
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_EC_Vote_Table &" where IsChecked=1"
		rs.open sql,conn,1,1
		%>
		<tr align="center" valign="middle"> 
			<td height="24" colspan="4"><font color="#000000">『</font><font color="#000073"><%=rs("Title")%></font><font color="#000000">』投票结果</font></td>
	  </tr>
		<tr>
			<td height="24" valign="top" colspan="4"><br>		  </td>
		</tr>
		<tr valign="middle">
			<td width="22%" height="27">序号</td>
			<td>百比分</td>
			<td colspan="2" align="center">人数</td>
		</tr>
		<%
		for i=1 to 8
			if rs("Select"&i)<>"" then
				total=total+rs("answer"&i)
			end if
		next
		for i=1 to 8
			if rs("Select"&i)<>"" then
				if total=0 then
					answer=0
				else
					answer=(rs("answer"&i)/total)*100
				end if%>
		<tr valign="middle">
			<td height="29"><%=i%>.<%=rs("select"&i)%>：</td>
			<td width="56%"><img src=images/RSCount.gif width=<%=int(answer*2)%> height=8></td>
			<td width="12%"><%=round(answer,3)%>%</td>
			<td width="10%"><%=rs("answer"&i)%>人</td>
			<%end if
		next%>
	  </tr>
		<tr>
		  <td colspan="4"> 共有【<%=total%>】人参加投票<br></td>
		</tr>
				<tr> <td colspan="4">
		<div align="center" style="background:#FFFFFF; color:#993300; font-size:13px;">非常感谢您的支持您可以如下操作：<BR><BR>【<a href="E_GuestBook.asp" target="_blank">留下建议</a>】【<a href="javascript:window.close()">关闭窗口</a>】<br>
	<% rs.close     
	set rs=nothing     
	conn.close     
	set conn=nothing %>
</div></td></tr>
</table>

</body>
