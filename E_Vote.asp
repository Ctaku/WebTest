<% Response.Buffer=True %>
<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<%
if request.QueryString("stype")="" then
	if Request.ServerVariables("REMOTE_ADDR")=request.cookies(eChsys)("IPAddress") then
		response.write"<SCRIPT language=JavaScript>alert('��л����֧�֣����Ѿ�Ͷ��Ʊ�ˣ������ظ�ͶƱ��лл��');"
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
<title><%=redcaff%>ͶƱ���</title>
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
			<td height="24" colspan="4"><font color="#000000">��</font><font color="#000073"><%=rs("Title")%></font><font color="#000000">��ͶƱ���</font></td>
	  </tr>
		<tr>
			<td height="24" valign="top" colspan="4"><br>		  </td>
		</tr>
		<tr valign="middle">
			<td width="22%" height="27">���</td>
			<td>�ٱȷ�</td>
			<td colspan="2" align="center">����</td>
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
			<td height="29"><%=i%>.<%=rs("select"&i)%>��</td>
			<td width="56%"><img src=images/RSCount.gif width=<%=int(answer*2)%> height=8></td>
			<td width="12%"><%=round(answer,3)%>%</td>
			<td width="10%"><%=rs("answer"&i)%>��</td>
			<%end if
		next%>
	  </tr>
		<tr>
		  <td colspan="4"> ���С�<%=total%>���˲μ�ͶƱ<br></td>
		</tr>
				<tr> <td colspan="4">
		<div align="center" style="background:#FFFFFF; color:#993300; font-size:13px;">�ǳ���л����֧�����������²�����<BR><BR>��<a href="E_GuestBook.asp" target="_blank">���½���</a>����<a href="javascript:window.close()">�رմ���</a>��<br>
	<% rs.close     
	set rs=nothing     
	conn.close     
	set conn=nothing %>
</div></td></tr>
</table>

</body>
