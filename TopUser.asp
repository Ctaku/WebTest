<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="98%" id="AutoNumber3" height="52">
<tr>
		<td align=left  width="70%" ><B>用户名称</B></td>
		<td align=right  width="30%" ><B>总计</B></td>
</tr>
	<%set rs10=server.createobject("adodb.recordset")
	sql10="select top " & topuser & " * from "& db_User_Table &" where number>=1 order by number desc," & db_User_Id &" desc"
	rs10.open sql10,ConnUser,1,1
	if not rs10.EOF then
	while not rs10.EOF
	%>
	<tr>
		<td align="left" width="70%"><a href="user.asp?user=<%=rs10(db_User_Name)%>" target="_blank" class=class><%=rs10(db_User_Name)%></a></td>
		<td align="right" width="30%"><%=rs10("number")%></a></td>
	</tr>
	<%rs10.MoveNext
	wend
	end if
	rs10.close
	set rs10=nothing
	%>
	<tr>
		<td align=right  width="75%" colspan="2"><a class=class href="E_AllUser.asp">更多</a></td>
	</tr>
</table>