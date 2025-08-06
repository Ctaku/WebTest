<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="98%" id="AutoNumber3" height="52">
<tr>
		<td align=left  width="70%" ><B>单位部门</B></td>
		<td align=right  width="30%" ><B>总计</B></td>
</tr>
	<%set rstopdep=server.createobject("adodb.recordset")
	sqltopdep="select top " & topuser & " * from "& db_EC_Dep_Table &" where depnumber>=0 order by depnumber desc"
	rstopdep.open sqltopdep,Conn,1,1
	if not rstopdep.EOF then
	while not rstopdep.EOF
	%>
	<tr>
		<td align="left" width="70%"><%=rstopdep("E_DepName")%>-<%=rstopdep("E_DepType")%></td>
		<td align="right" width="30%"><%=rstopdep("depnumber")%></a></td>
	</tr>
	<%rstopdep.MoveNext
	wend
	end if
	rstopdep.close
	set rstopdep=nothing
	%>
	<tr>
		<td align=right  width="75%" colspan="2"><a class=class href="E_Alldep.asp">更多</a></td>
	</tr>
</table>