<%set rs=server.CreateObject("ADODB.RecordSet") 
rs.Source="select top 1 * from " & db_User_Table & " order by " & db_User_ID & " desc"
rs.Open rs.Source,ConnUser,1,1
if not rs.EOF then
%>
	&nbsp;&nbsp;��- �����Ա��<a href=user.asp?user=<%=rs(db_User_Name)%>  target="_blank" class=class><font color=red><%=left(rs(db_User_Name),8)%></font></a>
<%end if
rs.Close
set rs=nothing
%>
