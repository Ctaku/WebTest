<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="chkuser.asp"-->
<!--#include file="ChkURL.asp"-->
<%
IF request.cookies(eChsys)("KEY")="bigmaster" THEN
	response.redirect "NewsAddd2.asp"
	response.end
else
IF request.cookies(eChsys)("KEY")="smallmaster" THEN
	response.redirect "E_NewsAddd3.asp"
	response.end
else
	dim jingyong
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_User_Table &" where  "& db_User_Name &"='"&request.cookies(eChsys)("username")&"'"
	rs.open sql,ConnUser,1,3
	jingyong=rs("jingyong")
	rs.close
	set rs=nothing
	
	NewsID=Request.QueryString("NewsID")
	set rs=server.CreateObject ("ADODB.RecordSet")
	rs.Source="select * from "& db_EC_Type_Table &" order by E_typeorder"
	rs.Open rs.source,conn,1,1

	%>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 添加文章</title>
</head>
<body topmargin="0">

<table border="1" width="100%" align=center cellpadding="0" cellspacing="0" style="border-collapse: collapse"  bordercolor="<%=border%>">
<form method="post" action="NewsAddd2.asp">
<tr align="center">
	<td colspan="2" class="TDtop" height=25>
		<div align="center" >┊ <B>添加文章</B> ┊</div>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align="center" height="160" >请选择文章总栏：
		<select size="1" name="E_typeid">
	<%if rs.EOF then %>
			<option value=0>暂无任何类别</option>
	<%else
		while not rs.EOF
			master=rs("typemaster")
			if master<>"" then
				tmaster=split(master,"|")
				for i=0 to ubound(tmaster)
					if request.cookies(eChsys)("username")=trim(tmaster(i)) then
						typemaster=true
						exit for
					else
						typemaster=false
					end if
				next
			end if
			if (request.cookies(eChsys)("key")="typemaster" and typemaster=true) or request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="check" or request.cookies(eChsys)("key")="selfreg" then
			%>
			<option value='<%=rs("E_typeid")%>'><%=trim(rs("E_typename"))%></option>
			<%
			end if
			rs.MoveNext
		wend
	end if
	%>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" width="100%" align="center" height="25" class="TDtop">
		<input type="button" value=" 返 回 " onClick="javascript:history.go(-1)" class="unnamed5" >&nbsp;
		<input type="submit" value="下 一 步" name="B1" >
	</td>
</tr>
</form>
</table>
</body>
</html>
<!--#include file="Admin_Bottom.asp"-->
	<%
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if
end if
%>