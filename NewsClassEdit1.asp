<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkURL.asp"-->
<%
if not(request.cookies(eChsys)("KEY")="super" or request.cookies(eChsys)("KEY")="typemaster" or request.cookies(eChsys)("KEY")="bigmaster" or request.cookies(eChsys)("key")="selfreg") then
	Show_Err("�Բ������ĺ�̨����Ȩ�޲���������")
	response.end
else
	dim jingyong
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_User_Table &" where  "& db_User_Name &"='"&request.cookies(eChsys)("username")&"'"
	rs.open sql,ConnUser,1,3
	jingyong=rs("jingyong")
	rs.close
	set rs=nothing
	
	if request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="check" or (request.cookies(eChsys)("key")="selfreg" and jingyong=0) or request.cookies(eChsys)("key")="typemaster" then
		NewsID=Request.QueryString("NewsID")
		set rs=server.CreateObject ("ADODB.RecordSet")
		rs.Source="select * from "& db_EC_Type_Table &" order by E_typeorder"
		rs.Open rs.source,conn,1,1
		%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - �޸��������</title>
</head>
<body topmargin="0">

<table border="1" width="100%" align=center cellpadding="0" cellspacing="0" style="border-collapse: collapse"  bordercolor="<%=border%>">
<form method="POST" action="NewsClassEdit2.asp?NewsID=<%=NewsID%>">
<tr align="center" >
<td colspan="2" height="25" width="100%"  class="TDtop">
	<div align="center">�� <B>�޸���������޸�ID��Ϊ<%=NewsID%>�����µ����</B> ��</div>
</td>
</tr>
<tr>
<td colspan="2" align="center" height="160" bgcolor="#FFFFFF">��ѡ������������
	<select size="1" name="E_typeid">
	<%if rs.EOF then %>
<option value=0>�����κ����</option>
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
	if (request.cookies(eChsys)("key")="typemaster" and typemaster=true) or request.cookies(eChsys)("key")="super" then%>
<option value='<%=rs("E_typeid")%>'><%=trim(rs("E_typename"))%></option>
	<%end if
	rs.MoveNext
	wend
end if
%>	
</select>
</td>
</tr>
<tr>
<td colspan="2" width="100%" align="center" height="25" class="TDtop">
	<input type="button" value="��  ��" onclick="javascript:history.go(-1)" class="unnamed5"  >&nbsp;
	<input type="submit" value="��һ��" name="B1"  >
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
end if%>
