<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="ChkUser.asp" -->
<!--#include file="ChkManage.asp" -->

<%IF not(request.cookies(eChsys)("ManageKEY")="super" or request.cookies(eChsys)("ManageKEY")="bigmaster" or request.cookies(eChsys)("ManageKEY")="check" or request.cookies(eChsys)("ManageKEY")="typemaster") THEN
	Show_Err("�Բ������Ĺ���Ȩ�޲�����")
	response.end
else
	aaas=(request.cookies(eChsys)("ManageKEY")="super")
	if votemana="1" or (request.cookies(eChsys)("ManagePurview")="99999" and request.cookies(eChsys)("purview")="99999") then
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_EC_Vote_Table &" order by id desc"
		rs.open sql,conn,1,1
		%>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - ͶƱ����</title>
</head>
<body topmargin="0">

<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFDFBF" width="100%" id="AutoNumber1">
<form method="POST" action="E_VoteSet.asp">
<tr class="TDtop">
	<td width="100%" height="25" colspan=5 align=center>�� <strong>ͶƱ����</strong> ��</td>
</tr>
<tr class="TDtop1" height=20>
	<td width="10%" align="center">ѡ��</td>
	<td width="10%" align="center">ID</td>
	<td width="60%" align="center">����</td>
	<td width="10%" align="center">�޸�</td>
	<td width="10%" align="center">ɾ��</td>
</tr>
		<%do while not rs.eof%>
<tr height=20>
	<td width="10%" align="center">
		<input type="radio" value=<%=rs("ID")%><%if rs("IsChecked")=1 then%> checked<%end if%> name="Checked">
	</td>
	<td width="10%" align="center"><%=rs("ID")%>��</td>
	<td width="60%"><%=rs("Title")%>��</td>
	<td width="10%" align="center">
		<input onclick="javascript:window.open('E_VoteModify.asp?id=<%=rs("ID")%>','_self','')" type="button" value="�޸�" name="button1" >
	</td>
	<td width="10%" align="center">
		<input onclick="javascript:window.open('E_VoteDel.asp?id=<%=rs("ID")%>','_self','')" type="button" value="ɾ��" name="button2" >
	</td>
</tr>
			<%rs.movenext
		loop
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		%>
<tr>
	<td colspan=5 align=center>
		<input type="submit" value="ѡ��ͶƱ��" name="submit" >
		<input onclick="javascript:window.open('E_VoteAdd.asp','_self','')" type="button" value="�����ͶƱ" name="button" >
	</td>
</tr>
</form>
</table>
</div>
</body>
</html>
<!--#include file="Admin_Bottom.asp"-->
	<%else
		Show_Err("�Բ��𣬸ù����Ѿ�������ϵͳ����Ա�رգ���û��Ȩ�޲�����")
		response.end
	end if
end if%>