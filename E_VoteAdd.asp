<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="ChkUser.asp" -->
<!--#include file="ChkManage.asp" -->
<%
IF not(request.cookies(eChsys)("ManageKEY")="super" or request.cookies(eChsys)("ManageKEY")="bigmaster" or request.cookies(eChsys)("ManageKEY")="check" or request.cookies(eChsys)("ManageKEY")="typemaster") THEN
	Show_Err("�Բ������Ĺ���Ȩ�޲�����")
	response.end
else
	if votemana="1" or (request.cookies(eChsys)("ManagePurview")="99999" and request.cookies(eChsys)("purview")="99999") then
	%>
<head><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - ���ͶƱ</title>
</head>
<body topmargin="0">

<table border="1" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="<%=border%>" align=center>
<form method="POST" action="E_VoteSave.asp">
<tr class="TDtop">
	<td width="100%" height="25" colspan=2 align=center>�� <strong>�� �� Ͷ Ʊ</strong> ��</td>
</tr>
<tr>
	<td colspan=2 align=center></td>
</tr>
<tr>
	<td width="40%" align="right">���⣺</td>
	<td width="60%">
		<input type="text" name="Title" size="20" >
	</td>
</tr>
		<%for i=1 to 8%>
<tr>
	<td align="right">ѡ��<%=i%>��</td>
	<td>
		<input type="text" name="select<%=i%>" size="20" >
		Ʊ����<input type="text" name="answer<%=i%>" size="5" >
	</td>
</tr>
		<%next%>
<tr>
	<td colspan=2 align=center></td>
</tr>
<tr class="TDtop">
	<td colspan=2 align=center>
		<input type="hidden" value="add" name="act">
		<input type="button" value=" �� �� " onclick="javascript:history.go(-1)" >&nbsp; 
		<input type="submit" value=" �� �� " name="cmdok" >&nbsp; 
		<input type="reset" value=" ȡ �� "  name="cmdcancel" >
	</td>
</tr>
<tr>
	<td colspan=2 align=center></td>
</tr>
</form>
</table>
</body>
</html>
<!--#include file="Admin_Bottom.asp"-->
		<%
		conn.close
		set conn=nothing
	else
		Show_Err("�Բ��𣬸ù����Ѿ�������ϵͳ����Ա�رգ���û��Ȩ�޲�����")
		response.end
	end if
end if%>