<!--#include file="Conn.ASP"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->

<%
IF request.cookies(eChsys)("ManageKEY")<>"super" then
	Show_Err("�Բ������ĺ�̨����Ȩ�޲�����")
	response.end
else
	Set rs6 = Server.CreateObject("ADODB.Recordset")
	sql6 ="select * from "& db_EC_Dep_Table &" where E_DepName='" & ChkRequest(Request("E_DepName"),0) & "' and E_DepType='" & ChkRequest(Request("E_DepType"),0) & "' order by ID"
	rs6.open sql6,Conn,3,3
	%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - �޸Ĳ���</title>
<LINK href=site.css rel=stylesheet>
</head>
<body topmargin="0">

<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFDFBF" width="100%" id="AutoNumber1">
<form method="POST" action="E_DepEditOK.asp">
<tr>
	<td width="100%" height="25" class="TDtop"> 
		<p align="center" >�� <strong>�� �� �� λ �� ��</strong> ��
	</td>
</tr>
<tr align="center"> 
	<td width="100%" bgcolor="#FFFFFF" height="150">
			<input type="hidden" name="name" value='<%=request("E_DepName")%>'>
			<input type="hidden" name="ID" value='<%=request("ID")%>'>
			<input type="hidden" name="type" value='<%=request("E_DepType")%>'>
		<table border="0" cellspacing="1" width="100%" height="70">
		<tr> 
			<td align="right" height="25" width="50%">��λ���ƣ�</td>
			<td align="left" width="50%">
				<input type="text" name="E_DepName" size="20" maxlength="20" value='<%=request("E_DepName")%>' >
			</td>
		</tr>
		<tr>
			<td align="right" height="25">�������ƣ�</td>
			<td align="left" height="25">
				<input type="text" name="E_DepType" size="20" maxlength="20" value='<%=request("E_DepType")%>' >
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center"> 
	<td width="100%" height="60" class="TDtop1">
		<input type="submit" value=" �� �� " name="cmdok" >&nbsp;
		<input type="reset" value=" �� ԭ " name="cmdcancel" >&nbsp; 
		<input type="button" name="ok" value=" �� �� " onClick="javascript:history.go(-1)" >
	</td>
</tr>
</form>
</table>
</body>
</html>
<!--#include file="Admin_Bottom.asp"-->
	<%rs6.close
	set rs6=nothing
	conn.close
	set conn=nothing
end if%>