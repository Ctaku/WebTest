<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="char.inc"-->
<!--#include file="config.asp"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp" -->
<%
IF not(request.cookies(eChsys)("ManageKEY")="super" or request.cookies(eChsys)("ManageKEY")="typemaster" or request.cookies(eChsys)("ManageKEY")="bigmaster") then
	Show_Err("�Բ������ĺ�̨����Ȩ�޲���������")
	response.end
else
    if request.cookies(eChsys)("purview")="99999" and request.cookies(eChsys)("KEY")="super" then
	%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - �޸�����</title>
<script LANGUAGE="javascript">
<!--
function FrmAddLink_onsubmit() {
 if (document.FrmAddLink.passwd.value=="")
	{
	  alert("�Բ��������������룡")
	  document.FrmAddLink.passwd.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd.value.length < 4)
	{
	  alert("Ϊ�˰�ȫ����������Ӧ�ó�һ�㣡")
	  document.FrmAddLink.passwd.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd.value.length > 16)
	{
	  alert("��������̫���˰ɣ�")
	  document.FrmAddLink.passwd.focus()
	  return false
	 }
else if (document.FrmAddLink.username.value==document.FrmAddLink.passwd.value)
	{
	  alert("Ϊ�˰�ȫ���û��������벻Ӧ����ͬ��")
	  document.FrmAddLink.passwd.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd2.value=="")
	{
	  alert("�Բ�������������֤���룡")
	  document.FrmAddLink.passwd2.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd2.value !== document.FrmAddLink.passwd.value)
	{
	  alert("�Բ�����������������벻һ�£�")
	  document.FrmAddLink.passwd2.focus()
	  return false
	 }
else if (document.FrmAddLink.fullname.value=="")
	{
	  alert("�Բ�����������ʵ������")
	  document.FrmAddLink.fullname.focus()
	  return false
	 }
else if (document.FrmAddLink.adder.value=="")
	{
	  alert("�Բ����������Ӧǰ̨��ͨ�����û���")
	  document.FrmAddLink.adder.focus()
	  return false
	 }
else if (document.FrmAddLink.depid.value=="")
	{
	  alert("�Բ�����ѡ������λ��")
	  document.FrmAddLink.depid.focus()
	  return false
	 }
}

//-->
</script>
</head>
<body topmargin="0">
<%
	dim sql
	dim rs
	sql="select * from "& db_EC_Manage_Table &" where "& db_ManageUser_ID &"="&ChkRequest(request("id"),1)	'��ע��
	set rs=server.createobject("adodb.recordset")
	rs.open sql,Conn,1,1
	if request.cookies(eChsys)("ManageKEY")="super" or rs("oskey")=request.cookies(eChsys)("ManageKEY") then
		%>

<table width="100%" border="1" align=center cellpadding="0" cellspacing="0" bordercolor="<%=border%>" bgcolor="#FFFFFF" style="border-collapse: collapse">
<form method="POST" name="FrmAddLink" LANGUAGE="javascript" onSubmit="return FrmAddLink_onsubmit()"   action="SaveAdminEdit.asp?id=<%=ChkRequest(request("id"),1)%>">
<tr>
	<td width="100%" height=32 colspan="2" class="TDtop">
	<p align="center" >�� <strong>�����û������޸�</strong> ��
	</td>
</tr>
<tr>
	<td width="46%" height="32" align="right" valign="middle">
		<div align="center"><span class="smallFont">�����û�����</span></div>
	</td>
	<td width="54%" height="32">
		<div align="left">
			<input name="username" size="20" value="<%=rs(db_ManageUser_Name)%>" <%if (rs(db_ManageUser_Name)<>request.cookies(eChsys)("ManageUserName")  and request.cookies(eChsys)("purview")="99999") or request.cookies(eChsys)("purview")<>"99999" then%>readonly<%end if%> class="smallInput" maxlength="30" >
		</div>
	</td>
</tr>
<tr>
	<td width="46%" height="32" align="right">
		<div align="center"><span class="smallFont">��&nbsp;&nbsp;&nbsp;�룺</span></div>
	</td>
	<td width="54%" height="32">
		<div align="left">
			<input name="passwd" type="password" size="20" value="<%=rs(db_ManageUser_Password)%>" class="smallInput" maxlength="50" >
		</div>
	</td>
</tr>
<tr>
	<td width="46%" height="32" align="right">
		<div align="center"><span class="smallFont">��֤���룺</span></div>
	</td>
	<td width="54%" height="32">
		<div align="left">
			<input name="passwd2" type="password" size="20" value="<%=rs(db_ManageUser_Password)%>" class="smallInput" maxlength="50" >
		</div>
	</td>
</tr>
<tr>
	<td width="46%" height="32" align="right">
		<div align="center"><span class="smallFont">��ʵ������</span></div>
	</td>
	<td width="54%" height="32">
		<div align="left">
			<input name="fullname" size="20" value="<%=rs("fullname")%>" class="smallInput" maxlength="50" >
		</div>
	</td>
</tr>
<tr>
	<td width="46%" height="32" align="right">
		<div align="center"><span class="smallFont">��Ӧǰ̨��ͨ�����û���</span></div>
	</td>
	<td width="54%" height="32">
		<div align="left">
			<input name="adder" size="20" value="<%=rs("adder")%>" class="smallInput" maxlength="50" >�����ʹ�������û��޷���½��
		</div>
	</td>
</tr>
<tr>
	<td width="46%" height="32" align="right">
		<div align="center"><span class="smallFont">������λ��</span></div>
	</td>
	<td width="54%" height="32">
		<%
		set rstype=createobject("adodb.recordset")
		sql="select * from "& db_EC_Dep_Table &" order by id"
		rstype.Open sql,conn,1,3
		%>
		<select name="depid" >
		<%do while not rstype.EOF%>
		<option value="<%=rstype("id")%>" <%if rs("depid")=Cint(rstype("id")) then%>selected<%end if%> ><%=rstype("E_DepName")%>==<%=rstype("E_DepType")%></option>
			<%
			rstype.MoveNext
		loop
		set rstype=nothing
		%>
		</select>
	</td>
</tr>
<tr>
	<td width="46%" height="32" align="right"><div align="center">�����û�Ȩ�ޣ�</div></td>
	<td width="54%" height="32">
		<select name="purview"  >
		<option <%if rs("purview")=99999 then%> selected <%end if%>value=99999>��������Ա</option>
		<option <%if rs("purview")=1 then%> selected <%end if%>value=1>ϵͳ����Ա</option>
		
		</select>
	</td>
</tr>
<tr>
	<td width="100%" colspan="2" height=32>
		<p align="center">
		<input type="button" value="�� ��" onClick="javascript:history.go(-1)" class="buttonface" >&nbsp;
		<input type="submit" value="�� ��" name="cmdok" class="buttonface" >&nbsp;
		<input type="reset" value="�� ��" name="cmdcancel" class="buttonface" >
	</td>
</tr>
		<%if rs(db_ManageUser_Name)=request.cookies(eChsys)("ManageUserName")  and request.cookies(eChsys)("ManagePurview")="99999" then%>
<tr>
	<td width="100%" height=32 colspan="2" align=center><font color=red>��������Ա <%=request.cookies(eChsys)("ManageUserName")%> ���ã�������������Լ����û��������ڸ��ĺ��˳�ϵͳ�����û������µ�¼���Ա�֤ϵͳ�����������С�</font></td>
</tr>
		<%end if
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		%>
</form>
</table>
</body>
</html>
<!--#include file="Admin_Bottom.asp"-->
	<%else
			Show_Err("�Բ�������Ȩ�޸ĸ��û���Ϣ��<br><br><a href='javascript:history.back()'>����</a>")
		response.end
	end if
	else
		Show_Err("�Բ������ĺ�̨�û�����Ȩ�޲���������")
		response.end
	end if
end if
%>