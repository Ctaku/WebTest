<!--#include file="Conn.ASP"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="Config.ASP"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="Function.asp"-->
<%
if not(request.cookies(eChsys)("ManageKEY")="super" or request.cookies(eChsys)("ManageKEY")="typemaster" or request.cookies(eChsys)("ManageKEY")="bigmaster") then
	Show_Err("�Բ������ĺ�̨����Ȩ�޲���������")
	response.end
else
	if shspecialgl="1" or (request.cookies(eChsys)("ManagePurview")="99999" and request.cookies(eChsys)("purview")="99999") then
	Set rs6 = Server.CreateObject("ADODB.Recordset")
	sql6 ="SELECT  * From "& db_EC_Special_Table &" order by SpecialID"
	RS6.open sql6,Conn,3,3
	%>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - ר�����</title>
</head>
<body topmargin="0">

<table width="100%" height="166" border="1" cellpadding="3" cellspacing="0" bordercolor="#FFDFBF" bgcolor="#FFFFFF" id="AutoNumber1" style="border-collapse: collapse">
<tr class="TDtop">
	<td  colspan="4" width="100%" height=25 align="center" valign="middle">��<strong>ר �� �� ��</strong>��</td>
</tr>
<tr align="center"  height=20>
	<td width="10%" >ID��</td>
	<td width="30%" >ר������</td>
	<td width="30%" >ר��ע��</td>
	<td width="30%" >ִ&nbsp;&nbsp;&nbsp;��</td>
</tr>
	<%do while not rs6.eof%>
<tr height=20>
	<td width="10%"  align="center" bgcolor="#FFFFFF"><%=rs6("SpecialID")%>��</td>
	<td width="30%"  align="center" bgcolor="#FFFFFF">
		<span style="CURSOR: hand" title="<%if rs6("specialzs")<>"" then%><%=rs6("specialzs")%><%else%>��<%end if%>"><%=rs6("E_SpecialName")%></span>
		<%if request.cookies(eChsys)("ManageKEY")="super" then%>
		<font color=red>(<%=rs6("specialmaster")%>)</font>
		<%end if%>
	</td>
	<td width="30%" > &nbsp;<%=rs6("Specialzs")%>��</td>
	<td width="30%"  align="center" bgcolor="#FFFFFF">
		<%if (specialgl="1" and request.cookies(eChsys)("ManageKEY")="bigmaster") or (shspecialgl="1" and request.cookies(eChsys)("ManageKEY")="check") or request.cookies(eChsys)("ManageKEY")="super" or request.cookies(eChsys)("ManageKEY")="typemaster" or rs6("specialmaster")=request.cookies(eChsys)("ManageUserName") then%>
		<a href="Specialedit.asp?SpecialID=<%=rs6("SpecialID")%>" onMouseOver="window.status='�༭ר�⡰<%=rs6("E_SpecialName")%>��������';return true;" onMouseOut="window.status='';return true;" title='�༭ר�⡰<%=rs6("E_SpecialName")%>��������'>�༭</a>
		<a href="Specialkill.asp?SpecialID=<%=rs6("SpecialID")%>" onMouseOver="window.status='ɾ��ר�⡰<%=rs6("E_SpecialName")%>��';return true;" onMouseOut="window.status='';return true;" title='ɾ��ר�⡰<%=rs6("E_SpecialName")%>��'>ɾ��</a>
		<%end if%>
	</td>
</tr>
		<%rs6.MoveNext
	Loop
	rs6.close
	set rs6=nothing
	conn.close
	set conn=nothing
	if (specialgl="1" and request.cookies(eChsys)("Managekey")="bigmaster") or (shspecialgl="1" and request.cookies(eChsys)("ManageKEY")="check") or request.cookies(eChsys)("ManageKEY")="super" or request.cookies(eChsys)("ManageKEY")="typemaster" then%>
<tr valign="middle">
	<form method="post" action="SpecialAdd.asp" name="type">
	<td width="100%" height="60" colspan="4" align="center">
		����ר�⣺<textarea name="E_SpecialName" cols="20" rows="3" class="text" style="font-family: ����; font-size: 9pt" title="����������Ҫ��ӵ�ר������" onMouseOver="window.status='����������Ҫ��ӵ�ר������';return true;" onMouseOut="window.status='';return true;"  ></textarea>
		ר��ע�ͣ�<textarea name="specialzs" cols="40" rows="3" class="text" style="font-family: ����; font-size: 9pt" title="����������Ҫ��ӵ�ר��ע��" onMouseOver="window.status='����������Ҫ��ӵ�ר��ע��';return true;" onMouseOut="window.status='';return true;"  ></textarea>
		<input type="hidden" name="Specialmaster" value="<%=request.cookies(eChsys)("ManageUserName")%>">
		<input type="submit" name="Submit" value="���" class=button onMouseOver="window.status='�������ť������ר��';return true;" onMouseOut="window.status='';return true;" title='�������ť������ר��' >
		<input type="reset" value="��д" name="B1" class=button onMouseOver="window.status='�������ť�������ר��';return true;" onMouseOut="window.status='';return true;" title='�������ť�������ר��' >
	</td>
	</form>
</tr>
	<%end if%>
</table>
</body>
</html>
<!--#include file="Admin_Bottom.asp"-->
	<%else
		Show_Err("�Բ��𣬸ù����Ѿ�������ϵͳ����Ա�رգ���û��Ȩ�޲�����")
		response.end
	end if
end if%>