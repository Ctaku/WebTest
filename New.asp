<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--'#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<%
if not(request.cookies(eChsys)("ManageKEY")="super" or request.cookies(eChsys)("KEY")="super") then
	Show_Err("�Բ������ĺ�̨����Ȩ�޲���������")
	response.end
else

	Set rs9 = Server.CreateObject("ADODB.Recordset")
	sql9 ="SELECT * From "& db_EC_System_Table &" Order By id DESC"
	RS9.open sql9,Conn,1,1

	if rs9("init")="1" or (request.cookies(eChsys)("ManagePurview")="99999" and request.cookies(eChsys)("purview")="99999") then
	%>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - ϵͳ��ʼ��</title>
<script language="JavaScript">
<!--
function CheckAll(form){
	for (var i=0;i<form.elements.length;i++) {
		var e = form.elements[i];
		if (e.name != 'chkall') e.checked = form.chkall.checked; 
	}
}
//-->
</script>
</head>
<body topmargin="0">

<table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#FFDFBF" bgcolor="#FFFFFF" id="AutoNumber1" style="border-collapse: collapse">
<form action=New_submit.asp method=post>
<tr>
	<td class="TDtop" width="100%" height=25 align="center" bgcolor="#ffb100" valign="middle">��ϵͳ��ʼ��ȷ�ϩ�</td>
</tr>
<tr>
	<td height="30" align="center" valign="middle"></td>
</tr>
<tr>
	<td align="center">
		<strong>��   �棡</strong><br><br>
		<strong>��ʼ��һ����ϵͳ��һ�ΰ�װʱ���У�</strong><br>
		<strong>��ϵͳ���������ڼ��ʼ�������ܶ�ʧ���ݡ�</strong><br><br>
ϵͳ������ʼ������ɾ�����ݿ��е�������Ϣ�����Իᱣ�����ݿ�ṹ��ȷ��Ҫ��ϵͳ��<font color=red>��ȫ��ʼ��</font>�������밴��ȷ�����������밴��ȡ������
	</td>
</tr>
<tr>
<td height="30"></td>
</tr>
<tr>
	<td width="100%" align=center > 
		<input type=submit value="ȷ��" name="alert_button" class="Bottom1" >&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" value="ȡ��" name="cancel" class="Botton" onClick="javascript:history.go(-1)" >
	</td>
</tr>
<tr>
	<td height="25" align="center" valign="middle"></td>
</tr>
</form>
</table>
<table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#FFDFBF" bgcolor="#FFFFFF" id="AutoNumber1" style="border-collapse: collapse">
<form action=New_submit.asp method=post>
<tr>
	<td width="100%" height="25" colspan="4" class="TDtop">��Ҳ����ѡ�񵥶���ʼ��ϵͳ�Ĳ���</td>
</tr>
<tr>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_EC_Dep_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">��&nbsp;&nbsp;&nbsp;&nbsp;��
	<%rs=conn.execute("Select count(*) from "& db_EC_Dep_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_EC_Type_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">��&nbsp;&nbsp;&nbsp;&nbsp;��
	<%rs=conn.execute("Select count(*) from "& db_EC_Type_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_EC_BigClass_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">��&nbsp;&nbsp;&nbsp;&nbsp;��
	<%rs=conn.execute("Select count(*) from "& db_EC_BigClass_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_EC_SmallClass_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">С&nbsp;&nbsp;&nbsp;&nbsp;��
	<%rs=conn.execute("Select count(*) from "& db_EC_SmallClass_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
</tr>
<tr>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_EC_Special_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">ר&nbsp;&nbsp;&nbsp;&nbsp;��
	<%rs=conn.execute("Select count(*) from "& db_EC_Special_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_EC_News_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">��&nbsp;&nbsp;&nbsp;&nbsp;��
	<%rs=conn.execute("Select count(*) from "& db_EC_News_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_NewsFile_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">��&nbsp;&nbsp;&nbsp;&nbsp;ʱ
	<%rs=conn.execute("Select count(*) from "& db_NewsFile_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_UploadPic_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">�ϴ���ʱ
	<%rs=conn.execute("Select count(*) from "& db_UploadPic_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
</tr>
<tr>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_EC_Vote_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">Ͷ&nbsp;&nbsp;&nbsp;&nbsp;Ʊ
	<%rs=conn.execute("Select count(*) from "& db_EC_Vote_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_EC_Link_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">��������
	<%rs=conn.execute("Select count(*) from "& db_EC_Link_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_EC_Board_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">��&nbsp;&nbsp;&nbsp;&nbsp;��
	<%rs=conn.execute("Select count(*) from "& db_EC_Board_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_EC_Attach_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">��&nbsp;&nbsp;&nbsp;&nbsp;��
	<%rs=conn.execute("Select count(*) from "& db_EC_Attach_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
</tr>
<tr>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_EC_Review_Table%>0" name="selectdel" class="Bottom1" style="background-color: #ffffff;">��&nbsp;&nbsp;&nbsp;&nbsp;��
	<%rs=conn.execute("Select count(*) from "& db_EC_Review_Table &" where NewsId<1")
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_EC_Review_Table%>1" name="selectdel" class="Bottom1" style="background-color: #ffffff;">��&nbsp;&nbsp;&nbsp;&nbsp;��
	<%rs=conn.execute("Select count(*) from "& db_EC_Review_Table &" where NewsId>0")
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_User_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">��ͨ�û�
	<%rs=ConnUser.execute("Select count(*) from "& db_User_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_EC_Manage_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">�����û�
	<%rs=conn.execute("Select count(*) from "& db_EC_Manage_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
</tr>
<tr>
<td colspan=4  height=25 align="center">
	<input type=hidden name=action value='del'>
	<input type=checkbox name=chkall value=on onClick="CheckAll(this.form)" style=" background-color: #ffffff;">ȫѡ&nbsp;&nbsp;
	<input type=submit name=action1 onClick="{if(confirm('��ʼ��ѡ������Ŀ��')){this.document.check;return true;}return false;}" value="��ʼ��" >
</td>
</tr>
</form>
</table>
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#FFDFBF" bgcolor="#FFFFFF" id="AutoNumber1" style="border-collapse: collapse">
<%
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from EC_counters"
rs.open sql,conn,1,3
if rs.eof and rs.bof then
	rs.addnew
end if
if request.form("Ch_Count")=1 then
	if request.form("total")="" then
		rs("total")=0
	else
		rs("total")=request.form("total")
	end if
	if request.form("today")="" then
		rs("today")=0
	else
		rs("today")=request.form("today")
	end if
	if request.form("yesterday")="" then
		rs("yesterday")=0
	else
		rs("yesterday")=request.form("yesterday")
	end if
	if request.form("months")="" then
		rs("month")=0
	else
		rs("month")=request.form("months")
	end if
	if request.form("bmonth")="" then
		rs("bmonth")=0
	else
		rs("bmonth")=request.form("bmonth")
	end if
	rs("date")=now()
	rs("inputdate")=now()
	rs("lastip")="127.0.0.1"
	rs.update
end if
%>
<form method="POST" action="New.asp?ID=<&=ID&>" id=form name=form>
<tr class="TDTop">
	<td width="100%" height="25" colspan=5 align="center"><b>ͳ�����ݳ�ʼ��</b></td>
</tr>
<tr> 
	<td width="20%" align="center">��������</td>
	<td width="20%" align="center">���շ���</td>
	<td width="20%" align="center">���շ���</td>
	<td width="20%" align="center">���·���</td>
	<td width="20%" align="center">���·���</td>
</tr>
<tr align="center">
	<td height="25"><input type="text" name="total" value="<%=rs("total")%>" size="9"></td>
	<td height="25"><input type="text" name="today" value="<%=rs("today")%>" size="9"></td>
	<td height="25"><input type="text" name="yesterday" value="<%=rs("yesterday")%>" size="9"></td>
	<td height="25"><input type="text" name="months" value="<%=rs("month")%>" size="9"></td>
	<td height="25"><input type="text" name="bmonth" value="<%=rs("bmonth")%>" size="9"></td>
</tr>
<tr align="center"> 
	<td height="25" colspan=5>
		<input type="hidden" name="Ch_Count" value="1">
		<input type="reset" value="�� ��" name="cmdcance2l" >&nbsp;&nbsp;
		<input type="submit" value="�� ��" name="cmdok2" >
	</td>
</tr>
</form>
<%             
	set rs=nothing             
%>
</table>
</body>
</html>
<!--#include file="Admin_Bottom.asp"-->
	<%
	rs9.close
	set rs9=nothing
	conn.close
	set conn=nothing
	
	else
		Show_Err("�Բ��𣬸ù����Ѿ�������ϵͳ����Ա�رգ���û��Ȩ�޲�����")
		response.end
	end if
end if%>