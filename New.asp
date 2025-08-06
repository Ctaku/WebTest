<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--'#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<%
if not(request.cookies(eChsys)("ManageKEY")="super" or request.cookies(eChsys)("KEY")="super") then
	Show_Err("对不起，您的后台管理权限不够操作！")
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
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 系统初始化</title>
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
	<td class="TDtop" width="100%" height=25 align="center" bgcolor="#ffb100" valign="middle">┊系统初始化确认┊</td>
</tr>
<tr>
	<td height="30" align="center" valign="middle"></td>
</tr>
<tr>
	<td align="center">
		<strong>警   告！</strong><br><br>
		<strong>初始化一般在系统第一次安装时进行！</strong><br>
		<strong>在系统正常运行期间初始化将可能丢失数据。</strong><br><br>
系统将被初始化，会删除数据库中的所有信息，但仍会保留数据库结构，确认要对系统做<font color=red>完全初始化</font>操作，请按“确定”，否则请按“取消”！
	</td>
</tr>
<tr>
<td height="30"></td>
</tr>
<tr>
	<td width="100%" align=center > 
		<input type=submit value="确定" name="alert_button" class="Bottom1" >&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" value="取消" name="cancel" class="Botton" onClick="javascript:history.go(-1)" >
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
	<td width="100%" height="25" colspan="4" class="TDtop">你也可以选择单独初始化系统的部分</td>
</tr>
<tr>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_EC_Dep_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">部&nbsp;&nbsp;&nbsp;&nbsp;门
	<%rs=conn.execute("Select count(*) from "& db_EC_Dep_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_EC_Type_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">总&nbsp;&nbsp;&nbsp;&nbsp;栏
	<%rs=conn.execute("Select count(*) from "& db_EC_Type_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_EC_BigClass_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">大&nbsp;&nbsp;&nbsp;&nbsp;栏
	<%rs=conn.execute("Select count(*) from "& db_EC_BigClass_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_EC_SmallClass_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">小&nbsp;&nbsp;&nbsp;&nbsp;栏
	<%rs=conn.execute("Select count(*) from "& db_EC_SmallClass_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
</tr>
<tr>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_EC_Special_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">专&nbsp;&nbsp;&nbsp;&nbsp;题
	<%rs=conn.execute("Select count(*) from "& db_EC_Special_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_EC_News_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">文&nbsp;&nbsp;&nbsp;&nbsp;章
	<%rs=conn.execute("Select count(*) from "& db_EC_News_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_NewsFile_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">临&nbsp;&nbsp;&nbsp;&nbsp;时
	<%rs=conn.execute("Select count(*) from "& db_NewsFile_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_UploadPic_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">上传临时
	<%rs=conn.execute("Select count(*) from "& db_UploadPic_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
</tr>
<tr>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_EC_Vote_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">投&nbsp;&nbsp;&nbsp;&nbsp;票
	<%rs=conn.execute("Select count(*) from "& db_EC_Vote_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_EC_Link_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">友情链接
	<%rs=conn.execute("Select count(*) from "& db_EC_Link_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_EC_Board_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">公&nbsp;&nbsp;&nbsp;&nbsp;告
	<%rs=conn.execute("Select count(*) from "& db_EC_Board_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_EC_Attach_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">附&nbsp;&nbsp;&nbsp;&nbsp;件
	<%rs=conn.execute("Select count(*) from "& db_EC_Attach_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
</tr>
<tr>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_EC_Review_Table%>0" name="selectdel" class="Bottom1" style="background-color: #ffffff;">留&nbsp;&nbsp;&nbsp;&nbsp;言
	<%rs=conn.execute("Select count(*) from "& db_EC_Review_Table &" where NewsId<1")
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_EC_Review_Table%>1" name="selectdel" class="Bottom1" style="background-color: #ffffff;">评&nbsp;&nbsp;&nbsp;&nbsp;论
	<%rs=conn.execute("Select count(*) from "& db_EC_Review_Table &" where NewsId>0")
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_User_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">普通用户
	<%rs=ConnUser.execute("Select count(*) from "& db_User_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
	<td width="25%"><p align="center">
		<input type=checkbox value="<%=db_EC_Manage_Table%>" name="selectdel" class="Bottom1" style="background-color: #ffffff;">管理用户
	<%rs=conn.execute("Select count(*) from "& db_EC_Manage_Table)
		response.write "&nbsp;&nbsp;<font color=red>["& rs(0) &"]</font>"
		set rs=nothing%>
	</td>
</tr>
<tr>
<td colspan=4  height=25 align="center">
	<input type=hidden name=action value='del'>
	<input type=checkbox name=chkall value=on onClick="CheckAll(this.form)" style=" background-color: #ffffff;">全选&nbsp;&nbsp;
	<input type=submit name=action1 onClick="{if(confirm('初始化选定的项目吗？')){this.document.check;return true;}return false;}" value="初始化" >
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
	<td width="100%" height="25" colspan=5 align="center"><b>统计数据初始化</b></td>
</tr>
<tr> 
	<td width="20%" align="center">访问总数</td>
	<td width="20%" align="center">今日访问</td>
	<td width="20%" align="center">昨日访问</td>
	<td width="20%" align="center">本月访问</td>
	<td width="20%" align="center">上月访问</td>
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
		<input type="reset" value="重 置" name="cmdcance2l" >&nbsp;&nbsp;
		<input type="submit" value="修 改" name="cmdok2" >
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
		Show_Err("对不起，该功能已经被超级系统管理员关闭，您没有权限操作！")
		response.end
	end if
end if%>