<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="ChkUser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="Function.asp"-->
<%
IF request.cookies(eChsys)("ManageKEY")<>"super" then
	Show_Err("对不起，您的后台管理权限不够！")
	response.end
else
	if depmanage="1" or (request.cookies(eChsys)("ManagePurview")="99999" and request.cookies(eChsys)("purview")="99999") then
    dim E_DepType
	E_DepType=ChkRequest(request("E_DepType"),0)	'防注入
	Set rs6 = Server.CreateObject("ADODB.Recordset")
	if E_DepType="" then
		sql6 ="select * from "& db_EC_Dep_Table &" order by ID"
	else
		sql6 ="select * from "& db_EC_Dep_Table &" where E_DepType="&E_DepType&" order by ID"
	end if
	rs6.open sql6,Conn,3,3
	%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 单位部门管理</title>
</head>
<body topmargin="0">

<center>
<table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#FFDFBF" bgcolor="#FFFFFF" id="AutoNumber1" style="border-collapse: collapse">
<tr class="TDtop"> 
	<td  colspan="5" width="100%" height=25 align="center" valign="middle" height="24">
		┊ <strong>单位部门管理</strong> ┊
	</td>
</tr>

<tr align="center" class="TDtop1"> 
	<td width="3%" >原序</td>
	<td width="10%" >单位名称</td>
	<td width="6%" height="25" >部门名称</td>
	<td width="4%" >累计稿件</td>
	<td width="10%" >单位编辑</td>
</tr>
	<%do while not rs6.eof%>
<tr> 
	<td width="3%" align="center" bgcolor="#FFFFFF"><%=rs6("ID")%></td>
	<td width="10%" align="center" bgcolor="#FFFFFF"><%=nohtml(rs6("E_DepName"))%></span></td>
	<td width="6%" height="25" align="center" bgcolor="#FFFFFF"> 
		<%E_DepType=nohtml(rs6("E_DepType"))
		response.write ""& E_DepType &""
		%>
	</td>
	<td width="4%" align="center" bgcolor="#FFFFFF"><%=rs6("depnumber")%></td>
	<td width="10%" bgcolor="#FFFFFF" align="center">
		<a href="E_DepEdit.asp?ID=<%=rs6("ID")%>&E_DepName=<%=nohtml(rs6("E_DepName"))%>&E_DepType=<%=nohtml(rs6("E_DepType"))%>" onMouseOver="window.status='编辑单位“<%=nohtml(rs6("E_DepName"))%>”的属性';return true;" onMouseOut="window.status='';return true;" title='编辑单位“<%=nohtml(rs6("E_DepName"))%>”的属性'>编辑</a>&nbsp;
		<a href="E_DepKill.asp?ID=<%=rs6("ID")%>&E_DepName=<%=nohtml(rs6("E_DepName"))%>" onMouseOver="window.status='删除单位“<%=nohtml(rs6("E_DepName"))%>”';return true;" onMouseOut="window.status='';return true;" title='删除单位“<%=nohtml(rs6("E_DepName"))%>”'>删除</a>
	</td>
</tr>
		<%
		rs6.MoveNext
	loop
	rs6.close
	set rs6=nothing
	%>
<tr> 
	<td colspan="5" height="40" align="center" width="100%" bgcolor="#FFFFFF">　</td>
</tr>
<form method="post" action="E_DepAdd.asp" name="name">
<tr> 
	<td align="center" colspan="5" width="100%"  height="34"> 
		增加单位：<input class=text type="text" name="E_DepName" size="15" onMouseOver="window.status='在这里输入要增加的单位名称';return true;" onMouseOut="window.status='';return true;" title="在这里输入要增加的单位名称" >&nbsp;
		单位部门：<input class=text type="text" name="E_DepType" size="15" onMouseOver="window.status='在这里输入要增加的单位部门';return true;" onMouseOut="window.status='';return true;" title="在这里输入要增加的部门名称" >&nbsp;

		<input type="submit" name="Submit" value="添加" title="按这个按钮添加这个单位" onMouseOver="window.status='按这个按钮重新添加小类';return true;" onMouseOut="window.status='';return true;" >
		<input type="reset" value="重写" name="B1" title="按这个按钮重新添加单位" onMouseOver="window.status='按这个按钮重新添加单位';return true;" onMouseOut="window.status='';return true;" >
	</td>
</tr>

</form>

</table>
</center>
</body>
</html>
<!--#include file="Admin_Bottom.asp"-->
	<%else
		Show_Err("对不起，该功能已经被超级系统管理员关闭，您没有权限操作！")
		response.end
	end if
end if%>