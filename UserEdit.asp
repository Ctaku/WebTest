<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp" -->
<%
if not(request.cookies(eChsys)("ManageKEY")="super" or request.cookies(eChsys)("ManageKEY")="typemaster" or request.cookies(eChsys)("ManageKEY")="bigmaster") then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	if request.cookies(eChsys)("purview")="99999" and request.cookies(eChsys)("KEY")="super" then
		%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 修改资料</title>
<script LANGUAGE="javascript">
<!--
function FrmAddLink_onsubmit() {
 if (document.FrmAddLink.passwd.value=="")
	{
	  alert("对不起，请您输入密码！")
	  document.FrmAddLink.passwd.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd.value.length < 4)
	{
	  alert("为了安全，您的密码应该长一点！")
	  document.FrmAddLink.passwd.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd.value.length > 16)
	{
	  alert("您的密码太长了吧！")
	  document.FrmAddLink.passwd.focus()
	  return false
	 }
else if (document.FrmAddLink.username.value==document.FrmAddLink.passwd.value)
	{
	  alert("为了安全，用户名与密码不应该相同！")
	  document.FrmAddLink.passwd.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd2.value=="")
	{
	  alert("对不起，请您输入验证密码！")
	  document.FrmAddLink.passwd2.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd2.value !== document.FrmAddLink.passwd.value)
	{
	  alert("对不起，您两次输入的密码不一致！")
	  document.FrmAddLink.passwd2.focus()
	  return false
	 }
else if (document.FrmAddLink.fullname.value=="")
	{
	  alert("对不起，请输入真实姓名！")
	  document.FrmAddLink.fullname.focus()
	  return false
	 }
else if (document.FrmAddLink.depid.value=="")
	{
	  alert("对不起，请选择工作单位！")
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
		sql="select * from "& db_User_Table &" where "& db_User_ID &"="&ChkRequest(request("id"),1)	'防注入
		set rs=server.createobject("adodb.recordset")
		rs.open sql,ConnUser,1,1
		if rs("adder")=request.cookies(eChsys)("username") or request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="typemanage" then%> 
<table border="1" width="100%" cellspacing="0" cellpadding="0"   bgcolor="#FFFFFF"  style="border-collapse: collapse"  bordercolor="#FF6600" align=center>
<form method="POST" name="FrmAddLink" LANGUAGE="javascript" onSubmit="return FrmAddLink_onsubmit()"   action="SaveEdit.asp?id=<%=request("id")%>">
<tr>
	<td width="100%" height=32 colspan="2" class="TDtop">
		<p align="center" >┊ <strong>用户资料修改</strong> ┊
	</td>
</tr>
<tr>
	<td width="46%" height="32" align="right" valign="middle"> 
		<div align="center"><span class="smallFont">用 户 名：</span></div>
	</td>
	<td width="54%" height="32"> 
		<div align="left">
			<input name="username" size="20" value="<%=rs(db_User_Name)%>" <%if (rs(db_User_Name)<>request.cookies(eChsys)("username")  and request.cookies(eChsys)("purview")="99999") or request.cookies(eChsys)("purview")<>"99999" then%>readonly<%end if%> class="smallInput" maxlength="30" >
		</div>
	</td>
	</tr>
<tr>
	<td width="46%" height="32" align="right"> 
		<div align="center"><span class="smallFont">密&nbsp;&nbsp;&nbsp;码：</span></div>
	</td>
	<td width="54%" height="32"> 
		<div align="left"> 
			<input name="passwd" type="password" size="20" value="<%=rs(db_User_Password)%>" class="smallInput" maxlength="50" >
		</div>
	</td>
</tr>
<tr>
	<td width="46%" height="32" align="right"> 
		<div align="center"><span class="smallFont">验证密码：</span></div>
	</td>
	<td width="54%" height="32"> 
		<div align="left"> 
			<input name="passwd2" type="password" size="20" value="<%=rs(db_User_Password)%>" class="smallInput" maxlength="50" >
		</div>
	</td>
</tr>
<tr>
	<td width="46%" height="32" align="right"> 
		<div align="center"><span class="smallFont">真实姓名：</span></div>
	</td>
	<td width="54%" height="32"> 
		<div align="left"> 
			<input name="fullname" size="20" value="<%=rs("fullname")%>" class="smallInput" maxlength="50" >
		</div>
	</td>
</tr>
<tr>
	<td width="46%" height="32" align="right"> 
		<div align="center"><span class="smallFont">工作单位：</span></div>
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
	<td width="46%" height="32" align="right"><div align="center">用户权限：</div></td>
	<td width="54%" height="32">
		<select name="oskey" >
			<%if request.cookies(eChsys)("key")="super" then%>
			<option <%if rs("oskey")="super" then%> selected <%end if%>value="super">系统用户</option>
			<option <%if rs("oskey")="typemaster" then%> selected <%end if%>value="typemaster">总栏用户</option>
			<option <%if rs("oskey")="bigmaster" then%> selected <%end if%>value="bigmaster">大类用户</option>
			<option <%if rs("oskey")="smallmaster" then%> selected <%end if%>value="smallmaster">小类用户</option>
			<option <%if rs("oskey")="check" then%> selected <%end if%>value="check">审核用户</option>
			<option <%if rs("oskey")="selfreg" then%> selected <%end if%>value="selfreg">注册用户</option>
			<%else%>
				<%if request.cookies(eChsys)("key")="typemaster" then%>
			<option <%if rs("oskey")="bigmaster" then%> selected <%end if%>value="bigmaster">大类用户</option>
			<option <%if rs("oskey")="smallmaster" then%> selected <%end if%>value="smallmaster">小类用户</option>
			<option <%if rs("oskey")="selfreg" then%> selected <%end if%>value="selfreg">注册用户</option>
				<%end if%>
			<%end if%>
		</select>
	</td>
</tr>
<%if request.cookies(eChsys)("key")="super" and Request.cookies(eChsys)("purview")="99999" And rs("oskey")="super" then%>
<tr> 
	<td height="32" align="right"> 
		<div align="center">系统管理员是否为超级管理员：<br>该功能仅对系统管理员有效！</div>
	</td>
		<td height="32"> 该用户
		<select name="ifpurview" >
		<option <%if rs("purview")="1" then%> selected <%end if%> value="1">不是</option>
		<option <%if rs("purview")="99999" then%> selected <%end if%> value="99999">是</option>
		</select> 超级管理员
	</td>
</tr>
<%end if%>

<%if rs("oskey")="smallmaster" then%>
<tr>
	<td width="46%" height="32" align="right"> 
		<div align="center">小类管理员文章默认状态：<br>该功能仅对小类用户有效！</div>
	</td>
	<td width="54%" height="32">
		<select name="shenhe" >
			<option <%if rs("shenhe")="0" then%> selected <%end if%> value="0">已审核</option>
			<option <%if rs("shenhe")="1" then%> selected <%end if%> value="1">未审核</option>
		</select>默认值为需要审核
	</td>
</tr>
<%End if%>

<tr>
	<td width="100%" colspan="2" height=32>
		<p align="center">
			<input type="reset" value=" 返 回 " onClick="javascript:history.go(-1)" class="buttonface" >&nbsp;
			<input type="submit" value=" 修 改 " name="cmdok" class="buttonface" >&nbsp;
			<input type="reset" value=" 复 原 " name="cmdcancel" class="buttonface" >
		</p>
	</td>
</tr>
			<%if rs(db_User_Name)=request.cookies(eChsys)("username") and request.cookies(eChsys)("purview")="99999" then%>
<tr>
	<td width="100%" height=32 colspan="2" align=center>
		<font color=red>超级用户 <%=request.cookies(eChsys)("username")%> 您好，如果您更改了自己的用户名，请在更改后退出系统以新用户名重新登录，以保证系统功能正常运行。</font>
	</td>
</tr>
			<%end if%>
</form>
</table>
</body>
</html>
			<%rs.close 
			set rs=nothing
			ConnUser.close
			set ConnUser=nothing
			conn.close
			set conn=nothing
		else
			Show_Err("对不起，你无权修改该用户信息！<br><br>－－－<a href='javascript:history.back()'>回去重来</a>－－－")
			response.end
		end if
	else
		Show_Err("对不起，您的前台用户管理权限不够操作！")
		response.end
	end if
end if%>