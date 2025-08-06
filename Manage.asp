<!--#include file="conn.asp" -->
<!--#include file="ConnUser.asp" -->
<!--#include file="config.asp" -->
<!--'#include file="char.inc"-->
<!--#include file="ChkUser.asp"-->
<script language=javascript>
function CheckFormUserLogin()
{
	if(document.UserLogin.UserName.value=="")
	{
		alert("请输入用户名！");
		document.UserLogin.UserName.focus();
		return false;
	}
	if(document.UserLogin.Passwd.value == "")
	{
		alert("请输入密码！");
		document.UserLogin.Passwd.focus();
		return false;
	}
	if(document.UserLogin.verifycode.value == "")
	{
		alert("请输入验证码！");
		document.UserLogin.verifycode.focus();
		return false;
	}
}
</script>
<%
response.cookies(eChsys)("ManageUserName")=""
response.cookies(eChsys)("ManageKEY")=""
response.cookies(eChsys)("ManagePurview")=""
response.cookies(eChsys)("ManageFullName")=""
response.cookies(eChsys)("ManageReglevel")=""
showerr=request("showmess")

Function getcode1()
	Dim test
	On Error Resume Next
	Set test=Server.CreateObject("Adodb.Stream")
	Set test=Nothing
	If Err Then
		Dim zNum
		Randomize timer
		zNum = cint(8999*Rnd+1000)
		Session("verifycode") = zNum
		getcode1= Session("verifycode")		
	Else
		getcode1= "<img src=""getcode.asp"">"		
	End If
End Function
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%><%=ver%></title>

</head>
<body bgcolor="#ffffff" topmargin="0" marginheight="0">
<table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#FFDFBF" bgcolor="#FFFFFF" id="AutoNumber1" style="border-collapse: collapse">
<form method="POST" action="ChkManageLogin.asp" name="UserLogin" onSubmit="return CheckFormUserLogin();">
<tr >
	<td height="30" colspan="2" align="center" background="images/admin_topbg.gif" class="title">
		[<%=copyright%>]&nbsp;<%=version%>&nbsp;<%=ver%>
	</td>
</tr>
<tr>
	<td height="30" colspan="2" align="center" background="images/admin_topbg.gif" class="title">| 管 理  登 录 |</td>
</tr>
<tr>
	<td width="45%" align="right">用户名：</td>
	<td width="55%" align="left">
		<input name="UserName" size="15" font face="宋体" >
	</td>
</tr>
<tr>
	<td align="right">密&nbsp;&nbsp;码：</td>
	<td align="left">
		<input type="password" name="Passwd" size="15" font face="宋体" >
	</td>
</tr>
<tr>
	<td align="right">验证码：</td>
	<td align="left">
		<input type="text" name="verifycode" size="15" font face="宋体" >
		<b><span><font color=#000000><%=getcode1()%></font></span></b>
	</td>
</tr>
<%if showerr<>"" then%>
<tr>
	<td colspan="2" align="center">
		<p><font size=3 color=red><%=showerr%></font></p>
	</td>
</tr>
<%end if%>
<tr>
	<td colspan="2" align="center">
		<p>
			<input type="submit" name="Submit" value="确定"  title="确定">&nbsp;
			<input type="reset" name="Submit2" value="重置"  title="重输">
		</p>
	</td>
</tr>
<tr >
	<td height="30" colspan="2" align="center" background="images/admin_topbg.gif">&nbsp;</td>
</tr>
</form>
</table>
</body>
</html>
<%response.end%>