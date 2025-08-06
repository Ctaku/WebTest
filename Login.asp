<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--'#include file="char.inc"-->
<%
response.cookies(eChsys)("UserName")=""
response.cookies(eChsys)("KEY")=""
response.cookies(eChsys)("purview")=""
response.cookies(eChsys)("fullname")=""
response.cookies(eChsys)("reglevel")=""
response.cookies(eChsys)("shenhe")=""
response.cookies(eChsys)("ViewUrl")="admin_index.asp"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=news.css rel=stylesheet>
<title><%=copyright%><%=version%><%=ver%></title>
<script language="JavaScript">
<!--

if (self != top) top.location.href = window.location.href

//-->
</script>

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

<style type="text/css">
<!--
.style1 {
	font-size: 10.5pt;
	font-weight: bold;
}
-->
</style>
</head>
<body topmargin="0" marginheight="0">
<br>
<p>&nbsp;</p>
<p>&nbsp;</p>
<form method="POST" action="ChkLogin.asp" name="UserLogin" onSubmit="return CheckFormUserLogin();">
  <table width="590" height="260" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td width="77" rowspan="2" valign="top"><img src="Images/login_01.gif" width="77" height="260"></td>
      <td width="74" rowspan="2" valign="top"><img src="Images/login_02.gif" width="74" height="260"></td>
      <td width="43" rowspan="2" valign="top"><img src="Images/login_03.gif" width="85" height="260"></td>
      <td width="354" height="86" valign="top"><img src="Images/login_04.gif" width="354" height="86"></td>
    </tr>
    <tr>
      <td width="354" height="174" valign="top" background="Images/login_05.gif"><table width="85%" border="0" cellspacing="4" cellpadding="0">
        
        <tr>
          <td width="36%" height="19"><div align="center">用户名：</div></td>
          <td width="64%"><input name="UserName"  size="15" font face="宋体"></td>
        </tr>
        <tr>
          <td height="17"><div align="center">密&nbsp;&nbsp;&nbsp;码：</div></td>
          <td><input type="password" name="Passwd" size="15" font face="宋体" ></td>
        </tr>
        <tr>
          <td height="21"><div align="center">验证码： </div></td>
          <td><div align="left">
            <%
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
            <input type="text" name="verifycode" size="8" font face="宋体" >
            <b><span><font color=#000000><%=getcode1()%></font></span></b> </div></td>
        </tr>
        <tr>
          <td height="25"><div align="right">
            <input name="Submit" type=image  src="images/button_dl.gif" height="23"  width="46" style="CURSOR: hand" >
          </div></td>
          <td>&nbsp;&nbsp;&nbsp;
           <IMG style="CURSOR: hand" onclick=reset() src="images/button_cx.gif" height=23  width=46></td></tr>
      </table></td>
    </tr>
	<tr><td colspan="4"><center>
        版权所有：<a href="<%=logourl%>" target="_blank"><font color="#000000">黔东南银监分局</font></a> 
      </center>
	</td>
    </tr>
  </table>
</form>
</body>
</html>