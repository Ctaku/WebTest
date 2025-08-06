<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<%
IF request.cookies(eChsys)("KEY")="" THEN
	Show_Err("对不起，您权限不够！")
	response.end
else
	usernamecookie=CheckStr(request.cookies(eChsys)("UserName"))
	passwdcookie=replace(trim(Request.cookies(eChsys)("passwd")),"'","''")
	KEYcookie=replace(trim(request.cookies(eChsys)("KEY")),"'","''")
	%>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 个人资料修改</title>
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
else if (document.FrmAddLink.question.value=="")
	{
	  alert("对不起，请您输入提示问题！")
	  document.FrmAddLink.question.focus()
	  return false
	 }
else if (document.FrmAddLink.answer.value=="")
	{
	  alert("对不起，请您输入问题答案！")
	  document.FrmAddLink.answer.focus()
	  return false
	 }
else if (document.FrmAddLink.question.value==document.FrmAddLink.answer.value)
	{
	  alert("为了安全，提示问题与问题答案不应该相同！")
	  document.FrmAddLink.answer.focus()
	  return false
	 }
else if (document.FrmAddLink.fullname.value=="")
	{
	  alert("对不起，请输入您的真实姓名！")
	  document.FrmAddLink.fullname.focus()
	  return false
	 }
else if (document.FrmAddLink.depid.value=="")
	{
	  alert("对不起，请输入您的工作单位！")
	  document.FrmAddLink.depid.focus()
	  return false
	 }
else if (document.FrmAddLink.sex.value=="")
	{
	  alert("对不起，请输入您的性别！")
	  document.FrmAddLink.sex.focus()
	  return false
	 }
else if (document.FrmAddLink.tel.value=="")
	{
	  alert("对不起，请输入您的联系电话！")
	  document.FrmAddLink.tel.focus()
	  return false
	 }
else if (document.FrmAddLink.email.value=="")
	{
	  alert("对不起，请输入您的电子邮件！")
	  document.FrmAddLink.email.focus()
	  return false
	 }
else if (document.FrmAddLink.email.value.indexOf("@",0)== -1||document.FrmAddLink.email.value.indexOf(".",0)==-1)
	  {
	  alert("对不起，您输入的电子邮件有误！")
	  document.FrmAddLink.email.focus()
	  return false
	 }
}
//-->
</script>
<!--生日选择日期处理开始-->
<SCRIPT language=JavaScript src="inc/User_Info_Modify.js"></SCRIPT>
<!--生日选择日期处理结束-->
</head>
<body topmargin="0">

<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFDFBF" width="100%" id="AutoNumber1">
<form method="POST" name="FrmAddLink" LANGUAGE="javascript" onSubmit="return FrmAddLink_onsubmit()"  action="Save.asp" >
	<%dim logins
	dim sql
	dim rs
	sql="select * from "& db_User_Table &" where "& db_User_Name &"='"&usernamecookie&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,ConnUser,1,1
	logins=rs(db_User_LoginTimes)
	if UserTableType = "Dvbbs" then
		photowidth=rs(db_User_FaceWidth)		''取论坛设定的图片宽度
		photoheight=rs(db_User_FaceHeight)		''取论坛设定的图片高度
	end if
	%> 
<tr>
	<td class="TDtop" colspan="2" height=25>
		<p align="center" >
			<%if logins=1 then%>
			<font color=red>这是您首次使用本系统，为了安全，请立即修改您的个人资料后再执行管理操作。</font><br>
			<%end if%>
			<%=request.cookies(eChsys)("fullname")%>修改资料
		</p>
	</td>
</tr>
<tr>
	<td width="37%"><p align="right">用 户 名：</td>
	<td width="63%">
		<input name="username" size="30" value="<%=rs(db_User_Name)%>" readonly class="smallInput" maxlength="30" >
	</td>
</tr>
<tr>
	<td><p align="right">密&nbsp;&nbsp;&nbsp; 码：</td>
	<td>
		<input name="passwd" type="password" size="30" value="<%=rs(db_User_Password)%>" class="smallInput" maxlength="50"  <%if request.cookies(eChsys)("key")="smallmaster" and inputmodpwd="0" then%>readonly<%end if%>><%if request.cookies(eChsys)("key")="smallmaster" and inputmodpwd="0" then%>该用户不允许修改密码<%end if%>
	</td>
</tr>
<tr>
	<td><p align="right">验证密码：</td>
	<td>
		<input name="passwd2" type="password" size="30" value="<%=rs(db_User_Password)%>" class="smallInput" maxlength="50"  <%if request.cookies(eChsys)("key")="smallmaster" and inputmodpwd="0" then%>readonly<%end if%>><%if request.cookies(eChsys)("key")="smallmaster" and inputmodpwd="0" then%>该用户不允许修改密码<%end if%>
	</td>
</tr>
<tr>
	<td><p align="right">提示问题：</td>
	<td>
		<input name="question" type="text" size="30" value="<%=rs(db_User_Question)%>" class="smallInput"  >
	</td>
</tr>
<tr>
	<td><p align="right">问题答案：</td>
	<td>
		<input name="answer" type="text" size="30" value="<%=rs(db_User_Answer)%>" class="smallInput" >
	</td>
</tr>
<tr>
	<td><p align="right">真实姓名：</td>
	<td>
		<input name="fullname" type="text" size="30" value="<%=rs("fullname")%>" class="smallInput" >
	</td>
</tr>
<tr>
	<td><p align="right">工作单位：</td>
	<td>
	<%
	set rstype=createobject("adodb.recordset")
	sql="select * from "& db_EC_Dep_Table &" order by id"
	rstype.Open sql,conn,1,3
	%>
		<select name="depid" >
			<option value="">请选择工作单位</option>
	<%do while not rstype.EOF%>
			<option value="<%=rstype("id")%>" <%if rs("depid")=cint(rstype("id")) then%> selected <%end if%>><%=rstype("E_DepName")%>==<%=rstype("E_DepType")%></option>
		<%
		rstype.MoveNext
	loop
	set rstype=nothing
	%>
		</select>
	</td>
</tr>
<tr>
	<td><p align="right">性&nbsp;&nbsp;&nbsp; 别：</td>
	<td> 
	<%if db_Sex_Select= "EChuang"	then %>
		<select size="1" name="sex"  >
			<option <%if rs(db_User_sex)="先生" then%>selected<%end if%> value="先生">先生</option>
			<option <%if rs(db_User_sex)="女士" then%>selected<%end if%> value="女士">女士</option>
			<option <%if rs(db_User_sex)="保密" then%>selected<%end if%> value="保密">保密</option>
		</select>
	<%else%>
		<%if  db_Sex_Select="Number" then%>
		<select size="1" name="sex"  >
			<option <%if rs(db_User_Sex)=1 then%>selected<%end if%> value="1">先生</option>
			<option <%if rs(db_User_Sex)=0 then%>selected<%end if%> value="0">女士</option>
		</select>
		<%end if%>
	<%end if%>
	</td>
</tr>
<tr>
	<td><p align="right">生&nbsp;&nbsp;&nbsp; 日：</td>
	<td > 
		<div align="left"> 
	<%if db_Birthday_Select= "EChuang" then %>
			<select size="1" name="birthyear" >
		<%for i=1950 to 2000%>
				<option <%if rs(db_User_birthyear)=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
			</select>年
			<select size="1" name="birthmonth" >
		<%for i=1 to 12%>
				<option <%if rs(db_User_birthmonth)=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
			</select>月
			<select size="1" name="birthday" >
		<%for i=1 to 31%>
				<option <%if rs(db_User_birthday)=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
			</select>日
	<%else
		if db_Birthday_Select="Text" then %>
			<INPUT  name=birthday onFocus="show_cele_date(birthday,'','',birthday)" value="<%=rs(db_User_birthday)%>" class="smallInput" ></div>
		<%end if%>
	<%end if%>
	</td>
</tr>
<tr>
	<td><p align="right">联系电话：</td>
	<td valign="middle"> 
		<input name="tel" size="30" class="smallInput" value="<%=rs("tel")%>" maxlength="100"  >
	</td>
</tr>
<tr>
	<td><p align="right">电子邮件：</td>
	<td>
		<input name="email" type="text" size="30" value="<%=rs(db_User_Email)%>" class="smallInput"   >
	</td>
</tr>
<tr>
	<td><p align="right">个人照片：</td>
	<td>
		<iframe name="ad" frameborder=0 width=100% height=20 scrolling=no src='UploadFace.asp'></iframe>
		<input name="photo" type="text" size="35" value="<%=rs(db_User_Face)%>" class="smallInput"  title="个人照片。您可以上传自己的照片，也可以直接填写您的网上照片的地址。"><br>
    用户头像:
	<%if UserTableType = "EChuang" then
	         if rs(db_User_Face)<>"" then%>
				<img src="<%=rs(db_User_Face)%>" border="0">
			<%else%>
				<img src="images/Image1.gif" border="0">
			<%end if
	else%>
	<%if UserTableType = "Dvbbs" then
			if rs(db_User_Face)<>"" then%>
				<img src="<%=BbsPath%><%=rs(db_User_Face)%>" border="0" width="<%=photowidth%>" height="<%=photoheight%>"> 
				<%''显示用户头像，加'bbs/'前缀路径,使图文系统直接显示定向到论坛头像%>
			<%else%>
				<img src="images/nopic.gif" border="0">
			<%end if
		end if%>
	<%end if%>
	</td>
</tr>
<tr>
	<td><p align="right">自我介绍：</td>
	<td><textarea rows="10" name="content" cols="30"  ><%=rs("content")%></textarea></td>
</tr>
<tr>
	<td colspan="2">
		<p align="center">
			<input type="submit" value="修 改" name="cmdok" class="buttonface" >&nbsp; 
			<input type="reset" value="重 置" name="cmdcancel" class="buttonface" >&nbsp; 
			<input type="button" value="返 回" onClick="javascript:history.go(-1)" class="buttonface" >
		</P>
	</td>
</tr>
</form>
</table>
</body>
</html>
<!--#include file="Admin_Bottom.asp"-->
	<%rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	ConnUser.close
	set ConnUser=nothing
end if%>