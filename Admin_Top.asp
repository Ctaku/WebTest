<!--#include file="Conn.ASP"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--'#include file="char.inc"-->
<!--#include file="chkuser.asp" -->

<%dim jingyong
set rs=server.createobject("adodb.recordset")
sql="select * from "& db_User_Table &" where  "& db_User_Name &"='"&Request.cookies(eChsys)("username")&"'"
rs.open sql,ConnUser,1,3
if rs.bof or rs.eof then
	response.redirect "login.asp"
	response.end
end if'
jingyong=rs("jingyong")
rs.close
set rs=nothing

if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="check" or Request.cookies(eChsys)("KEY")="bigmaster" or Request.cookies(eChsys)("KEY")="smallmaster" or Request.cookies(eChsys)("KEY")="typemaster" or (Request.cookies(eChsys)("key")="selfreg" and jingyong=0) then
%>

<HTML><HEAD><TITLE></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<META content="MSHTML 6.00.3790.118" name=GENERATOR>
<LINK href=site.css rel=stylesheet>
<base target="main">
<script language="javascript">
<!--
var displayBar=true;
function switchBar(obj){
	if (displayBar)
	{
		parent.frame.cols="0,*";
		displayBar=false;
		obj.src="images/bar2.gif";
		obj.title="打开左边管理菜单";
	}
	else{
		parent.frame.cols="180,*";
		displayBar=true;
		obj.src="images/bar1.gif";
		obj.title="关闭左边管理菜单";
	}
}
//-->
</script>
<SCRIPT language=javascript>
function checkclick(msg){
	if(confirm(msg)){
		event.returnValue=true;
	}else{
		event.returnValue=false;
	}
}
</SCRIPT>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber1">
<tr>
<td width="15%" background="images/admin_topbg.gif">
  <a href="http://www.echsys.com" target="_blank"><img  src="Images/top_logo.gif" width="147" height="32" border="0" ></a></td>
<td width="2%" background="images/admin_topbg.gif"><img onClick="switchBar(this)" alt="关闭左边管理菜单" style="cursor:hand" src="Images/bar1.gif" width="15" height="15"></td>
<td width="55%" background="images/admin_topbg.gif"><table cellpadding=0 cellspacing=0 align=center width=100%>
  <tr>
    <td height=20 class="title">欢迎您： <%=Request.cookies(eChsys)("UserName")%>
        <%if db_Birthday_Select="EChuang" then		'性别字段是原字段%>
        <%=Request.cookies(eChsys)("sex")%>
        <%else
				if db_Birthday_Select="Text" then		'性别字段是BBS的文本型阿拉伯数字
					if Request.cookies(eChsys)("sex")=1 then%>
      男
      <%else
						if Request.cookies(eChsys)("sex")=0 then%>
      女
      <%else%>
      保密
      <%end if
					end if
				end if
			end if%>
      您的权限：
      <%if Request.cookies(eChsys)("KEY")="super" and Request.cookies(eChsys)("purview")="99999" then%>
      <font color="#990000">超级管理员</font>
      <%end if%>
      <%if Request.cookies(eChsys)("KEY")="super" and Request.cookies(eChsys)("purview")<>"99999" then%>
      <font color="#990000">系统管理员</font>
      <%end if%>
      <%if Request.cookies(eChsys)("KEY")="check" then%>
      <font color="#990000">文章审核员</font>
      <%end if%>
      <%if Request.cookies(eChsys)("KEY")="selfreg" then%>
      <font color="#990000">注册用户</font>
      <%end if%>
      <%if Request.cookies(eChsys)("KEY")="smallmaster" then%>
      <font color="#990000">小类管理员</font>
      <%end if%>
      <%if Request.cookies(eChsys)("KEY")="bigmaster" then%>
      <font color="#990000">大类管理员</font>
      <%end if%>
      <%if Request.cookies(eChsys)("KEY")="typemaster" then%>
      <font color="#990000">总栏管理员</font>
      <%end if%>
      您的等级：
      <%if Request.cookies(eChsys)("KEY")<>"selfreg" then%>
      <font color="#990000">内部成员</font>
      <%end if%>
      <%if Request.cookies(eChsys)("KEY")="selfreg" and Request.cookies(eChsys)("reglevel")="1" then%>
      <font color="red">普通</font>
      <%end if%>
      <%if Request.cookies(eChsys)("KEY")="selfreg" and Request.cookies(eChsys)("reglevel")="2" then%>
      <font color="red">高级</font>
      <%end if%>
      <%if Request.cookies(eChsys)("KEY")="selfreg" and Request.cookies(eChsys)("reglevel")="3" then%>
      <font color="red">特级</font>
      <%end if%>
      登录次数：<%=Request.cookies(eChsys)("UserLoginTimes")%><br>
    </td>
  </tr>
</table></td>
<td width="21%" background="images/admin_topbg.gif">
</td>
</tr>
</table>
</BODY>
</HTML>
<%else%>
	<script language=javascript>
		history.back()
		alert("对不起，管理员尚未开通您的帐号，请稍候！")
	</script>
<%end if%>