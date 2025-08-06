<!--#include file="ConnUser.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="char.inc"-->
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
<SCRIPT language="JavaScript" type="text/javascript">
    // Begin morelink
      function morelink(morelink)
      {
        url = 'MoreLink.asp?linktype=1';
        window.open(url,morelink);
      }
    // End morelink-->
 // Begin linkreg
      function linkreg(linkreg)
      {
        url = 'LinkReg.asp';
        window.open(url,linkreg);
      }
    // End linkreg-->
// Begin vote
      function vote(vote)
      {
        url = 'E_Vote.asp?stype=view';
        window.open(url,vote);
      }
    // End vote-->
// Begin adduser
      function adduser(adduser)
      {
        url = 'AddUser.asp';
        window.open(url,adduser);
      }
    // End adduser-->
// Begin getpwd
      function getpwd(getpwd)
      {
        url = 'GetPwd.asp';
        window.open(url,getpwd);
      }
    // End getpwd-->
</script>
<%if TransShow="on" then%>
<META content=revealTrans(Transition=23,Duration=1.0) http-equiv=Page-Enter>
<META content=revealTrans(Transition=23,Duration=1.0) http-equiv=Page-Exit>
<SCRIPT language=JavaScript>
<!--
//页面随机切换效果
	function transDemo(n) {
	if (document.all && navigator.userAgent.indexOf("Mac")==-1) {
	t=document.all.transmeta;
	t.style.width=document.body.clientWidth;
	t.style.height=document.body.offsetHeight;
	t.style.top=document.body.scrollTop;
	t.style.backgroundColor="#003333";
	t.style.visibility="visible";
	t.filters[0].transition=n;
	setTimeout("transShow()"); // separated to force screen paint
	} else {
	alert("You can view transitions only on Windows IE 4.0 and later.");
	} 
	}
	
	function transShow() {
	t.filters[0].Apply();
	t.style.visibility="hidden";
	t.filters[0].Play();
	}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</SCRIPT>
<%end if%>
<%		'获取当前 URL
dim ViewUrl
if Request.ServerVariables("QUERY_STRING")<>"" then
	ViewUrl=Request.ServerVariables("url") &"?"& Request.ServerVariables("QUERY_STRING")&""
else
	ViewUrl=Request.ServerVariables("url") &""
end if
response.cookies(eChsys)("ViewUrl")=ViewUrl%>
<div id=menuDiv style='Z-INDEX: 2; VISIBILITY: hidden; WIDTH: 1px; POSITION: absolute; HEIGHT: 1px; BACKGROUND-COLOR: #9cc5f8'></div>
<%if topbg=1 then %>
<script src="clearevents.js"></script>
<%end if%>
<%if R_BG=1 then %>
<SCRIPT language=JavaScript src="Float_AD.asp"></SCRIPT>
<%end if%>
<table width="1002" height="25" border="0" align="center" cellpadding="0" cellspacing="0" background="images/topbg01.gif">
  <tr>
    <td width="772" align="left"><table width="760" border="0" align="left" cellpadding="0" cellspacing="0">
      <tr>
        <%if Request.cookies(eChsys)("username")="" then%>
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
        <td height="24"><form method="POST" action="ChkLogin.asp" name="UserLogin" onSubmit="return CheckFormUserLogin();" class="login">
            &nbsp;&nbsp;<span align="left" valign="middle" class="login_user">用户：</span>
            <input name="UserName" size="8" class="login_username" onMouseOver ="this.style.backgroundColor='#F1F7FC'" onMouseOut ="this.style.backgroundColor='#FFFFFF'">
            <span align="left" valign="middle" class="login_user">密码：</span>
            <input type="password" name="Passwd" size="8" class="login_username" onMouseOver ="this.style.backgroundColor='#F1F7FC'" onMouseOut ="this.style.backgroundColor='#FFFFFF'">
            <span align="left" valign="middle" class="login_user">验码：</span>
            <input type="text" name="verifycode" size="4" class="login_username" onMouseOver ="this.style.backgroundColor='#F1F7FC'" onMouseOut ="this.style.backgroundColor='#FFFFFF'">
            <span><%=getcode1()%></span>
            <input type="submit" name="Submit" value="登录"  class="login_username" onMouseOver ="this.style.backgroundColor='#F1F7FC'" onMouseOut ="this.style.backgroundColor='#FFFFFF'" title="登录系统">
            <%if reg=1 then%>
          &nbsp;
          <input type="button" name="Submit2" value="注册" class="login_username" onMouseOver ="this.style.backgroundColor='#F1F7FC'" onMouseOut ="this.style.backgroundColor='#FFFFFF'" onClick="javascript:adduser()" title="注册新会员">
          <%end if%>
          &nbsp;
          <input type="button" name="Submit2" value="忘密" class="login_username" onMouseOver ="this.style.backgroundColor='#F1F7FC'" onMouseOut ="this.style.backgroundColor='#FFFFFF'" onClick="javascript:getpwd()" title="忘记密码了？">
        </form></td>
        <%else%>
        <td height="24"><span align="right" valign="middle" style="font-size: 9pt;color: #000000"> 欢迎：<b><%=Request.cookies(eChsys)("UserName")%></b>&nbsp;&nbsp;</span>
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
          <font color="#ff0000">超级管理员</font>
          <%end if%>
          <%if Request.cookies(eChsys)("KEY")="super" and Request.cookies(eChsys)("purview")<>"99999" then%>
          <font color="#ff0000">系统管理员</font>
          <%end if%>
          <%if Request.cookies(eChsys)("KEY")="check" then%>
          <font color="#ff0000">文章审核员</font>
          <%end if%>
          <%if Request.cookies(eChsys)("KEY")="selfreg" then%>
          <font color="#ff0000">注册用户</font>
          <%end if%>
          <%if Request.cookies(eChsys)("KEY")="smallmaster" then%>
          <font color="#ff0000">小类管理员</font>
          <%end if%>
          <%if Request.cookies(eChsys)("KEY")="bigmaster" then%>
          <font color="#ff0000">大类管理员</font>
          <%end if%>
          <%if Request.cookies(eChsys)("KEY")="typemaster" then%>
          <font color="#ff0000">总栏管理员</font>
          <%end if%>
          您的等级：
          <%if Request.cookies(eChsys)("KEY")<>"selfreg" then%>
          <font color="#ff0000">内部成员</font>
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
          <a href="admin_index.asp" class=my>[发文入口]</a>&nbsp;<a href="Exit.asp" class=my>[退出登录]</a>&nbsp;
		  <!--
		  <a href="bloguser.asp?user=<%=Request.cookies(eChsys)("UserName")%>" class=my>[我的文章]</a>
		  -->
		  </span> </td>
        <%end if%>
      </tr>
    </table>
	
	</td>
    <td width="230" align="left"><div align="center"><IMG height=16 src="images/setindex.gif" 
            width=16 align=absMiddle>　<A 
            onclick="this.style.behavior='url(#default#homepage)';this.setHomePage('http://10.124.0.13/');" 
            href="http://10.124.21.10" class="login_user">设为首页</A> 
            　<IMG height=16 src="images/setfav.gif" width=16 
            align=absMiddle> <A 
            onclick="javascript:window.external.AddFavorite('http://10.124.21.10/', '贵州银监局');" 
            href="http://10.124.21.10/#" class="login_user">加入收藏</A></div></td>
  </tr>
</table>
<table width="1002" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#FFFFFF" >
  <tr>
    <td height="60"><a href="<%=TopBannerurl%>">
        <%if gd2="1" then%>
        <img src="<%=TopBanner%>" width="1002" height="123" border="0" align="absmiddle">
        <%else%>
        <object codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' 
      height="123" width="1002" classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'>
          <param name="movie" value='<%=TopBanner%>'>
          <param name="quality" value="high">
          <param name="wmode" value="transparent">
          <embed src="photo/top.swf" quality=high 
      pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" 
      type="application/x-shockwave-flash" width="1002" height="123"> </embed>
      </object>
        <%end if%></a>
	</td>
  </tr>
  <tr>
    <td height="30" valign="bottom">
	  <table width="1002" height="30" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="240" height="30" align="left"  valign="middle" background="images/top_dh_bg2.gif" bgcolor="#ffffff" class="TOP1"><%
						dim mymenu
						mymenu=echuangmenu
						Response.Write mymenu
						%></td>
        <td  valign="middle" align="right" class="TOP1" bgcolor="#ffffff">
			<form method="post" name="myform" action="Result.asp" target="newwindow" class="login">
              <%if showsearch="1" then%>
                <%if search=1 then%>
                <!--#include file="E_Search.asp"-->
               <%else%>
               <!--#include file="E_Search1.asp"-->
               <%end if%>
              <%end if%>
           </form>	
		</td>
      </tr>
    </table>
	</td>
  </tr>
 
  <tr>
    <td height="34"  class="top"  background="Images/a_bg.gif">
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="100%" align="left"><!--#include file="menu.asp"--></td>
      </tr>
    </table>
	</td>
  </tr>
</table>
<table width="1002" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
    <tr >
	   <td >
	     <a href = "http://10.124.0.13/E_Type.asp?E_typeid=12"> <img src="images\kxfzg.gif" width="1002" height="123" border="0" align="absmiddle">
	   </td>
	</tr>
</table>