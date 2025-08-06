<%@ LANGUAGE = VBScript %>
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
<style type="text/css">
BODY{ 
      background-color: #FA7727;MARGIN: 0px; FONT: 9pt 宋体;}
</style>
<SCRIPT src="inc/js.js" type=text/javascript></SCRIPT>
<SCRIPT src="inc/exit.js" type=text/javascript></SCRIPT>
<SCRIPT language=JavaScript>
<!--
/*for ie and ns*/
document.onclick=function(evt){
var evt=evt?evt:(window.event)?window.event:""
var e=evt.target?evt.target:evt.srcElement
evt.cancelBubble=true
if(e.tagName=="A"&&evt.shiftKey)return false
}
//-->
</SCRIPT>

<SCRIPT language=javascript>
  var whichOpen="";
  var whichContinue='';
</SCRIPT>

</HEAD>
<BODY oncontextmenu=window.event.returnValue=false onkeydown="if(event.keyCode==78&amp;&amp;event.ctrlKey)return false;" onresize=maxWin() leftMargin=0 topMargin=0 onload=maxWin() marginwidth="0" marginheight="0">
<SCRIPT language=JavaScript1.2 src="inc/menu.js"></SCRIPT>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=42 valign=bottom  align=middle ><img src="images/admin_title.gif" width=158 height=38 border="0" style="vertical-align:bottom; border:0px;"></td>
  </tr>
</table>
<%if (request.cookies(eChsys)("KEY")="super" or request.cookies(eChsys)("KEY")="bigmaster" or request.cookies(eChsys)("KEY")="check" or request.cookies(eChsys)("key")="typemaster") then%>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=25 background=images/title_bg_quit.gif class=menu_title1 onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title1';><div align="center"><a href="admin_main.asp" target=main>管理首页</a> | <A onClick="checkclick('您真的退出网站？')" href="Admin_Exit.asp" target=main>退 出 </a></div></td>
  </tr>
  <tr>
    <td style="display:">&nbsp;</td>
  </tr>
</table>
<!--
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=25 valign="middle" background=images/admin_left_1.gif class=menu_title1 id=menuTitle0 onClick="showsubmenu(0)" onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title1';><div align="center">系统配置</div></td>
  </tr>
  <tr>
    <td id='submenu0'><div class="sec_menu" style="width:158">
      <table cellpadding=0 cellspacing=0 align=center width=135 style="POSITION: relative; TOP: 5px">
<%if Request.cookies(eChsys)("ManageUserName")<>"" then%>
        <tr>
          <td height=20><div align="center"><img src="Images/bullet.gif" width="15" height="20"><A href="E_System.asp" target=main> 网站属性</a></div></td>
        </tr>
        <tr>
          <td height=20><div align="center"><img src="Images/bullet.gif" width="15" height="20"><A href="E_System1.asp" target=main> 功能设置</A></div></td>
        </tr>
		<tr>
          <td height=20><div align="center"><img src="Images/bullet.gif" width="15" height="20"><A href="CssEdit.asp" target=main> 模板编辑</a></div></td>
        </tr>
        <tr>
          <td height=20><div align="center"><img src="Images/bullet.gif" width="15" height="20"><A href="new.asp" target=main> 系统初始</a></div></td>
        </tr>
<%else%>
		<tr>
          <td height=20><div align="center"><img src="Images/bullet.gif" width="15" height="20"><A href="Manage.asp" target=main> 管理登录</a></div></td>
        </tr>
	<%end if%>
      </table>
    </div>
	-->
        <div  style="width:158">
          <table cellpadding=0 cellspacing=0 align=center width=135>
            <tr>
              <td height=5></td>
            </tr>
          </table>
        </div></td>
  </tr>
</table>
<%end if%>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=25 valign="middle" background=images/admin_left_2.gif class=menu_title1 id=menuTitle1 onClick="showsubmenu(1)" onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title1';><div align="center">文章管理</div></td>
  </tr>
  <tr>
    <td  id='submenu1'><div class="sec_menu" style="width:158">
      <table cellpadding=0 cellspacing=0 align=center width=135 style="POSITION: relative; TOP: 5px">
        <tr>
          <td height=20>
		  <div align="center"><img src="images/bullet.gif"> <A title=展开/收缩菜单 onclick=showsubmenu(11) href="E_TypeManage.asp" target=main>栏目管理</a></div></td>
        </tr>
					<TR>
					<TD id=submenu11 style="DISPLAY: none" vAlign=center height=20>
<%
response.write "	<!--==========================-->"& chr(10)
response.write "	<!--      栏目菜单开始        -->"& chr(10)
response.write "	<!--==========================-->"
sqltype="select * from "& db_EC_Type_Table & " order by E_typeorder"
Set rstype= Server.CreateObject("ADODB.Recordset")
rstype.open sqltype,conn,1,1

dim Type_Content,BigClass_Content
Type_Content=""
BigClass_Content=""

typeNum=rstype.recordcount

if rstype.bof and rstype.eof then 
	response.write "&nbsp;&nbsp;&nbsp;&nbsp;栏目正在建设中！"
else
	menu_pic1="images/menu1.gif"
	menu_pic2="images/menu2.gif"	
	menu_spacer_pic="images/spacer.gif"
	
	do while not rstype.eof
		SqlBigClass="select * from "& db_EC_BigClass_Table & "  where E_typeid="& rstype("E_typeid") &" Order by E_bigclassorder"	'选择大组类
		Set rsBigClass= Server.CreateObject("ADODB.Recordset")
		rsBigClass.open SqlBigClass,conn,1,1
		if rsBigClass.bof and rsBigClass.eof then
			'总栏下无大组栏
			Type_Content=Type_Content & chr(10)
			Type_Content=Type_Content &"<DIV class=parent id=KB"& rstype("E_typeid") &"Parent>"& chr(10)
			Type_Content=Type_Content &"&nbsp;&nbsp;<IMG src="& menu_pic1 &">"& chr(10)
			Type_Content=Type_Content &"	<A href='E_BigClassManage.asp?E_typeid="& rstype("E_typeid") &"' target=main>"& rstype("E_typename") &"</A>"& chr(10)
			Type_Content=Type_Content &"</DIV>"& chr(10)
			response.write Type_Content
			Type_Content=""
		else
			'总栏下有大组栏
			Type_Content=Type_Content & chr(10)
			Type_Content=Type_Content &"<DIV class=parent id=KB"& rstype("E_typeid") &"Parent>"& chr(10)
			Type_Content=Type_Content &"&nbsp;&nbsp;<IMG title='展开/收缩菜单' style='CURSOR: hand' onclick='expandIt("& chr(34) &"KB"& rstype("E_typeid") &""& chr(34) &")' src="& menu_pic2 &">"& chr(10)
			Type_Content=Type_Content &"	<A href='E_BigClassManage.asp?E_typeid="& rstype("E_typeid") &"' target=main>"& rstype("E_typename") &"</A>"& chr(10)
			Type_Content=Type_Content &"</DIV>"& chr(10)
			Type_Content=Type_Content &"<DIV class=child id=KB"& rstype("E_typeid") &"Child >"& chr(10)
			response.write Type_Content
			Type_Content=""

			do while not rsBigClass.eof
				SqlSmallClass="select * from "& db_EC_SmallClass_Table & "  where E_BigClassID="& rsBigClass("E_BigClassID") &" Order by smallclassorder"	'选择小组类
				Set rsSmallClass= Server.CreateObject("ADODB.Recordset")
				rsSmallClass.open SqlSmallClass,conn,1,1
				if rsSmallClass.bof and rsSmallClass.eof then
					'大组下无小组栏
					BigClass_Content=BigClass_Content &"	&nbsp;&nbsp;"& chr(10)
					BigClass_Content=BigClass_Content &"	<IMG src="& menu_pic1 &">"& chr(10)
					BigClass_Content=BigClass_Content &"	<A href='E_SmallClassManage.asp?E_BigClassID="& rsBigClass("E_BigClassID") &"' target=main>"& rsBigClass("E_BigClassName") &"</A>"& chr(10)
					BigClass_Content=BigClass_Content &"	<BR>"& chr(10)
					response.write BigClass_Content
					BigClass_Content=""
				else
					'大组下有小组栏
					BigClass_Content=BigClass_Content &"	&nbsp;&nbsp;"& chr(10)
					BigClass_Content=BigClass_Content &"	<IMG title=展开/收缩菜单 style='CURSOR: hand' onclick=showsubmenu("& rstype("E_typeid") & rsBigClass("E_BigClassID") &") src="& menu_pic2 &">"& chr(10)
					BigClass_Content=BigClass_Content &"	<A href='E_SmallClassManage.asp?E_BigClassID="& rsBigClass("E_BigClassID") &"' target=main>"& rsBigClass("E_BigClassName") &"</A>"& chr(10)
					BigClass_Content=BigClass_Content &"	<BR>"& chr(10)
					BigClass_Content=BigClass_Content &"	<DIV id=submenu"& rstype("E_typeid") & rsBigClass("E_BigClassID") &" style='DISPLAY: none'>"& chr(10)
					response.write BigClass_Content
					BigClass_Content=""
					
					rsSmallClass.movefirst
					do while not rsSmallClass.eof
						response.write "		&nbsp;&nbsp;"
						response.write "&nbsp;&nbsp;"& chr(10)
						response.write "		<IMG src="& menu_pic1 &">"& chr(10)
						response.write "		<A href='ListNews.asp?E_SmallClassID="& rsSmallClass("E_SmallClassID") &"' target=main>"& rsSmallClass("E_smallclassname") &"</A>"& chr(10)
						response.write "		<BR>"& chr(10)
						rsSmallClass.movenext
					loop
					response.write "	</DIV>"& chr(10)
				end if
				rsSmallClass.close
				set rsSmallClass=nothing
				
				rsBigClass.movenext
			loop
			response.write "</DIV>"& chr(10)
		end if
		rsBigClass.close
		set rsBigClass=nothing

		rstype.movenext
	loop
end if

rstype.close
set rstype=nothing

conn.close
set conn=nothing

response.write "	<!--==========================-->"& chr(10)
response.write "	<!--      栏目菜单结束        -->"& chr(10)
response.write "	<!--==========================-->"
%>
					</TD>
					</TR>		
        <tr>
          <td height=20><div align="center"><img src="images/bullet.gif"> <A href="newsaddd1.asp" target="main">添加文章</a></div></td>
        </tr>
        <tr>
          <td height=20><div align="center"><img src="images/bullet.gif"> <a href="MyNews.asp" target="main">我的文章</a></div></td>
        </tr>
      </table>
    </div>
        <div  style="width:158">
          <table cellpadding=0 cellspacing=0 align=center width=135>
            <tr>
              <td height=5></td>
            </tr>
          </table>
        </div></td>
  </tr>
</table>
<!--
<%if (request.cookies(eChsys)("KEY")="super" or request.cookies(eChsys)("KEY")="bigmaster" or request.cookies(eChsys)("KEY")="check" or request.cookies(eChsys)("key")="typemaster") then%>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=25 valign="middle" background=images/admin_left_3.gif class=menu_title1 id=menuTitle2 onClick="showsubmenu(2)" onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title1';><div align="center">言论管理</div></td>
  </tr>
  <tr>
    <td id='submenu2'><div class="sec_menu" style="width:158">
      <table cellpadding=0 cellspacing=0 align=center width=135 style="POSITION: relative; TOP: 5px">
<%if Request.cookies(eChsys)("ManageUserName")<>"" then%>
        <tr>
          <td height=20><img src="images/bullet.gif"> <a href="GuestReview.asp" target="main">留言管理</a>| <a href="ReViewManage.asp" target="main">评论管理</a></td>
        </tr>
        <tr>
          <td height=20><img src="images/bullet.gif"> <A href="E_LeadMailList.asp" target="main">领导信箱</a>| <A href="ComplainList.asp" target="main">投诉举报</a></td>
        </tr>
        <tr>
          <td height=20><img src="images/bullet.gif"> <a href="E_OpinionList.asp" target="main">建言献策</a>|</td>
        </tr>
<%else%>
		<tr>
          <td height=20><div align="center"><img src="Images/bullet.gif" width="15" height="20"><A href="Manage.asp" target=main> 管理登录</a></div></td>
        </tr>
	<%end if%>
      </table>
    </div>
	-->
        <div  style="width:158">
          <table cellpadding=0 cellspacing=0 align=center width=135>
            <tr>
              <td height=5></td>
            </tr>
          </table>
        </div></td>
  </tr>
</table>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=25 valign="middle" background=images/admin_left_4.gif class=menu_title1 id=menuTitle3 onClick="showsubmenu(3)" onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title1';><div align="center">附加管理</div></td>
  </tr>
  <tr>
    <td  id='submenu3'><div class="sec_menu" style="width:158">
      <table cellpadding=0 cellspacing=0 align=center width=135 style="POSITION: relative; TOP: 5px">
<%if Request.cookies(eChsys)("ManageUserName")<>"" then%>
        <tr>
          <td height=20><img src="images/bullet.gif"> <a href="CheckNews1.asp" target="main">文章审核</a>| <a href="allnewsmanage.asp" target="main">文章检索</a></td>
        </tr>
        <tr>
          <td height=20><img src="images/bullet.gif"> <a href="SpecialManage.asp" target="main">专题管理</a>| <a href="E_VoteManage.asp" target="main">投票管理</a></td>
        </tr>
        <tr>
          <td height=20><img src="images/bullet.gif"> <A href="E_BoardManage.asp" target="main">公告管理</a>| <A href="LinkManage.asp" target="main">友情链接</a></td>
        </tr>
        <tr>
           <a href="Admin_UpFileManage.asp" target="main">附件管理</a></td>
        </tr>
		  <tr>
          <td height=20><img src="images/bullet.gif"> <%if DbType = "MSSQL" then%><A href="SQLBackRest.asp" target=main>备份恢复</A><%elseif DbType = "ACCESS" then%><A href="Db_compact.asp" target=main>备份压缩</A><%end if%>| <a href="AspCheck.asp" target="main">信息探测</a></td>
        </tr>
<%else%>
		<tr>
          <td height=20><div align="center"><img src="Images/bullet.gif" width="15" height="20"><A href="Manage.asp" target=main> 管理登录</a></div></td>
        </tr>
	<%end if%>
      </table>
    </div>
        <div  style="width:158">
          <table cellpadding=0 cellspacing=0 align=center width=135>
            <tr>
              <td height=5></td>
            </tr>
          </table>
        </div></td>
  </tr>
</table>

<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=25 valign="middle" background=images/admin_left_5.gif class=menu_title1 id=menuTitle4 onClick="showsubmenu(4)" onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title1';><div align="center">用户管理</div></td>
  </tr>
  <tr>
    <td id='submenu4'><div class="sec_menu" style="width:158">
      <table cellpadding=0 cellspacing=0 align=center width=135 style="POSITION: relative; TOP: 5px">
<%if Request.cookies(eChsys)("ManageUserName")<>"" then%>
        <tr>
          <td height=20><img src="images/bullet.gif"> <a href="Edit.asp" target="main">修改资料</a>| <a href="E_DepManage.asp" target="main">部门管理</a></td>
		</tr>
        <tr>
          <td height=20><img src="images/bullet.gif"> 
		  <A href="E_UserManage.asp" target="main">普通用户</a>| <A href="useradd1.asp" target="main">添加用户</a></td>
        </tr>
        <tr>
          <td height=20><img src="images/bullet.gif"> <a href="AdminManage.asp" target="main">超管管理</a>| <a href="AdminAdd1.asp" target="main">添加超管</a></td>
        </tr>
<%else%>
		<tr>
          <td height=20><div align="center"><img src="Images/bullet.gif" width="15" height="20"><A href="Manage.asp" target=main> 管理登录</a></div></td>
        </tr>
	<%end if%>
      </table>
    </div>
        <div  style="width:158">
          <table cellpadding=0 cellspacing=0 align=center width=135>
            <tr>
              <td height=5></td>
            </tr>
          </table>
        </div></td>
  </tr>
</table>
<%end if%>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=25 valign="middle" background=images/admin_left_5.gif class=menu_title1 id=menuTitle5 onClick="showsubmenu(5)" onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title1';><div align="center">个人资料</div></td>
  </tr>
  <tr>
    <td id='submenu5'><div class="sec_menu" style="width:158">
      <table cellpadding=0 cellspacing=0 align=center width=135 style="POSITION: relative; TOP: 5px">
        <tr>
          <td height=20><div align="center"><img src="images/bullet.gif"> <a href="Edit.asp" target="main">修改资料</a></div></td>
        </tr>
      </table>
    </div>
        <div  style="width:158">
          <table cellpadding=0 cellspacing=0 align=center width=135>
            <tr>
              <td height=5></td>
            </tr>
          </table>
        </div></td>
  </tr>
</table>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=25 valign="middle" background=images/admin_left_8.gif class=menu_title1 id=menuTitle6 onClick="showsubmenu(6)" onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title1';><div align="center">登陆管理</div></td>
  </tr>
  <tr>
    <td  id='submenu6'><div class="sec_menu" style="width:158">
      <table cellpadding=0 cellspacing=0 align=center width=135 style="POSITION: relative; TOP: 5px">
        
        
        <tr>
          <td height=20><div align="center"><img src="images/bullet.gif"> <A onClick="checkclick('您是否要重新登录？')" href="Login.asp" target="main">重新登录</a></div></td>
        </tr>
        <tr>
          <td height=20><div align="center"><img src="images/bullet.gif"> <A onClick="checkclick('您真的退出网站？')" href="Admin_Exit.asp" target="main">退出管理</a></div></td>
        </tr>
      </table>
    </div>
        <div  style="width:158">
          <table cellpadding=0 cellspacing=0 align=center width=135>
            <tr>
              <td height=5></td>
            </tr>
          </table>
        </div></td>
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