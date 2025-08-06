<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="function.asp"-->
<!--'#include file="char.inc"-->
<%
IF request.cookies(eChsys)("KEY")="" THEN

else
	usernamecookie=CheckStr(request.cookies(eChsys)("UserName"))
	passwdcookie=CheckStr(trim(Request.cookies(eChsys)("passwd")))
	KEYcookie=CheckStr(trim(request.cookies(eChsys)("KEY")))
	if usernamecookie="" or passwdcookie="" then
		response.cookies(eChsys)("UserName")=""
		response.cookies(eChsys)("KEY")=""
		response.cookies(eChsys)("purview")=""
		response.cookies(eChsys)("fullname")=""
		response.cookies(eChsys)("reglevel")=""
	else
		'判断用户的合法性
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_User_Table &" where "& db_User_Name &"='"&usernamecookie&"'"
		rs.open sql,ConnUser,1,1
		if rs.eof and rs.bof then
			response.cookies(eChsys)("UserName")=""
			response.cookies(eChsys)("KEY")=""
			response.cookies(eChsys)("purview")=""
			response.cookies(eChsys)("fullname")=""
			response.cookies(eChsys)("reglevel")=""
		end if
		IF passwdcookie<>rs(db_User_Password) THEN
			response.cookies(eChsys)("UserName")=""
			response.cookies(eChsys)("KEY")=""
			response.cookies(eChsys)("purview")=""
			response.cookies(eChsys)("fullname")=""
			response.cookies(eChsys)("reglevel")=""
		END IF
		'下面判断用户级别实际在有用户级别是都应该判断
		if KEYcookie<>rs("OSKEY") then
			response.cookies(eChsys)("UserName")=""
			response.cookies(eChsys)("KEY")=""
			response.cookies(eChsys)("purview")=""
			response.cookies(eChsys)("fullname")=""
			response.cookies(eChsys)("reglevel")=""
		end if
		rs.close
		set rs=nothing
	END IF
END IF



'该文件需要进行调整和设置
dim E_typename
NewsID=Request.QueryString("NewsID")
Page=Request.QueryString("page")

if page="" then
	page=1
	elseif not IsNumeric(page) then
		Show_Err("非法参数！<br><br><a href='javascript:history.back()'>返回</a>")
		response.end
	end if
	page=int(page)
	if newsid="" then
		Show_Err("未指定参数！<br><br><a href='javascript:history.back()'>返回</a>")
		response.end
	elseif not IsNumeric(newsid) then
		Show_Err("非法参数！<br><br><a href='javascript:history.back()'>返回</a>")
		response.end
	else
		'判断该篇文章是否审核
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_EC_News_Table &" where NewsId="&NewsId
		rs.open sql,conn,3,3
		if rs.eof and rs.bof then
			rs.close
			set rs=nothing
			Show_Err("指定的文章不存在！<br><br><a href='javascript:history.back()'>返回</a>")
			response.end
		else
			checked=rs("checkked")
			if checked=1 or Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" then	'文章已审核并且是相关管理员
				Click=rs("Click")
				if isnull(rs("Click"))  then
					conn.execute("update "& db_EC_News_Table &" Set Click=1 where NewsID=" & NewsID )
				else 
					conn.execute("update "& db_EC_News_Table &" Set Click=click+1 where NewsID=" & NewsID )
				end if
			end if
		
			rs.close
			set rs=nothing
		end if

		set rs=server.CreateObject("ADODB.RecordSet")
		if uselevel=1 then
			if Request.cookies(eChsys)("key")="" then
				rs.Source="select * from "& db_EC_News_Table &" where checkked=1 and newslevel=0 and newsid="&newsid
			end if
			if Request.cookies(eChsys)("key")="selfreg" then
				if Request.cookies(eChsys)("reglevel")=3 then
					rs.Source="select * from "& db_EC_News_Table &" where checkked=1 and newslevel<=3 and newsid="&newsid
				end if
				if Request.cookies(eChsys)("reglevel")=2 then
					rs.Source="select * from "& db_EC_News_Table &" where checkked=1 and newslevel<=2 and newsid="&newsid
				end if
				if Request.cookies(eChsys)("reglevel")=1 then
					rs.Source="select * from "& db_EC_News_Table &" where checkked=1 and newslevel<=1 and newsid="&newsid
				end if
			end if
		end if
			if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then
				rs.Source="select * from "& db_EC_News_Table &" where newsid="&newsid
			else
				rs.Source="select * from "& db_EC_News_Table &" where newsid="&newsid
			end if
			rs.Open rs.Source,conn,1,1
			E_BigClassID=rs("E_BigClassID")
			E_SmallClassID=rs("E_SmallClassID")
			title=htmlencode4(trim(rs("title")))
			title1=htmlencode4(trim(rs("title")))
			about=htmlencode4(trim(rs("about")))
			Author=htmlencode4(trim(rs("Author")))
			editor=htmlencode4(trim(rs("editor")))
			Original=htmlencode4(trim(rs("Original")))
			image=rs("image")
			UpdateTime=trim(rs("UpdateTime"))
			News_Content=rs("Content")
			SpecialID=rs("SpecialID")
			SpecialID2=rs("SpecialID2")
			click=rs("click")
			EnCode=trim(rs("EnCode"))
			E_typeid=rs("E_typeid")
			titletype=rs("titletype")
			titlecolor=rs("titlecolor")
			titleface=rs("titleface")
			editor=rs("editor")
			wzdj=rs("newslevel")
			backtype=rs("backtype")
			backtype1=rs("backtype1")
			rs.Close
			set rs=nothing

			set rs=server.CreateObject("ADODB.RecordSet")
			rs.Source="select * from "& db_EC_Type_Table &" where E_typeid=" & E_typeid
			rs.Open rs.Source,conn,1,1
			E_typename=rs("E_typename")
			rs.Close
			set rs=nothing

			set rs=server.CreateObject("ADODB.RecordSet")
			If(isnull(E_BigClassID)) Then
				E_BigClassName=" "
				E_BigClassID = 0
			Else
				rsSource="select * from "& db_EC_BigClass_Table &" Where E_BigClassID="& E_BigClassID
				rs.Open rsSource,conn,1,3
				If (rs.eof or rs.bof) Then
					E_BigClassName=" "
				else
					E_BigClassName=rs("E_BigClassName")
				End if
				rs.close
				set rs=nothing
			End if
			
			if E_SmallClassID<>"" then
				set rs=server.CreateObject("ADODB.RecordSet")
				rs.Source="select * from "& db_EC_SmallClass_Table &" Where E_SmallClassID=" & E_SmallClassID
				rs.Open rs.Source,conn,1,1
				E_smallclassname=rs("E_smallclassname")
				rs.close
				set rs=nothing
			end if%>
<%
'文章ip限制
	if cINT(backtype1)=1 then
	dim byte2
	byte2=split(newsipType,"|")
	for i=0 to ubound(byte2)
		if Request.serverVariables("REMOTE_ADDR")<>byte2(i) then
		Show_Err("对不起，此篇文章只供内部浏览，请联系管理员！。<br><br><a href='javascript:history.back()'>返回</a>")
		Response.End
		end if
	next
	end if
%>

<html><head><meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=title%><%if E_SmallClassID<>"" then%>_<%=E_smallclassname%><%end if%>_<%=E_BigClassName%>_<%=E_typename%>_<%=eChSys%></title>
<LINK href="news.css" rel="stylesheet">
<script language="JavaScript" type="text/JavaScript">

function validateMode()
{
  if (!bTextMode) return true;
  alert("请先取消查看HTML源代码，进入“编辑”状态，然后再使用系统编辑功能!");
  message.focus();
  return false;
}
function validateModea()
{
  if (!bTextMode) return true;
  alert("请先取消查看HTML源代码!");
  message.focus();
  return false;
}

function sign()
{if (!validateMode()) return;
message.focus();

var range =message.document.selection.createRange();
str1="<font color='red'><b>|签收|</b>文件已经阅读</font>"
   range.pasteHTML(str1);
}

function img_zoom(e, o)		//图片鼠标滚轮缩放
{
	var zoom = parseInt(o.style.zoom, 10) || 100;
	zoom += event.wheelDelta / 12;
	if (zoom > 0) o.style.zoom = zoom + '%';
		return false;
}
</script>

<script language=javascript>
function user_login()
{
	document.UserLogin.UserName.focus();
	return false;
}
</script>

<script language="JavaScript" type="text/JavaScript">
<!--
function MM_preloadImages() { //v3.0
	var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
	var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
	if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}
//-->
</script>

<script language="JavaScript" type="text/JavaScript">

var size=14;			//字体大小

function fontZoom(fontsize){	//设置字体
	size=fontsize;
	document.getElementById('fontZoom').style.fontSize=size+'px';
}

function fontMax(){	//字体放大
	size=size+2;
	document.getElementById('fontZoom').style.fontSize=size+'px';
}

function fontMin(){	//字体缩小
	size=size-2;
	if (size < 2 ){
		size = 2
	}
	document.getElementById('fontZoom').style.fontSize=size+'px';
}

</SCRIPT>

<script language="javascript">
<!--
function changedata() {
	if ( document.AddReview.none.checked ) {
		document.AddReview.Author.value = "游客";
	} 
}
function changedata1() {
	if ( document.AddReview.none1.checked ) {
		document.AddReview.email.value = "guest@echsys.com";
	} 
}

//-->
</script>

<script language=javascript>
function CheckFormAddReview()	//检测评论发表栏填写项目是否为空
{
	if(document.AddReview.Author.value=="")
	{
		alert("请输入用户名！");
		document.AddReview.Author.focus();
		return false;
	}
	if(document.AddReview.email.value == "")
	{
		alert("请输入您的EMAIL！");
		document.AddReview.email.focus();
		return false;
	}
	if(document.AddReview.content.value == "")
	{
		alert("请输入评论内容！");
		return false;
	}
}
</script>
<style type="text/css">
<!--
.STYLE1 {color: #FF3333}
-->
</style>
</head>
<%if ScrollClick = "double" then%>
	<SCRIPT language=JavaScript>
	//双击滚屏代码
	var currentpos,timer;
	
	function initialize()
	{
	timer=setInterval("scrollwindow()",50);
	}
	
	function sc(){
	clearInterval(timer);
	}
	
	function scrollwindow()
	{
	currentpos=document.body.scrollTop;
	window.scroll(0,++currentpos);
	if (currentpos != document.body.scrollTop)
		sc();
	}
	document.onmousedown=sc
	document.ondblclick=initialize
	</script>
<%else%>
	<SCRIPT>
	//单击拖动屏幕方式
	var old_y=0;  //记录鼠标按下时的Ｙ轴位置
	var yn=0;  //用于记录鼠标状态
	function moveit()
	{
	if(yn==1 &&  event.button==1)  //如果鼠标左键按下就移动页面
	document.body.scrollTop=(old_y-event.clientY); //移动页面
	}
	function downit()
	{old_y=event.clientY+document.body.scrollTop; //记住鼠标按下时的Ｙ轴位置
	yn=1; //用于记住鼠标已按下
	}
	
	function upit()
	{yn=0;}  //记住鼠标已放开
	
	document.onmouseup=upit; //鼠标放开时执行upit()函数
	document.onmousedown=downit; //鼠标按下时执行downit()函数
	document.onmousemove =moveit; //鼠标移动时执行moveit()函数
	</SCRIPT>
<%end if%>
<body topmargin="0" marginheight="0">
<SCRIPT language=javascript>
<!--
  function do_color(vobject,vvar)
  { document.getElementById(vobject).style.color=vvar; }
  function do_zooms(vobject,vvar)
  { document.getElementById(vobject).style.fontsize=vvar+'px'; }
-->
</SCRIPT>
<br><br>
<table width="80%" border="0" align="center" cellpadding="0" cellspacing="0"  bgcolor="#F1F7FC">
<tr><td width="80%" height="124" valign="top"  class="daohang" ><table width="780" height="94" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="212" height="31"><div align="center"></div></td>
    <!--
	<td width="303"><div align="center"><a href=send.asp?NewsID=<%=NewsID%> target=_blank  class="bottom">将本信息发给好友</a></div></td>
	-->
    <td width="97"><div align="center"><a href="javascript:window.print()" class="bottom">打印本页</a></div></td>
    <td width="78"><div align="center"><a href="javascript:window.close()" class="bottom">关闭</a></div></td>
    <td width="90"><div align="center"></div></td>
  </tr>
  <tr>
    <td colspan="5"><div align="left">&nbsp;&nbsp;发表日期：<%=year(updateTime)%>年<%=month(updateTime)%>月<%=day(updateTime)%>日&nbsp; 共浏览<font color=red><%=click%></font> 次 &nbsp;
        <%if Original<>"" then%>
出处：<%=Original%>
<%end if%>

  </tr>
</table></td>
</tr>
<tr>
<td >
 <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="100%" align=center> <br>
                      <font color="<%=titlecolor%>" size="+3" face="楷体_GB2312"><strong><%=title1%></strong></font>
                  <HR align="center" width="95%" noshade style="color:#EFF5FA"></td>
                </tr>
				<tr>
               <td height="20" ><div align="center">发布人：<%=editor%>&nbsp;&nbsp;
			    </td>
              </tr>
              <tr>
             <td>
                <%	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;" & Content%><br>
              </td>
             </tr>
                <%
set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from "& db_EC_News_Table &" where NewsID=" & NewsID
rs.Open rs.source,conn,1,1
E_typeid=rs("E_typeid")
Title=trim(rs("Title"))
image=rs("image")

dim mode
    
set rs11=server.CreateObject("ADODB.RecordSet")
rs11.Source="select * from "& db_EC_Type_Table &" where E_typeid="&E_typeid&" order by E_typeid"
rs11.Open rs11.Source,conn,1,1
mode=rs11("mode")
rs11.close
set rs11=nothing

''添加图片鼠标滚轮缩放效果
if mouse_wheel_zoom="on" then
	News_Content=replace(News_Content,"<IMG","<IMG onmousewheel='return img_zoom(event,this)' onload='javascript:if(this.width>screen.width-333)this.width=screen.width-333'",1,-1,1) 
end if
''图片上传路径还原为 config.asp 中设定的 [FileUploadPath] 值
News_Content=replace(News_Content,"="&chr(34)&"uploadfile/","="&chr(34)&FileUploadPath,1,-1,1)

arr_Content=split(News_Content,"[---分页---]")
MaxPages=ubound(arr_Content)
%>
                <tr>
                  <td width="100%"  align="center">
                    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" style="TABLE-LAYOUT: fixed">
                      <tr>
                        <td width="100%" align=center></td>
                      </tr>
                      <tr>
                        <TD class=newstitle id=fontzoom vAlign=top><br>
                            <%if (checked<>1) and ((Request.cookies(eChsys)("key")<>"super") and (Request.cookies(eChsys)("key")<>"typemaster") and (Request.cookies(eChsys)("key")<>"bigmaster") and (Request.cookies(eChsys)("key")<>"smallmaster")) then	'文章未审核,并且是非相关管理员
	response.write "<P><CENTER><strong><font color='ff0000' size='+2' face='隶书'>文章还未经过审核<br>请等待或者通知管理员审核才能阅览！</font></strong></CENTER>"
	response.write "<P><CENTER><IMG SRC='" &ReadNews_CopyRight_Logo& "' BORDER='0' ALT=''></CENTER>"
else	'文章已审核
	if checked<>1 then
		response.write "<P><CENTER><strong><font color='ff0000' size='+2' face='隶书'>提醒：该文章还未通过审核</font></strong></CENTER>"
	end if



if uselevel=1 then

	if cINT(wzdj)<1 then
		Response.Write arr_Content(Page-1)%>
                            <% else %>
                            <%if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then %>
                            <%Response.Write arr_Content(Page-1)%>
                            <% else %>
                            <%if  Request.cookies(eChsys)("key")="" then%>
                            <br>
                            <font color="#cc0000"><b>内容简介：</b></font><br>
                            <br>
                            <%=CutStr(nohtml(rs("Content")),10)%>... <br>
                            <br>
                            <br>
                            <font color="#cc0000"><b>友情提醒：</b></font><br>
                            <br>
                  你还没注册？或者没有登录？或者权限不够？这篇文章要求是本站符合权限要求的注册用户才能阅读！<br>
                  <br>
                  <%
				response.write "文章级别："
				response.write cINT(wzdj)
				response.write "级"
				%>
                  <br>
                  <%
				response.write "您的权限："
				response.write "未注册"
				%>
                  <br>
                  <br>
                  如果你还没注册，请赶紧点此<a href="#" OnClick="adduser()"  class=bottom><font color="#cc0000">注册</font></a>吧！<br>
                  <br>
                  如果你已经注册但还没登录，请赶紧点此<a href="#" OnClick="user_login()"  class=bottom><font color="#cc0000">登录</font></a>吧！<br>
                  <% else %>
                  <%if  Request.cookies(eChsys)("key")="selfreg" then
					if cINT(Request.cookies(eChsys)("reglevel"))>=cINT(wzdj) then%>
                  <%Response.Write arr_Content(Page-1)%>
                  <% else %>
                  <br>
                  <font color="#cc0000"><b>内容简介：</b></font> <br>
                  <br>
                  <%=CutStr(nohtml(rs("Content")),10)%>... <br>
                  <br>
                  <br>
                  <font color="#cc0000"><b>友情提醒：</b></font><br>
                  <br>
                  你还没注册？或者没有登录？或者权限不够？这篇文章要求是本站符合权限要求的注册用户才能阅读！<br>
                  <br>
                  <%
						response.write "文章级别："
						response.write cINT(wzdj)
						response.write "级"
						%>
                  <br>
                  <%
						response.write "您的权限："
						response.write (Request.cookies(eChsys)("reglevel"))
						response.write "级"
						%>
                  <br>
                  <br>
                  如果你还没注册，请赶紧点此<a href="#" OnClick="adduser()" class=bottom><font color="#cc0000">注册</font></a>吧！<br>
                  <br>
                  如果你已经注册但还没登录，请赶紧点此<a href="#" OnClick="user_login()"  class=bottom><font color="#cc0000">登录</font></a>吧！<br>
                  <br>
                  <%end if%>
                  <%end if%>
                  <%end if%>
                  <%end if%>
                  <%end if%>
                  <%else%>
                  <%Response.Write arr_Content(Page-1)%>
                  <%end if%>
                  <%end if%>
                  <br>
                  <div align="right">
                    <%
url="E_ReadNews.asp?NewsId="&newsid
if MaxPages >0 then
	Response.write "<a class=black href='"& Url &"&page=1' title='第1页'>首页</a> "
	if Page-1 > 0 then
		Prev_Page = Page - 1
		Response.write "<a class=black href='"& Url &"&page="& Prev_Page &"' title='第"& Prev_Page &"页'>上一页</a> "
	end if

	for PageCounter=0 to MaxPages
		PageLink = PageCounter+1
		if PageLink <> Page Then
			Response.write "<a class=black href='"& Url &"&page="& PageLink &"'>["& PageLink &"]</a> "
		else
			Response.Write "<font color='#FF0000'><B>["& PageLink &"]</B></font> "
		end if
		If PageLink = MaxPages+1 Then Exit for
	Next
	if page <= Maxpages then
		bdd_Page = Page + 1
		Response.write "<a class=black href='" & Url & "&page=" & bdd_Page & "' title='第" & bdd_Page & "页'>下一页</A>"
	end if
	Response.write " <A class=black href='" & Url & "&page=" & Maxpages+1 & "' title='第"& Maxpages+1 &"页'>尾页</A>"
end if
%>
                </div></td>
                      </tr>
                  </table></td>
                </tr>
                <tr>
                  <td width="100%"  height="25">
                    <div align="center">
                      <!--#include file=attach.asp -->
                  </div></td>
                </tr>
                <tr>
                  <td height="22"> <b><IMG SRC="IMAGES/006.gif"> 上一篇：</b><a class=blacklink href="E_ReadNews.asp?NewsID=<%=getSideNewsID(E_BigClassID,newsid,-1)%>"><%=getnewstitle(getSideNewsID(E_BigClassID,newsid,-1))%></a><br>                  </td>
                </tr>
                <tr>
                  <td height="22"> <b><IMG SRC="IMAGES/006.gif"> 下一篇：</b><a  class=blacklink href="E_ReadNews.asp?NewsID=<%=getSideNewsID(E_BigClassID,newsid,1)%>"><%=getnewstitle(getSideNewsID(E_BigClassID,newsid,1))%></a> </td>
                </tr>
                
<!--
              <tr>
                  <td width=100% ><hr size=1 noshade style="color:#EFF5FA"></td>
                </tr>
                <tr>
                  <td  height=25>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="32%" valign="top" bgcolor="#EFF5FA">
                          <table width="100%" border="0" cellpadding="3" cellspacing="0" bgcolor="#EFF5FA">
                            <tr>
                              <td height="25" bgcolor="#E2EDF5" class="white_link">&nbsp;<B class="yellow_title">相关专题：</b></td>
                            </tr>
                            <tr>
                              <td bgcolor="#EFF5FA"><%set rs4=server.CreateObject("ADODB.RecordSet")
if SpecialID<>0 then
	rs4.Source="select * from "& db_EC_Special_Table &" where SpecialID=" & SpecialID
	rs4.Open,conn,1,3
	if not rs4.eof then%>
                                  <br>
                                  <a class=class href='Special_News.asp?SpecialID=<%=SpecialID%>'>·<%=CutStr(trim(rs4("E_SpecialName")),28)%></a>
                                  <%			
    end if
	rs4.Close
	set rs4=nothing

else

	Response.Write "<br><font color=red >·专题1信息无</font>"

	end if%>
                                  <br>
                                  <br>
                                  <%
	set rs4=server.CreateObject("ADODB.RecordSet")
	if specialid2<>0 then
		rs4.Source="select * from "& db_EC_Special_Table &" where SpecialID=" & SpecialID2
		rs4.Open,conn,1,3
		if not rs4.eof then %>
                                  <a class=class href='Special_News.asp?SpecialID=<%=SpecialID2%>'>·<%=CutStr(trim(rs4("E_SpecialName")),28)%></a>
                              <%end if
		rs4.Close
		set rs4=nothing

else

	Response.Write "<font color=red >·专题2信息无</font>"

	end if
	%>                              </td>
                            </tr>
                        </table></td>
                        <td width="1%">&nbsp;</td>
                        <td width="34%" valign="top" bgcolor="#EFF5FA"><table width="100%" border="0" cellspacing="0" cellpadding="3">
                            <tr>
                              <td height="25" bgcolor="#E2EDF5" class="white_link">&nbsp;<strong class="yellow_title"> 热门文章：</strong></td>
                            </tr>
                            <%
set rs=server.CreateObject("ADODB.RecordSet")
if uselevel=1 then
if Request.cookies("key")="super" or Request.cookies("key")="typemaster" or Request.cookies("key")="bigmaster" or Request.cookies("key")="smallmaster" or Request.cookies("key")="check" then
rs.Source="select top 4 * from "& db_EC_News_Table &" where  checkked=1 order by click DESC,newsid desc"
end if
if Request.cookies("key")="" then
rs.Source="select top 4 * from "& db_EC_News_Table &" where (checkked=1 and newslevel=0) order by click DESC,newsid desc"
end if
if Request.cookies("key")="selfreg" then
	if Request.cookies("reglevel")=3 then
rs.Source="select top 4 * from "& db_EC_News_Table &" where (checkked=1 ) order by click DESC,newsid desc"
	end if
	if Request.cookies("reglevel")=2 then
rs.Source="select top 4 * from "& db_EC_News_Table &" where (checkked=1 ) order by click DESC,newsid desc"
	end if
	if Request.cookies("reglevel")=1 then
rs.Source="select top 4 * from "& db_EC_News_Table &" where (checkked=1 ) order by click DESC,newsid desc"
	end if
end if
else
rs.Source="select top 4 * from "& db_EC_News_Table &" where  checkked=1 order by click DESC,newsid desc"
end if
rs.Open rs.Source,conn,1,1
while not rs.EOF
title=trim(rs("title"))
title=replace(title,"<br>","")
%>
                            <tr>
                              <td> &nbsp;· <a class=middle href="E_ReadNews.asp?NewsID=<%=rs("NewsID")%>" target="_blank" title="<%=title%>"><font color="<%=rs("titlecolor")%>">
                                <%if len(title)>13 then%>
                                <%=left(title,13)%>
                                <%else%>
                                <%=title%>
                                <%end if%>
                              </font></a>[<%=rs("click")%>] </td>
                            </tr>
                            <%
rs.MoveNext
wend
rs.close
%>
                        </table></td>
                        <td width="1%">&nbsp;</td>
                        <td width="32%" valign="top" bgcolor="#EFF5FA">
                          <table width="100%" border="0" cellspacing="0" cellpadding="3">
                            <tr>
                              <td height="25" bgcolor="#E2EDF5" class="yellow_title">&nbsp;<B class="yellow_title">相关文章：</b></td>
                            </tr>
                            <%
Response.Write ""
Response.Write ""
if about<>""  then
	sql="select top 4 * from "& db_EC_News_Table &" where (about like '%" & about & "%' or title like '%" & about & "%') and checkked=1 order by newsid desc"
	set rs=conn.execute(sql)

	do while not rs.eof
		title=htmlencode4(trim(rs("title")))
		%>
                            <tr>
                              <td> <img src=images/goxp.gif> <a class=middle target=_top href='E_ReadNews.asp?NewsID=<%=rs("NewsID")%>'><%=Title%></a>
                                  <!--<font color=#666666>(<%=month(trim(rs("updateTime")))%>月<%=day(trim(rs("updateTime")))%>日)-->
                                  [<font color=#666666><%=rs("click")%></font>] </td>
                            </tr>
                            <% rs.movenext
	loop
	Response.Write "<tr><td align='right'><a class=lift  href='Result.asp?keyword=" & about &"'><img src='images/more.gif' border='0'></a></td></tr>"
	rs.close   
	set rs=nothing  
else
	Response.write "<tr><td><font color=red ><br>·没有相关文章</font></td></tr>"
end if
%>
                        </table></td>
                      </tr>
                  </table></td>
                </tr>

-->
                <%
set rs1=server.CreateObject("ADODB.RecordSet")
rs1.Source="select top "&reviewnum&" * from "& db_EC_Review_Table &" where NewsID=" & NewsID & " and passed=1 order by reviewid desc"
rs1.Open rs1.Source,conn,1,1
if rs1.EOF then  NoReview=1
	Response.Write "<tr><td width=100% ><hr size=1 noshade style=color:#EFF5FA></td></tr><tr><td height=8></td></tr>"
	Response.Write "<tr><td width=100% ><B></b></td></tr><tr><td height=8></td></tr>"
	%>
                <tr>
                  <td width="100%">
                    <%
	if NoReview then 						
		Response.Write "<font color=red ><b>相关评论无</b></font>"
	end if
	%>                  </td>
                </tr>
                <%
	if not NoReview then
		while not rs1.EOF
			author=server.HTMLEncode(trim(rs1("author")))
			email=server.HTMLEncode(trim(rs1("email")))
			reviewip=rs1("reviewip")
			content=trim(rs1("content"))
            content=replace(content,"table","ｔａｂｌｅ")
			content=replace(content,"TABLE","ｔａｂｌｅ")
			ContentLen=len(Content)
			%>
                <tr>
                  <td>
                    <table width="100%" border="0" cellspacing="0" cellpadding="5" style="table-layout:fixed; word-break:break-all">
                      <tr bgcolor="#EFEFEF">
                        <td width="322" bgcolor="#EFF5FA">发表人：<%=author%></td>
                        <td width="270" bgcolor="#EFF5FA">
                          <p align="right">
                            <%if Request.cookies(eChsys)("key")="super" or showip="1" then%>
                    IP：<%=reviewip%>
                    <%end if%>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="2">
                          <table width="100%" border="0" cellspacing="0" cellpadding="0" style="table-layout:fixed; word-break:break-all">
                            <tr>
                              <td>发表人邮件：<a href='mailto:<%=email%>'><%=email%></a></td>
                              <td align="right">发表时间：<%=rs1("updatetime")%></td>
                            </tr>
                        </table></td>
                      </tr>
                      <tr bgcolor="#FFFFFF">
                        <td colspan="2"><%
			If ContentLen=<50  then	
				DisplayContent=Content
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;" & displaycontent
										%></td>
                      </tr>
                      <%
			else
									%>
                      <tr bgcolor="#FFFFFF">
                        <td colspan="2"><%
				DisplayContent=replace(nohtml(trim(Content)),"&nbsp;","",1,-1,1) '获取表格中留言字段内容并替换格式符
				DisplayContent=replace(DisplayContent,vbcrlf,"",1,-1,1)
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"& CutStr(displaycontent,50) &"<a href='ReadView.asp?ReviewID=" & rs1("ReviewID") &"&NewsID=" & NewsID &"' target=_blank class=class>详细内容</a>"
			end if
										%></td>
                      </tr>
                  </table></td>
                </tr>
                <%
			rs1.MoveNext
		wend
	end if

	rs1.Close
	set rs1=nothing	
						%>
                <tr>
                  <td width="98%" height="28"></td>
                </tr>
    </table></td>
</tr>
<tr valign="top">
    <td background="Images/table_bg.gif">
        <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
          <%if reviewable="1" then%>
          <form method="POST" action="AddReview.asp" name=AddReview  onSubmit="return CheckFormAddReview();">
            <%if Request.cookies(eChsys)("username")<>"" then
		
				%>
            <tr>
              <td width="100%" align="center">
                <table border="0" cellspacing="0" width="100%" cellpadding="4">
                  <input type=hidden name=NewsID value=<%=NewsID%>>
                  <tr>
                    <td width="15%" align="left"><a name="send"><font color="#FF0000">*</font>用&nbsp;户&nbsp;名：</a></td>
                    <td width="35%"><input type="text" name="Author" size="30" value="<%=Request.cookies(eChsys)("UserName")%>" readonly></td>
                    <td width="15%" align="left"><font color="#FF0000">*</font>电子邮件：</td>
                    <td width="35%"><input type="text" name="email" size="30" value="<%=Request.cookies(eChsys)("UserEmail")%>" ></td>
                  </tr>
                  <tr>
                    <td colspan="4" align="left" valign="top"><div align="center"><font color="#FF0000">*</font>（100字以内）
                            <% if M_MAIN=1 then %>
                            <font color=red >签收文件</font>： <img class=None  src="images/watermark.gif" align="absmiddle" border="0" style="cursor:hand;" alt="签收文件" lANGUAGE="javascript" onClick="sign()"></div></td>
                    <% end if%>
                  </tr>
                  <tr>
                    <td width="70%" colspan="4" align="center" valign="top">
                      <script language="javascript">
									var bTextMode=false;
									document.write ('<iframe src="guestbox.asp?action=new" id="message" width="500" height="100" frameborder=yes scrolling=no align=center></iframe>')
									frames.message.document.designMode = "On";
									</script>                    </td>
                  </tr>
                  <tr>
                    <td width="70%" colspan="4" align="center" height="50">
                      <input type="hidden" name="editor" value="<%=editor%>">
                      <input name="passed" type="hidden" value="<%if reviewcheck="1" then%>1<%else%>0<%end if%>">
                      <input type="submit" value="发表评论" name="Submit3" OnClick="document.AddReview.content.value = frames.message.document.body.innerHTML;">
                      <input name="reset" type=reset value="重新填写">
                      <input type="hidden" name="content" value="">
                      <input type="hidden" name="title" value="评论：<%=title%>">                    </td>
                  </tr>
              </table></td>
            </tr>
            <%else%>
            <tr>
              <td width="100%" align="center">
                <table border="0" cellspacing="0" width="100%" cellpadding="4">
                  <input type=hidden name=NewsID value=<%=NewsID%>>
                  <tr>
                    <td width="15%" align="left"><a name="send"><font color="#FF0000">*</font>用&nbsp;户&nbsp;名：</a></td>
                    <td width="35%">
                      <input type="text" name="Author" size="18">
        &nbsp;游客：
                      <INPUT onclick=changedata() type=checkbox value=true  name=none></td>
                    <td width="15%" align="left"><font color="#FF0000">*</font>电子邮件：</td>
                    <td width="35%"><input name="email" type="text" size="18" value="">
&nbsp;游客：
              <INPUT onclick=changedata1() type=checkbox value=true  name=none1></td>
                  </tr>
                  <tr>
                    <td colspan="4" align="left" valign="top" ><div align="center"><font color="#FF0000">*</font></div></td>
                  </tr>
                  <tr>
                    <td width="70%" colspan="4" align="center" valign="top">
                      <script language="javascript">
									document.write ('<iframe src="guestbox.asp?action=new" id="message" width="500" height="100" frameborder=yes scrolling=no align=center></iframe>')
									frames.message.document.designMode = "On";
									</script>                    </td>
                  </tr>
                  <tr>
                    <td width="70%" colspan="4" align="center" height="34">
                      <input type="hidden" name="editor" value="<%=editor%>">
                      <input name="passed" type="hidden" value="<%if reviewcheck="1" then%>1<%else%>0<%end if%>">
                      <input type="submit" value="发表评论" name="Submit3" OnClick="document.AddReview.content.value = frames.message.document.body.innerHTML;">
                      <input name="reset2" type=reset value="重新填写">
                      <input type="hidden" name="content" value="">
                      <input type="hidden" name="title" value="评论：<%=title%>">                    </td>
                  </tr>
              </table></td>
            </tr>
            <%end if%>
          </form>
          <%end if%>
      </table>
        </td>
  </tr>
<tr valign="top">
  <td height="47" valign="bottom" background="Images/news_bottom.gif"><div align="center">
    <table width="780" height="33" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="210">&nbsp;</td>
		<!--
        <td width="304"><div align="center"><a href=Review.asp?NewsID=<%=NewsID%> target="_blank" class="bottom"><img src="images/icon1.gif"  border="0">发表、查看更多关于该信息的评论</a></div></td>
		-->
        <td width="95"><div align="center"><a href="javascript:window.print()" class="bottom">打印本页</a></div></td>
        <td width="82"><div align="center"><a href="javascript:window.close()" class="bottom">关闭</a></div></td>
        <td width="89">&nbsp;</td>
      </tr>
    </table>
  </div></td>
</tr>
</table>
<br><br>
</body>
</html>
<%end if%>
<%
conn.close
set conn=nothing
%> 