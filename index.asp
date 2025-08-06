<%@ Language=VBScript%>
<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="function.asp"-->
<!--'#include file="char.inc"-->
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="黔东南银监分局">
<title>黔东南银监分局</title>
<LINK href="news.css" rel=stylesheet>
<link rel="Shortcut Icon" href=""><!--IE地址栏前换成自己的图标--> 
<link rel="Bookmark" href=""><!--可以在收藏夹中显示出你的图标-->
<noscript><iframe scr="*"></ifriame></noscript>
<script type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
<SCRIPT> 
<!-- 
window.defaultStatus="<%=gg1%>"; 
//--> 
</SCRIPT>
<script language="JavaScript">
<!--
if (self != top) top.location.href = window.location.href
//-->
</script>
<script language=JavaScript>
<!--
//
var version = "other"
browserName = navigator.appName;   
browserVer = parseInt(navigator.appVersion);
if (browserName == "Netscape" && browserVer >= 3) version = "n3";
else if (browserName == "Netscape" && browserVer < 3) version = "n2";
else if (browserName == "Microsoft Internet Explorer" && browserVer >= 4) version = "e4";
else if (browserName == "Microsoft Internet Explorer" && browserVer < 4) version = "e3";
function marquee1()
{
	if (version == "e4" | version == "n3")
	{
		document.write("<marquee style='BOTTOM: 0px; FONT-WEIGHT: 100px; HEIGHT:120px;  TEXT-ALIGN: left; TOP: 0px' id='news' scrollamount='1' scrolldelay='10' behavior='loop' direction='up' border='0' onmouseover='this.stop()' onmouseout='this.start()'>")
	}
}

function marquee2()
{
	if (version == "e4" | version == "n3")
	{
		document.write("</marquee>")
	}
}

function marquee_logo_news()
{
	if (version == "e4")
	{
		document.write("<marquee style='BOTTOM: 0px; FONT-WEIGHT: 1200px; HEIGHT:31px;  TEXT-ALIGN: left; TOP: 0px' id='link_map' scrollamount='2' scrolldelay='10' behavior='alternate' direction='right' border='0' onmouseover='this.stop()' onmouseout='this.start()'>")
	}
}

function marquee3()
{
	if (version == "e4" | version == "n3")
	{
		document.write("<marquee direction='left' border='0' onmouseover='this.stop()' onmouseout='this.start()'>")
	}
}

function marquee4()
{
	if (version == "e4" | version == "n3")
	{
		document.write("</marquee>")
	}
}

function marquee5()
{
	if (version == "e4" | version == "n3")
	{
		document.write("<marquee style='BOTTOM: 0px; FONT-WEIGHT: 100px; HEIGHT:100px;  TEXT-ALIGN: left; TOP: 0px' id='news' scrollamount='1' scrolldelay='10' behavior='loop' direction='up' border='0' onmouseover='this.stop()' onmouseout='this.start()'>")
	}
}

function marquee6()
{
	if (version == "e4" | version == "n3")
	{
		document.write("</marquee>")
	}
}

//-->
</script>
<script language="JavaScript1.2">
function makevisible(cur,which){
if (which==0)
cur.filters.alpha.opacity=100
else
cur.filters.alpha.opacity=20
}
</script>

<script LANGUAGE="javascript">
<!--
function checkdata()
{
if (document.user.UserName.value=="")
	{
	  alert("对不起，请输入您的论坛用户名！");
	  document.user.UserName.focus();
	  return false;
	 }
else if (document.user.Password.value=="")
	{
	  alert("对不起，请输入您的论坛密码！");
	  document.user.Passwd.focus();
	  return false;
	 }
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
//-->
</script>
</head>
<body>
<%if E_announceview = 1 then%>
<script language="JavaScript">
  window.open("gonggao.asp","公告窗","width=300,height=400,toolbar=0,status=0,menubar=0,resize=0");
</script>
<%end if%>

<!-- 公告-->
<script language="JavaScript">
  //window.open("gonggao.asp","公告窗","width=300,height=400,toolbar=0,status=0,menubar=0,resize=0");
</script>
<!--#include file="Top.asp"-->

<%
dim E_typeid
dim E_typename
dim typecontent
set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from "& db_EC_Type_Table &" where E_typeview=1 order by E_typeorder"
rs.Open rs.Source,conn,1,1
i=1

Dim ArrayE_typeid(10000),ArrayE_typename(10000)
rs.close
set rs=nothing

set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from "& db_EC_Type_Table &" where E_typeview=1 order by E_typeorder"
rs.Open rs.Source,conn,1,1
i=1
Dim ArraytyID(10000),ArraytyName(10000),Arraytyview(10000)
if not rs.EOF then
rseof=1

while not rs.EOF
RecordCount=rs.RecordCount
tyID=rs("E_typeid")
tyName=trim(rs("E_typename"))
tycontent=rs("typecontent")
tyview=trim(rs("E_typeview"))
ArraytyID(i)=tyID
ArraytyName(i)=tyName
Arraytyview(i)=tyview
i=i+1
rs.MoveNext
wend
rs.close
%>



<table width="1002" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td width="190" height="226" valign="top" background="images\left_bg.gif">
	<!--#include file="index_left.asp"-->
    <div align="center"></div></td>
    <td width="622" valign="top"><div align="center">
       <!--#include file="E_daodu.asp"-->
<!--#include file="newtopphoto.asp"-->
<br>
<%
for i=1 to RecordCount
E_typeid=Arraytyid(i)
E_typename=ArraytyName(i)
%>
<table width="610" border="0" align="center" cellpadding="0" cellspacing="0" id="AutoNumber6" style="border-collapse: collapse">
  <tbody>
    <tr>
      <td width="379" valign="top"><table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber12">
        <tbody>
          <tr>
            <td height="43" colspan="2" valign="top" background="images/type_daohang.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="33%" height="26"><div align="right"><font><A class=top1 HREF="E_Type.asp?E_typeid=<%=E_typeid%>"><%=E_typename%></A></font></div></td>
                <td width="67%"><div align="right"><a class="class" href="E_Type.asp?E_typeid=<%=E_typeid%>"><img src="Images/more.gif" alt="更多<%=E_typename%>" width="32" height="6" border="0" /></a></div></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td width="90%" valign="top"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber11">
              <% dim rs3
	set rs3=server.CreateObject("ADODB.RecordSet")
	if uselevel=1 then
		if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then
			rs3.Source="select top " & top_news & " * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &" and checkked=1) order by newsid DESC"
		end if
		if Request.cookies(eChsys)("key")="" then
			rs3.Source="select top " & top_news & " * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &" and checkked=1) order by newsid DESC"
		end if
		if Request.cookies(eChsys)("key")="selfreg" then
			if Request.cookies(eChsys)("reglevel")=3 then
				rs3.Source="select top " & top_news & " * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &" and checkked=1 ) order by newsid DESC"
			end if
			if Request.cookies(eChsys)("reglevel")=2 then
				rs3.Source="select top " & top_news & " * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &" and checkked=1 ) order by newsid DESC"
			end if
			if Request.cookies(eChsys)("reglevel")=1 then
				rs3.Source="select top " & top_news & " * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &" and checkked=1 ) order by newsid DESC"
			end if
		end if
	else
		rs3.Source="select top " & top_news & " * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &" and checkked=1) order by newsid DESC"
	end if
	rs3.Open rs3.Source,conn,1,1
	while not rs3.EOF
		newsurl="E_ReadNews.asp?NewsID=" & rs3("NewsID")
		newswwwurl=rs3("titleface")
		fileExt=lcase(getFileExtName(rs3("picname")))
		if showyear=1 then
			datetime="<font class=newsdate>" & year(rs3("UpdateTime"))  &"/"& Month(rs3("UpdateTime"))  &"/"& Day(rs3("UpdateTime")) &"</font>"
		else
			datetime="<font class=newsdate>"& Month(rs3("UpdateTime"))  &"/"& Day(rs3("UpdateTime")) &"</font>"
		end if
		if rs3("picnews")=1 then
			if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then
			img="<font class=pic_word>[图]</font>"
				else
			img="<font class=pic_word>[媒]</font>"
			end if
		else
			img=""
		end if

		'增加小类名称程序开始
	if rs3("E_BigClassID")<>"" then 
		B=rs3("E_BigClassID")
		set rs22=server.createobject("adodb.recordset")
		sql="SELECT E_BigClassID,E_BigClassName FROM "& db_EC_BigClass_Table &" WHERE E_BigClassID="&B
		Set rs22=Conn.Execute(sql)
		A=rs22("E_BigClassName")
		rs22.close
		set rs22=Nothing
	else		A="本栏"
	end if

		'增加小类名称程序结束
		%>
              <tbody>
                <tr>
                  <td><table width="100%" border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
                    <tbody>
                      <tr>
                        <td height="22"><img src="IMAGES/point.gif" width="10" height="7">[<%=A%>]<a class="middle" href="<%if rs3("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" title="<%=htmlencode4(rs3("title"))%>" target="_blank"> <font color="<%=rs3("titlecolor")%>"> <%=CutStr(htmlencode4(rs3("title")),24)%> </font> </a><%=img%>
                                  <!--标题后评论提示-->
                                  <% if rs3("titlesize")>=1 then %>
                                  <A class=middle HREF="<%=path%>Review.asp?NewsID=<%=rs3("NewsID")%>" target="_blank" ><font color=red>评</font></A>
                                  <%end if %>
                                  <%=datetime%>
                                  <!--标题后评论提示-->
                                  <%''if showauthor="1" then%>
                                  <%''=rs3("Author")%>
                                  <%''end if%></td>
                      </tr>
                      <tr>
                        <td height="1" background="IMAGES/bg_line.gif"></td>
                      </tr>
                    </tbody>
                  </table></td>
                </tr>
                <%
								rs3.MoveNext
wend
%>
              </tbody>
            </table></td>
          </tr>
        </tbody>
      </table></td>
      <%i=i+1
              E_typeid=Arraytyid(i)
              E_typename=ArraytyName(i)
              if i<=RecordCount then
              %>
      <td  width="1" valign="top" background="images/t.gif" bgcolor="#F1F7FC"></td>
      <td width="380" valign="top"><table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber13">
        <tbody>
          <tr>
            <td  height="43" colspan="2" valign="top" background="images/type_daohang_r.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="33%" height="26"><div align="right"><font><A class=top1 HREF="E_Type.asp?E_typeid=<%=E_typeid%>"><%=E_typename%></A></font></div></td>
                <td width="67%"><div align="right"><a class="class" href="E_Type.asp?E_typeid=<%=E_typeid%>"><img src="Images/more.gif" alt="更多<%=E_typename%>" width="32" height="6" border="0" /></a></div></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td width="90%" valign="top"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber14">
              <%
set rs3=server.CreateObject("ADODB.RecordSet")
if uselevel=1 then
if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then
rs3.Source="select top " & top_news & " * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &" and checkked=1) order by newsid DESC"
end if
if Request.cookies(eChsys)("key")="" then
rs3.Source="select top " & top_news & " * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &" and checkked=1) order by newsid DESC"
end if
if Request.cookies(eChsys)("key")="selfreg" then
	if Request.cookies(eChsys)("reglevel")=3 then
	rs3.Source="select top " & top_news & " * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &" and checkked=1 ) order by newsid DESC"
	end if
	if Request.cookies(eChsys)("reglevel")=2 then
	rs3.Source="select top " & top_news & " * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &" and checkked=1 ) order by newsid DESC"
	end if
	if Request.cookies(eChsys)("reglevel")=1 then
	rs3.Source="select top " & top_news & " * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &" and checkked=1 ) order by newsid DESC"
	end if
end if
else
	rs3.Source="select top " & top_news & " * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &" and checkked=1) order by newsid DESC"
end if
rs3.Open rs3.Source,conn,1,1
while not rs3.EOF
newsurl="E_ReadNews.asp?NewsID=" & rs3("NewsID")
newswwwurl=rs3("titleface")
fileExt=lcase(getFileExtName(rs3("picname")))
if showyear=1 then
datetime="<font class=newsdate>" & year(rs3("UpdateTime"))  &"/"& Month(rs3("UpdateTime"))  &"/"& Day(rs3("UpdateTime")) &"</font>"
else
newsurl="E_ReadNews.asp?NewsID=" & rs3("NewsID")
newswwwurl=rs3("titleface")
datetime="<font class=newsdate>"& Month(rs3("UpdateTime"))  &"/"& Day(rs3("UpdateTime")) &"</font>"
end if

if rs3("picnews")=1 then
			if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then
			img="<font class=pic_word>[图]</font>"
				else
			img="<font class=pic_word>[媒]</font>"
			end if
			else
			img=""
		end if
%>
              <%
		'增加小类名称程序开始
	if rs3("E_BigClassID")<>"" then 
		B=rs3("E_BigClassID")
		set rs22=server.createobject("adodb.recordset")
		sql="SELECT E_BigClassID,E_BigClassName FROM "& db_EC_BigClass_Table &" WHERE E_BigClassID="&B
		Set rs22=Conn.Execute(sql)
		A=rs22("E_BigClassName")
		rs22.close
		set rs22=Nothing
	else
		A="本栏"
	end if

		'增加小类名称程序结束
%>
              <tbody>
                <tr>
                  <td><table width="100%" border="0" cellpadding="2" cellspacing="0">
                    <tbody>
                      <tr>
                        <td height="22"><img src="IMAGES/point.gif" width="10" height="7">[<%=A%>] <a class="middle" href="<%if rs3("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" title="<%=htmlencode4(rs3("title"))%>" target="_blank"> <font color="<%=rs3("titlecolor")%>"> <%=CutStr(htmlencode4(rs3("title")),24)%> </font> </a><%=img%>
                                  <!--标题后评论提示-->
                                  <% if rs3("titlesize")>=1 then %>
                                  <A class=middle HREF="<%=path%>review.asp?NewsID=<%=rs3("NewsID")%>" target="_blank" ><font color=red>评</font></A>
                                  <%end if %>
                                  <%=datetime%>
                                  <!--标题后评论提示-->
                                  <%''if showauthor="1" then%>
                                  <%''=rs3("Author")%>
                                  <%''end if%>                        </td>
                      </tr>
                      <tr>
                        <td background="IMAGES/bg_line.gif" height="1"></td>
                      </tr>
                    </tbody>
                  </table></td>
                </tr>
                <%rs3.MoveNext
wend
%>
              </tbody>
            </table></td>
          </tr>
        </tbody>
      </table></td>
      <%end if%>
    </tr>
  </tbody>
</table>
<%
next
rs3.close
set rs3=nothing
%>
    </div></td>
    <td width="190" rowspan="2" valign="top" background="images\right_bg.gif"><div align="center">
      <!--#include file="index_right.asp" -->
    </div></td>
  </tr>
</table>
<table width="1002" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="782" height="18" valign="top" bgcolor="#FFFFFF"><center>
      <!--#include file="FriendSite.asp" -->
      <%else
Response.Write "<table width=760 align=center border=0 height=50 bgcolor=#F1F7FC><tr><td align=center>暂 无 文 章 类 别，请 <a href=login.asp>登 陆</a> 后 添 加！</td></tr></table>"
end if%>
  </center>
	</td>
  </tr>
</table>
<!--#include file="Bottom.asp" -->
</body>
</HTML>