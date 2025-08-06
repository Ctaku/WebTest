<!--#include file="function.asp"-->

<table width="610" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
<tr>
  <td height="49" valign="top" background="Images/recommend.gif"><table width="87%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td width="40%" height="39"></td>
      <td width="60%" valign="middle"><script language="JavaScript" type="text/javascript">marquee3();</script>
          <%
set rs9=server.CreateObject("ADODB.RecordSet") 
rs9.Source="select top 6 * from "& db_EC_News_Table &" where checkked=1 order by newsid desc"
rs9.Open rs9.Source,conn,1,1
do while not rs9.EOF
title=trim(rs9("title"))
newsurl="E_ReadNews.asp?NewsID=" & rs9("NewsID")
newswwwurl=rs9("titleface")
%>
          <a class="middle" target="_blank" href="<%if rs9("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" title="<%=htmlencode4(title)%>"> <img src="images/dd.gif" width="5" height="5" border="0">&nbsp;<font color="<%=rs9("titlecolor")%>"><%=CutStr(htmlencode4(rs9("title")),28)%></font>&nbsp;&nbsp;
          <%rs9.movenext
loop
rs9.close
set rs9=nothing%>
          
        </a></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td>
<table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="226" height="217" rowspan="3" valign="middle"><table width="226" height="217" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td background="Images/pic_bg.gif"><div align="center">
          <script type="text/javascript">

<%
set rs3=server.CreateObject("ADODB.RecordSet") 
rs3.Source="select top 5 *  from "& db_EC_News_Table &" where picnews=1 and goodnews=1 and checkked=1 and picname is not null order by NewsID DESC"
rs3.Open rs3.Source,conn,1,1  
for i=1 to rs3.RecordCount
 
response.Write "imgUrl"&i&"=""uploadfile/"&rs3("picname")&""";"&vbCrLf 
response.Write "imgtext"&i&"="""&CutStr(htmlencode4(rs3("title")),30)&""""&vbCrLf 
response.Write "imgLink"&i&"=escape(""E_ReadNews.asp?newsid="&rs3("Newsid")&""");"&vbCrLf 
rs3.movenext 
if i=5 then exit for ''那里的代码只定义了5个循环所以数据库里的只取5条就可以了 
next 

%>
 var focus_width=206
 var focus_height=176
 var text_height=18
 var swf_height = focus_height+text_height
 
 var pics=imgUrl1+"|"+imgUrl2+"|"+imgUrl3+"|"+imgUrl4+"|"+imgUrl5
 var links=imgLink1+"|"+imgLink2+"|"+imgLink3+"|"+imgLink4+"|"+imgLink5
 var texts=imgtext1+"|"+imgtext2+"|"+imgtext3+"|"+imgtext4+"|"+imgtext5
 
 document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ focus_width +'" height="'+ swf_height +'">');
 document.write('<param name="allowScriptAccess" value="sameDomain"><param name="movie" value="images/focus.swf"><param name="quality" value="high"><param name="bgcolor" value="#F0F0F0">');
 document.write('<param name="menu" value="false"><param name=wmode value="opaque">');
 document.write('<param name="FlashVars" value="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'">');
 document.write('<embed src="pixviewer.swf" wmode="opaque" FlashVars="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'" menu="false" bgcolor="#F0F0F0" quality="high" width="'+ focus_width +'" height="'+ focus_height +'" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />');  document.write('</object>'); 
            </script>
        </div></td>
      </tr>
    </table></td>
    <td width="374" height="200" align="left" valign="top"><center>
        <table width="98%" border="0" cellspacing="2" cellpadding="0">
		            <% set rs99=server.CreateObject("ADODB.RecordSet") 
	rs99.Source="select top 1 * from "& db_EC_News_Table &" where daodu=1 and  checkked=1 order by newsid desc"
	rs99.Open rs99.Source,conn,1,1
	title=trim(rs99("title"))%>
<tr>
            
</tr>
            <%
rs9.close
set rs9=nothing
%>
          <% set rs9=server.CreateObject("ADODB.RecordSet") 
	rs9.Source="select top 7 * from "& db_EC_News_Table &" where goodnews=1  and  checkked=1 order by newsid desc"
	rs9.Open rs9.Source,conn,1,1
	if rs9.EOF then
	Response.Write "<tr><td align=center>目前暂无推荐内容</td></tr>"
	end if
	do while not rs9.EOF
			if showyear=1 then
			datetime="<font class=newsdate>(" & year(rs9("UpdateTime"))  &"-"& Month(rs9("UpdateTime"))  &"-"& Day(rs9("UpdateTime")) &")</font>"
		else
			datetime="<font class=newsdate>(" & year(rs9("UpdateTime"))  &"-"& Month(rs9("UpdateTime"))  &"-"& Day(rs9("UpdateTime")) &")</font>"
			end if
	title=trim(rs9("title"))%>
          <tr>
            <td  height="20">&nbsp;&nbsp;<img src="Images/dd.gif" width="4" height="4">&nbsp;&nbsp;<a class="tj" target="_blank" href="E_ReadNews.asp?NewsId=<%=rs9("newsid")%>" title="<%=CutStr(nohtml(rs9("Content")),100)%>"> <%=CutStr(htmlencode4(title),38)%></a></td>
            <td  align="right"><%=datetime%></td>
          </tr>
		    <tr>
              <td height="1" colspan="2" background="IMAGES/bg_line.gif"></td>
              </tr>
          <%rs9.movenext
	loop
	rs9.close
	set rs9=nothing%>
        </table>
        </center></td>
  </tr>
</table>
   </td>
  </tr>
    <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<!--
<table width="610" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="5" colspan="2"><div align="center">
      <script language="JavaScript" src="./zongg/ad.asp?i=12" type="text/javascript"></script>
    </div></td>
  </tr>
   <tr>
    <td height="5" colspan="2"></td>
  </tr>
</table>
<table width="610" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="420" height="180" valign="top">
	  <table width="420" border="0" cellspacing="0" cellpadding="0"  id=Layer1>
        <tr>
          <td width="14" rowspan="3" valign="top"><img src="Images/daodu_bg_01.gif" width="15" height="99" /></td>
          <td width="391" valign="top"><img src="Images/daodu_bg_02.gif" width="392" height="53" border="0" usemap="#Map" /></td>
          <td width="15" rowspan="3" valign="top"><img src="Images/daodu_bg_03.gif" width="13" height="99" ></td>
        </tr>
        <tr>
          <td width="392" height="127" valign="top" background="Images/daodu_bg_04.gif"><div align="center">
            <div align="center">
              <table width="98%" border="0" align="center" cellpadding="0" cellspacing="3">
                <% set rs9=server.CreateObject("ADODB.RecordSet") 
			rs9.Source="select top 6 * from "& db_EC_News_Table &" where istop=1 and checkked=1 order by newsid desc"
			rs9.Open rs9.Source,conn,1,1
			if rs9.EOF then
	        Response.Write "<tr><td align=center>目前暂无固顶文章</td></tr>"
	        end if
			do while not rs9.EOF
			title=trim(rs9("title"))
			%>
                <tr>
                  <td width="79%">&nbsp; <img src="Images/dd.gif" width="4" height="4" />&nbsp; <a class="middle" target="_blank" href="E_ReadNews.asp?NewsId=<%=rs9("newsid")%>" title="<%=htmlencode4(title)%>"><%=CutStr(nohtml(rs9("title")),46)%></a>
                      
                      <% if rs9("titlesize")>=1 then %>
                      <a class="middle" href="<%=path%>Review.asp?NewsID=<%=rs9("NewsID")%>" target="_blank" > <font color="red">评</font></a>
                      <%end if %>
                      
                  </td>
                  <td width="21%"><%=datetime%></td>
                </tr>
                <tr>
                  <td height="1" colspan="2" background="IMAGES/bg_line.gif"></td>
                </tr>
                <%
			rs9.movenext
			loop
			rs9.close
			set rs9=nothing
			%>
              </table>
            </div>
          </div></td>
        </tr>
        <tr>
          <td valign="top"><img src="Images/daodu_bg_08.gif" width="392" height="11" /></td>
        </tr>
      </table>
	  <table width="420" border="0" cellspacing="0" cellpadding="0"  id="Layer2">
        <tr>
          <td width="14" rowspan="3" valign="top"><img src="Images/daodu_bg_01.gif" width="15" height="99" /></td>
          <td width="391" valign="top"><img src="Images/daodu_bg2_02.gif" width="392" height="53" border="0" usemap="#Map2" />
            <map name="Map2" id="Map2">
              <area onmouseover="javascript:showtb(Layer1);" shape="rect" coords="5,14,74,42" />
              <area onmouseover="javascript:showtb(Layer2);shape=&quot;rect&quot;" coords="92,15,161,43" />
              <area onmouseover="javascript:showtb(Layer3);" shape="rect" coords="186,15,250,42" />
              <area onmouseover="javascript:showtb(Layer4);" shape="rect" coords="273,14,339,42" />
            </map></td>
          <td width="15" rowspan="3" valign="top"><img src="Images/daodu_bg_03.gif" width="13" height="99"></td>
        </tr>
        <tr>
          <td width="392" height="127" valign="top" background="Images/daodu_bg_04.gif"><div align="center">
            -->
			<!--#include file="ReviewIndex.asp"-->
			<!--
          </div></td>
        </tr>
        <tr>
          <td valign="top"><img src="Images/daodu_bg_08.gif" width="392" height="11" /></td>
        </tr>
      </table>
	  <table width="420" border="0" cellspacing="0" cellpadding="0"  id="Layer3">
        <tr>
          <td width="14" rowspan="3" valign="top"><img src="Images/daodu_bg_01.gif" width="15" height="99"></td>
          <td width="391" valign="top"><img src="Images/daodu_bg3_02.gif" width="392" height="53" border="0" usemap="#Map3" /></td>
          <td width="15" rowspan="3" valign="top"><img src="Images/daodu_bg_03.gif" width="13" height="99"></td>
        </tr>
        <tr>
          <td height="127" valign="top" background="Images/daodu_bg_04.gif"><div align="center">
              <div align="center">
                <table width="98%" height="80" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#D1DDE3" id="AutoNumber1" style="border-collapse: collapse">
                  <tr class="TDtop1" align="center">
                    <td width="7%" height="21" bgcolor="#FFFFFF"><span class="top1">状态</span></td>
                    <td width="43%" bgcolor="#FFFFFF" class="top1">信件标题</td>
                    <td width="12%" bgcolor="#FFFFFF" class="top1">来信日期</td>
                  </tr>
                  <%
	set rs=server.CreateObject("ADODB.RecordSet")
	rs.Source="select top 5 * from "& db_EC_LeadMail_Table &"  order by LeadMailID desc"	
	rs.Open rs.Source,conn,1,1
	if not rs.EOF then
    Do While Not Rs.EOF
    topic=trim(rs("topic"))
    DisplayTopic=mid(Topic,1,18)
%>
                  <tr bgcolor="ffffff">
                    <td height="23" align="center" bgcolor="#FFFFFF" ><%if trim(rs("ReplyContent"))<>"" then %>
                        <font color="red" >已阅批</font>
                        <%else%>
                        <font color="red" >未阅批</font>
                        <%end if%>
                    </td>
                    <td align="left" bgcolor="#FFFFFF"> ·<a  class="middle" href='E_ReadLeadMail.asp?LeadMailID=<%=rs("LeadMailID")%>' target="_black"><%=htmlencode(DisplayTopic)%>...</a></td>
                    <td align="center" bgcolor="#FFFFFF"><%=year(rs("UpdateTime")) %>-<%=Month(rs("UpdateTime")) %>-<%=Day(rs("UpdateTime")) %></td>
                  </tr>
                  <%
Rs.Movenext
Loop
Rs.Close
Set Rs=Nothing
else
Response.Write "<tr><td height=25 colspan=3 bgcolor=ffffff align=center><font color=red>暂无来信！</font></td></tr>"
End If
%>
                </table>
              </div>
          </div></td>
        </tr>
        <tr>
          <td valign="top"><img src="Images/daodu_bg_08.gif" width="392" height="11" /></td>
        </tr>
      </table>
	  <table width="420" border="0" cellspacing="0" cellpadding="0"  id="Layer4">
        <tr>
          <td width="14" rowspan="3" valign="top"><img src="Images/daodu_bg_01.gif" width="15" height="99" /></td>
          <td width="391" valign="top"><img src="Images/daodu_bg4_02.gif" width="392" height="53" border="0" usemap="#Map4" /></td>
          <td width="15" rowspan="3" valign="top"><img src="Images/daodu_bg_03.gif" width="13" height="99" /></td>
        </tr>
        <tr>
          <td height="127" valign="top" background="Images/daodu_bg_04.gif"><div align="center">
              <div align="center">
                <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#B3C7D0" id="AutoNumber1" style="border-collapse: collapse">
                  <tr class="TDtop1" align="center">
                    <td width="7%" height="21" bgcolor="#FFFFFF"><span class="top1">状态</span></td>
                    <td width="43%" bgcolor="#FFFFFF" class="top1">发言标题</td>
                    <td width="12%" bgcolor="#FFFFFF" class="top1">发言日期</td>
                  </tr>
                  <%

	set rs=server.CreateObject("ADODB.RecordSet")
	rs.Source="select top 5 * from "& db_EC_Opinion_Table &"  order by OpinionID desc"
	rs.Open rs.Source,conn,1,1
	If Not rs.eof then
    Do While Not Rs.EOF
    topic=trim(rs("topic"))
    DisplayTopic=mid(Topic,1,18)
%>
                  <tr bgcolor="ffffff">
                    <td height="23" align="center" bgcolor="#FFFFFF" ><%if trim(rs("ReplyContent"))<>"" then %>
                        <font color="red" >已处理</font>
                        <%else%>
                        <font color="red" >未处理</font>
                        <%end if%>
                    </td>
                    <td align="left" bgcolor="#FFFFFF"> ·<a  class="middle" href='E_ReadOpinion.asp?OpinionID=<%=rs("OpinionID")%>' target="_black"><%=htmlencode(DisplayTopic)%>...</a></td>
                    <td align="center" bgcolor="#FFFFFF"><%=year(rs("UpdateTime")) %>-<%=Month(rs("UpdateTime")) %>-<%=Day(rs("UpdateTime")) %></td>
                  </tr>
                  <%
Rs.Movenext
Loop
Rs.Close
Set Rs=Nothing
else
Response.Write "<tr><td height=25 colspan=3 bgcolor=ffffff align=center><font color=red>暂无建言献策</font></td></tr>"
End If
%>
                </table>
              </div>
          </div></td>
        </tr>
        <tr>
          <td valign="top"><img src="Images/daodu_bg_08.gif" width="392" height="11" /></td>
        </tr>
      </table>
    </td>
    <td rowspan="2" valign="top"><table width="185" border="0" align="right" cellpadding="0" cellspacing="0">
      <tr>
        <td><img src="Images/q_bg1.gif" width="185" height="6" /></td>
        </tr>
      <tr>
        <td height="169" valign="top" background="images\q_bg2.gif"><table width="90%" border="0" align="center" cellpadding="0" cellspacing="4">
          <tr>
            <td height="30"><div align="center"><a href="E_LeadMail.asp" target="_blank"><img src="Images/q1.gif" width="144" height="40" border="0" /></a></div></td>
          </tr>
          <tr>
            <td><div align="center"><a href="E_Guestbook.asp" target="_blank"><img src="Images/q2.gif" width="144" height="40" border="0" /></a></div></td>
          </tr>
          <tr>
            <td><div align="center"><a href="Complain.asp" target="_blank"><img src="Images/q3.gif" width="144" height="40" border="0" /></a></div></td>
          </tr>
          <tr>
            <td><div align="center"><a href="E_Opinion.asp" target="_blank"><img src="Images/q4.gif" width="144" height="40" border="0" /></a></div></td>
          </tr>
        </table></td>
        </tr>
      <tr>
        <td><img src="Images/q_bg3.gif" width="185" height="11" /></td>
        </tr>
    </table></td>
  </tr>
</table>
-->
<map name="Map3">
      <area onmouseover=javascript:showtb(Layer1); shape="rect" coords="5,14,74,42">
	  <area onmouseover=javascript:showtb(Layer2);shape="rect" coords="92,15,161,43">
      <area onmouseover=javascript:showtb(Layer3); shape="rect" coords="186,15,250,42">
      <area onmouseover=javascript:showtb(Layer4); shape="rect" coords="273,14,339,42">
</map>

<map name="Map4">
      <area onmouseover=javascript:showtb(Layer1); shape="rect" coords="7,15,75,41">
	  <area onmouseover=javascript:showtb(Layer2);shape="rect" coords="92,15,161,43">
      <area onmouseover=javascript:showtb(Layer3); shape="rect" coords="184,14,249,41">
      <area onmouseover=javascript:showtb(Layer4); shape="rect" coords="275,15,339,42">
</map>
<!--
<SCRIPT>
       document.all.Layer2.style.display="none";
	   document.all.Layer3.style.display="none";
	   document.all.Layer4.style.display="none";
       document.all.Layer2.style.visibility="hidden";
	   document.all.Layer3.style.visibility="hidden";
	   document.all.Layer4.style.visibility="hidden";

function showtb(tbobj) {
       document.all.Layer1.style.display="none";
       document.all.Layer2.style.display="none";
       document.all.Layer3.style.display="none";
	   document.all.Layer4.style.display="none";

	tbobj.style.display="";
	tbobj.style.visibility="visible";

}
</SCRIPT>


-->
<map name="Map" id="Map">
<area onmouseover=javascript:showtb(Layer1); shape="rect" coords="8,14,72,43">
<area onmouseover=javascript:showtb(Layer2); shape="rect" coords="96,14,160,43">
<area onmouseover=javascript:showtb(Layer3); shape="rect" coords="182,15,250,44">
<area onmouseover=javascript:showtb(Layer4); shape="rect" coords="273,14,341,42">
</map>