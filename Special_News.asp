<%@ Language=VBScript%>
<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="function.asp"-->
<!--#include file="char.inc"-->
<%
request_SpecialID=ChkRequest(Request.QueryString("SpecialID"),1)	'防注入
if request_specialid="" then
Response.Write "<script>alert('未指定参数');history.back()</script>"
response.end
else

if not IsNumeric(request_specialid) then
 response.write "<script>alert('非法参数');history.back()</script>"
 response.end
else

set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from "& db_EC_Special_Table &" where Specialid="&request_SpecialID&" order by SpecialID"
rs.Open rs.Source,conn,1,1
if rs.EOF then
Response.Write "<script>alert('该专题不存在');history.back()</script>"
else
E_SpecialName=rs("E_SpecialName")
specialzs=rs("specialzs")

rs.close
set rs=nothing
PageShowSize = 10            '每页显示多少个页
MyPageSize   = 20           '每页显示多少条新闻
	
If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
MyPage=1
Else
MyPage=Int(Abs(Request("page")))
End if
end if

set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from "& db_EC_Type_Table &" order by E_typeorder"
rs.Open rs.Source,conn,1,1

dim ArrayE_typeid(10000),ArrayE_typename(10000),Arraytypecontent(10000)
typeCount=rs.RecordCount
for i=1 to typeCount
ArrayE_typeid(i)=rs("E_typeid")
ArrayE_typename(i)=rs("E_typename")
Arraytypecontent(i)=rs("typecontent")
rs.MoveNext
next
rs.Close
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=E_SpecialName%>_专题信息__<%=eChSys%></title>
<LINK href=news.css rel=stylesheet>
</head>

<body topmargin="0">
<!--#include file="other_Top.asp"-->
<table width="780" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr valign="top">
    <td width="210" rowspan="2" bgcolor="F1F7FC">
      <div align="center">
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber2">
          <tr>
            <td height=50 valign="middle" background="IMAGES/type_dh.gif"><div align="center" class="yellow_title"><img src="Images/arrow_dh_1.gif" width="7" height="7"><font color="#ffffff" class="yellow_title">&nbsp;&nbsp;</font>专题图文</div></td>
          </tr>
          <tr>
            <td height=25 valign="top" bgcolor="F1F7FC"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td>
                    <%
set rs2=server.CreateObject("ADODB.RecordSet")
if uselevel=1 then
if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then
rs2.Source="select top " & 3 & "  * from "& db_EC_News_Table &" where SpecialID=" & request_SpecialID & " or SpecialID2=" & request_SpecialID & " and picnews=1 and checkked=1 order by NewsID DESC"
end if
if Request.cookies(eChsys)("key")="" then
rs2.Source="select top " & 3 & "  * from "& db_EC_News_Table &" where SpecialID=" & request_SpecialID & " or SpecialID2=" & request_SpecialID & " and picnews=1 and checkked=1 and newslevel=0 order by NewsID DESC"
end if
if Request.cookies(eChsys)("key")="selfreg" then
	if Request.cookies(eChsys)("reglevel")=3 then
rs2.Source="select top " & 3 & "  * from "& db_EC_News_Table &" where SpecialID=" & request_SpecialID & " or SpecialID2=" & request_SpecialID & " and picnews=1 and checkked=1 order by NewsID DESC"
	end if
	if Request.cookies(eChsys)("reglevel")=2 then
rs2.Source="select top " & 3 & "  * from "& db_EC_News_Table &" where SpecialID=" & request_SpecialID & " or SpecialID2=" & request_SpecialID & " and picnews=1 and checkked=1 order by NewsID DESC"
	end if
	if Request.cookies(eChsys)("reglevel")=1 then
rs2.Source="select top " & 3 & "  * from "& db_EC_News_Table &" where SpecialID=" & request_SpecialID & " or SpecialID2=" & request_SpecialID & " and picnews=1 and checkked=1 order by NewsID DESC"
	end if
end if
else
rs2.Source="select top " & 3 & "  * from "& db_EC_News_Table &" where SpecialID=" & request_SpecialID & " or SpecialID2=" & request_SpecialID & " and picnews=1 and checkked=1 order by NewsID DESC"
end if
rs2.Open rs2.Source,conn,3,1

If Not rs2.eof then

while not rs2.EOF
fileExt=lcase(getFileExtName(rs2("picname")))
%>
                    <table width="100%" border="0" cellspacing="0" cellpadding="3" align="center" style="TABLE-LAYOUT: fixed">
                      <tr>
                        <td style="WORD-WRAP: break-word"> <div align="center"><a class=middle href="E_ReadNews.asp?NewsID=<%=rs2("NewsID")%>" target=_blank title="<%=rs2("title")%>">
                          <%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
                          <img  src="<%=FileUploadPath & rs2("picname")%>" width="150"  border=0 align="left">
                          <%end if%>
                          <%if fileext="swf" then%>
                          <object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="150" border=0 align="left">
                            <param name=movie value="<%=FileUploadPath & rs2("picname")%>">
                            <param name=quality value=high>
                            <param name='Play' value='-1'>
                            <param name='Loop' value='0'>
                            <param name='Menu' value='-1'>
                            <embed src="<%=FileUploadPath & rs2("picname")%>" width="150"  border=0 align="left" pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
                          </object>
                          <%end if%>
                        </a> </div></td>
                      </tr>
                    </table>
                    <%
											rs2.MoveNext
										wend
									else
										Response.Write "<tr><td width=100% align=center height=18>暂无</td></tr>"
									end if
									rs2.close
									set rs2=nothing
									%>
                  </td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td width="100%" height=50 valign="middle" background="IMAGES/type_dh.gif">
            <p align="center" class="yellow_title"><img src="Images/arrow_dh_1.gif" width="7" height="7"><font color="#ffffff" class="yellow_title">&nbsp;&nbsp;</font>所有专题</td>
          </tr>
          <tr>
            <td width="100%" valign="top">
              <div align="center">
                <table width="100%" border="0" cellpadding="3" cellspacing="0">
                  <tr>
                    <td bgcolor="F1F7FC" > <br>
                        <%set rs2=server.CreateObject("ADODB.RecordSet")  '专题
rs2.Source="select Top " & top_sp & " * from "& db_EC_Special_Table &"  order by SpecialID DESC "
rs2.Open rs2.Source,conn,1,1
if not rs2.EOF then
while not rs2.EOF

TrString="&nbsp;·&nbsp;<a class=class href='Special_News.asp?SpecialID=" & rs2("SpecialID") &"'>" & trim(CutStr(htmlencode4(rs2("E_SpecialName")),22)) & "</a><br>"
Response.Write TrString

rs2.MoveNext
wend
%>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a class=class href=E_Special.asp>更多专题...</a>
              <%
else
Response.Write "<tr><td width=100% align=center height=18 >暂无专题</td></tr>"
end if
rs2.Close
set rs2=nothing

  
%>                    </td>
                  </tr>
                </table>
            </div></td>
          </tr>
        </table>
        <!--广告调用-->
    </div></td>
    <td width="360" rowspan="2">
      <div align="center">
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td height="25" background="Images/dh_bg.gif">&nbsp;&nbsp;&nbsp;<img src="Images/arrow_dh_1.gif" width="7" height="7">&nbsp;&nbsp;&nbsp;栏目导航<b>&nbsp;&nbsp;</b><span class="daohang"><a class="daohang" href="./" >网站首页</a>&gt;&gt;<a class=daohang href=E_Special.asp>专题</a></span></td>
          </tr>
          <tr>
            <td>
              <div align="left"><a class="daohang" href="./" > </a> <a class=daohang href="./" > </a> </div>
              <div align="center">
                <table width="100%" border="0" cellpadding="3" cellspacing="0">
                  <tr>
                    <td><div align="center"><b><font size=+2 color="#FF0000" face="楷体_GB2312"><%=htmlencode4(E_SpecialName)%></font></b></div></td>
                  </tr>
                  <tr>
                    <td width="90%" class="news"> &nbsp;&nbsp;&nbsp;&nbsp;<%=specialzs%><br>
                        <br></td>
                  </tr>
                </table>
            </div></td>
          </tr>
        </table> 
        <table border="0" cellpadding="0" cellspacing="0" width="100%">
          <tr>
            <td width=100% height="25" valign="middle" background="Images/dh_bg.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class=m_tittle>&nbsp;</font>本专题文章</td>
          </tr>
          <!------------------------------------------------------------------------------------------------------------->
          <%
set rs2=server.CreateObject("ADODB.RecordSet")
if uselevel=1 then
if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then
rs2.Source="select * from "& db_EC_News_Table &" where SpecialID=" & request_SpecialID & " or SpecialID2=" & request_SpecialID & " and checkked=1 order by NewsID DESC"
end if
if Request.cookies(eChsys)("key")="" then
rs2.Source="select * from "& db_EC_News_Table &" where SpecialID=" & request_SpecialID & " or SpecialID2=" & request_SpecialID & " and checkked=1 and newslevel=0 order by NewsID DESC"
end if
if Request.cookies(eChsys)("key")="selfreg" then
	if Request.cookies(eChsys)("reglevel")=3 then
rs2.Source="select * from "& db_EC_News_Table &" where SpecialID=" & request_SpecialID & " or SpecialID2=" & request_SpecialID & " and checkked=1 order by NewsID DESC"
	end if
	if Request.cookies(eChsys)("reglevel")=2 then
rs2.Source="select * from "& db_EC_News_Table &" where SpecialID=" & request_SpecialID & " or SpecialID2=" & request_SpecialID & " and checkked=1 order by NewsID DESC"
	end if
	if Request.cookies(eChsys)("reglevel")=1 then
rs2.Source="select * from "& db_EC_News_Table &" where SpecialID=" & request_SpecialID & " or SpecialID2=" & request_SpecialID & " and checkked=1 order by NewsID DESC"
	end if
end if
else
rs2.Source="select * from "& db_EC_News_Table &" where SpecialID=" & request_SpecialID & " or SpecialID2=" & request_SpecialID & " and checkked=1 order by NewsID DESC"
end if
rs2.Open rs2.Source,conn,3,1

If Not rs2.eof then
rs2.PageSize     = MyPageSize
MaxPages         = rs2.PageCount
rs2.absolutepage = MyPage
total            = rs2.RecordCount
i = 0
do until rs2.Eof or i = rs2.PageSize
if rs2("picname")<>"" then
img=" <font class=pic_word>[图]</font>"
else
img=""
end if
title=htmlencode4(trim(rs2("title")))

%>
          <tr>
            <td width="100%">
              <div align="center">
                <table width="100%" border="0" cellpadding="2" cellspacing="0">
                  <tr>
                    <td width="78%"><a class=middle href="E_ReadNews.asp?NewsID=<%=rs2("NewsID")%>" target=_blank title="<%=title%>"><img src="IMAGES/point.gif" width="10" height="7" border="0">&nbsp;&nbsp;<%=img%><font color="<%=rs2("titlecolor")%>"> <%=CutStr(title,28)%> </font></a> <font class=middle>&nbsp;</font></td>
                    <!--
				    <td width="13%" align="right"> 
                    <%if showauthor="1" then%>
                    <a href="User.asp?user=<%=rs2("author")%>" target="_blank" class=class><%=rs2("Author")%></a> 
					<% datetime="<font class=middle>[" & Month(rs2("UpdateTime"))  &"/"& Day(rs2("UpdateTime")) &"]</font>" %>
                    <%end if%>
					</td>
					-->
                    <td width="22%" align="left"><div align="right"><font class=middle><%=datetime%></font></div></td>
                  </tr>
                </table>
            </div></td>
          </tr>
          <%
rs2.MoveNext
i = i + 1
loop
				
%>
          <tr>
            <td width="100%">&nbsp; </td>
          </tr>
          <tr>
            <td width="100%" align=center>共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 页，每页 <%=MyPageSize%> 条
            <%
url="Special_News.asp?SpecialID=" & request_SpecialID				
PageNextSize=int((MyPage-1)/PageShowSize)+1
Pagetpage=int((total-1)/rs2.PageSize)+1

if PageNextSize >1 then
PagePrev=PageShowSize*(PageNextSize-1)
Response.write "<a class=black href='" & Url & "&page=" & PagePrev & "' title='上" & PageShowSize & "页'>上一翻页</a> "
Response.write "<a class=black href='" & Url & "&page=1' title='第1页'>页首</a> "
end if
if MyPage-1 > 0 then
Prev_Page = MyPage - 1
Response.write "<a class=black href='" & Url & "&page=" & Prev_Page & "' title='第" & Prev_Page & "页'>上一页</a> "
end if

if Maxpages>=PageNextSize*PageShowSize then
PageSizeShow = PageShowSize
Else
PageSizeShow = Maxpages-PageShowSize*(PageNextSize-1)
End if
If PageSizeShow < 1 Then PageSizeShow = 1
for PageCounterSize=1 to PageSizeShow
PageLink = (PageCounterSize+PageNextSize*PageShowSize)-PageShowSize
if PageLink <> MyPage Then
Response.write "<a class=black href='" & Url & "&page=" & PageLink & "'>[" & PageLink & "]</a> "
else
Response.Write "<B>["& PageLink &"]</B> "
end if
If PageLink = MaxPages Then Exit for
Next

if Mypage+1 <=Pagetpage  then
Next_Page = MyPage + 1
Response.write "<a class=black href='" & Url & "&page=" & Next_Page & "' title='第" & Next_Page & "页'>下一页</A>"
end if

if MaxPages > PageShowSize*PageNextSize then
PageNext = PageShowSize * PageNextSize + 1
Response.write " <A class=black href='" & Url & "&page=" & Pagetpage & "' title='第"& Pagetpage &"页'>页尾</A>"
Response.write " <a class=black href='" & Url & "&page=" & PageNext & "' title='下" & PageShowSize & "页'>下一翻页</a>"
End if

else
Response.write "<tr><td valign=top height=80>&nbsp;本专题暂无信息</td></tr>"				
End If
rs2.close

%>            </td>
          </tr>
        </table>
    </div></td>
    <td width="210" bgcolor="F1F7FC"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" id="AutoNumber2" style="border-collapse: collapse">
      <tr>
        <td width="100%" height=50 valign="middle" background="IMAGES/type_dh.gif"><div align="center" class="yellow_title"><img src="Images/arrow_dh_1.gif" width="7" height="7"><font color="#ffffff" class="yellow_title">&nbsp;&nbsp;</font>专题文章导读</div></td>
      </tr>
      <tr>
        <td height=56 valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" style="TABLE-LAYOUT: fixed">
            <tr>
              <td>
                <%
set rs2=server.CreateObject("ADODB.RecordSet")
if uselevel=1 then
if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then
rs2.Source="select Top " & 5 & " * from "& db_EC_News_Table &" where SpecialID=" & request_SpecialID & " or SpecialID2=" & request_SpecialID & " and checkked=1 order by NewsID DESC"
end if
if Request.cookies(eChsys)("key")="" then
rs2.Source="select Top " & 5 & " * from "& db_EC_News_Table &" where SpecialID=" & request_SpecialID & " or SpecialID2=" & request_SpecialID & " and checkked=1 and newslevel=0 order by NewsID DESC"
end if
if Request.cookies(eChsys)("key")="selfreg" then
	if Request.cookies(eChsys)("reglevel")=3 then
rs2.Source="select Top " & 5 & " * from "& db_EC_News_Table &" where SpecialID=" & request_SpecialID & " or SpecialID2=" & request_SpecialID & " and checkked=1 order by NewsID DESC"
	end if
	if Request.cookies(eChsys)("reglevel")=2 then
rs2.Source="select Top " & 5 & " * from "& db_EC_News_Table &" where SpecialID=" & request_SpecialID & " or SpecialID2=" & request_SpecialID & " and checkked=1 order by NewsID DESC"
	end if
	if Request.cookies(eChsys)("reglevel")=1 then
rs2.Source="select Top " & 5 & " * from "& db_EC_News_Table &" where SpecialID=" & request_SpecialID & " or SpecialID2=" & request_SpecialID & " and checkked=1 order by NewsID DESC"
	end if
end if
else
rs2.Source="select Top " & 5 & " * from "& db_EC_News_Table &" where SpecialID=" & request_SpecialID & " or SpecialID2=" & request_SpecialID & " and checkked=1 order by NewsID DESC"
end if
rs2.Open rs2.Source,conn,3,1

If Not rs2.eof then

do until rs2.Eof 
if rs2("picname")<>"" then
img="<img src='images/news_img.gif' border='0'>  "
else
img=""
end if
title=htmlencode4(trim(rs2("title")))

%>
              <tr>
                  <td width="100%" >
                    <div align="center">
                      <table width="96%" border="0" cellpadding="3" cellspacing="0">
                        <tr>
                          <td width="46%"><b><%=CutStr(rs2("title"),18)%></b><br>
                              <a class=middle href="E_ReadNews.asp?NewsID=<%=rs2("NewsID")%>" target=_blank title="<%=CutStr(rs2("title"),18)%>"> &nbsp;&nbsp;&nbsp;&nbsp;<%=CutStr(nohtml(rs2("Content")),80)%><br>
                            </a> <font class=middle>&nbsp;</font></td>
                        </tr>
                      </table>
                  </div></td>
              </tr>
                <%
rs2.MoveNext

loop
				
%>
                <%
else
Response.write "<tr><td valign=top height=80>&nbsp;本专题暂无信息</td></tr>"				
End If
rs2.close
%>
        </table></td>
      </tr>
    </table></td>
  </tr>

</table>
<!--#include file="other_bottom.asp"-->
</body>

</html>
<%end if
end if
%>