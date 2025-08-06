<%@ Language=VBScript%>
<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="function.asp"-->
<!--#include file="char.inc"-->
<%
dim linktype
linktype=checkstr(request("linktype"))
if linktype="" then
Response.Write "未指定参数"
else
if not IsNumeric(linktype) then
 response.write "<script>alert('非法参数');history.back()</script>"
 response.end
else

PageShowSize = 10            '每页显示多少个页
MyPageSize   = 20           '每页显示多少条新闻
	
If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
MyPage=1
Else
MyPage=Int(Abs(Request("page")))
End if
end if


%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>友情链接__<%=eChSys%></title>
<LINK href=news.css rel=stylesheet>
</head>
<body topmargin="0">
<!--#include file="other_Top.asp"-->
<table width="780" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr valign="top"> 
    <td width="210" bgcolor="#F1F7FC"> 
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" id="AutoNumber2" style="border-collapse: collapse" valign=top >
        <tr> 
          <td width="100%" height="50" valign="middle" background="Images/type_dh.gif"> 
          <p align="center" class="yellow_title"><img src="Images/arrow_dh_1.gif" width="7" height="7">&nbsp;&nbsp;链接类型</td>
        </tr>
        <tr> 
          <td width="100%" height="24" align="center" valign="middle"> 
            <a class=class href="morelink.asp?linktype=2"><br>首页LOGO链接</a><br>
			<a class=class href="morelink.asp?linktype=1"><br>其他LOGO链接</a><br>
			<a class=class href="morelink.asp?linktype=0"><br>首页文字链接</a><br>
			<a class=class href="morelink.asp?linktype=3"><br>其他文字链接</a><br>            </td>
        </tr>
    </table>    </td>
    <td> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="25" background="Images/dh_bg.gif">&nbsp;&nbsp;&nbsp;<img src="Images/arrow_dh_1.gif" width="7" height="7">&nbsp;&nbsp;<span class="daohang">&nbsp;栏目导航</span><span class="white_link"><b class="daohang">&nbsp;&nbsp; 
            </b></span><span class="daohang">当前位置：</span><span class="white_link"><a class="daohang" href="./" >网站首页</a></span><span class="daohang">&gt;&gt;友情链接</span></td>
        </tr>
        <tr> 
          <td height="25"><div align="center"><b> 
              <font color=red><a class=daohang href="./" > 
              <script language=javascript src=./zongg/ad.asp?i=14></script> 
              </a> </font></b></div></td>
        </tr>
        <tr> 
          <td height="25" background="Images/dh_bg.gif" class="daohang">&nbsp;</td>
        </tr>
        <tr> 
          <td height="108"> 
            <div align="center"> 
              <table width="97%"  border="0" align="center" cellpadding="0" cellspacing="0" id="AutoNumber3" style="border-collapse: collapse">
                <tr> 
                  <td valign=top>&nbsp;</td>
                </tr>
                <tr> 
                  <td valign=top><table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" height="51">
                      <%

set rs3=server.CreateObject("ADODB.RecordSet")

rs3.Source="select * from "& db_EC_Link_Table &" where (linktype="&linktype&" and pass=1) order by id DESC"
rs3.Open rs3.Source,conn,3,1

If Not rs3.eof then
Rs3.PageSize     = MyPageSize
MaxPages         = Rs3.PageCount
Rs3.absolutepage = MyPage
total            = Rs3.RecordCount
%>
                      <tr> 
                        <td width=100% height="20" valign="middle" > <p align="center">&nbsp;共 
                            <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 页，每页 <%=MyPageSize%> 
                            条</td>
                      </tr>
                      <tr> 
                        <td width="100%" height="27"> 
                          <%
i = 0
do until rs3.Eof or i = rs3.PageSize
%>
                          <table width="98%" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td width="5%" height="31" valign="middle">&nbsp;<img src="IMAGES/point.gif" width="10" height="7" border="0">&nbsp;                              </td>
                              <%
                              if  rs3("linktype")=0 or rs3("linktype")=3  then
							  %>
							  <td width="18%" valign="middle">文字链接</td>
							  <%
							  else
							  %>
<%
                              if  rs3("linktype")=2  then
							  %>
							  <td width="18%" valign="middle"><a class=middle href="<%=rs3("weburl")%>" target=_blank title="<%=trim(rs3("content"))%>"><img src="<%=rs3("logo")%>" width="88" height="31"border="0"></a></td>
							  <%
							  else
							  %>

							  <td width="18%" valign="middle"><a class=middle href="<%=rs3("weburl")%>" target=_blank title="<%=trim(rs3("content"))%>"><img src="<%=rs3("logo")%>" width="88" height="31"border="0"></a></td>
<%end if%><%end if%>

                              <td width="40%" valign="middle"><font class=middle>......<a class=middle href="<%=rs3("weburl")%>" target=_blank title="<%=trim(rs3("content"))%>"><%=trim(rs3("webname"))%></font></a></td>
                              <td width="27%" valign="middle"><font class=middle>(<%=trim(rs3("dateandTime"))%>)</font></td>
                            </tr>
                          </table>
                          <br> 
                          <%
rs3.MoveNext
i = i + 1
loop
%>                        </td>
                      </tr>
                      <tr> 
                        <td width="100%" align=center height="20" valign="middle">共 
                          <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 页，每页 <%=MyPageSize%> 
                          条 
                          <%
url="morelink.asp?linktype="&linktype&""
PageNextSize=int((MyPage-1)/PageShowSize)+1
Pagetpage=int((total-1)/rs3.PageSize)+1

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
Response.write "<tr><td align=center>&nbsp;本类暂无信息</td></tr>"
				
End If
Rs3.close
set rs3=nothing

%>                        </td>
                      </tr>
                    </table></td>
                </tr>
              </table>
            </div></td>
        </tr>
      </table>    </td>
  </tr>
</table>
<!--#include file="bottom.asp" -->
</body>
</html><%end if%>