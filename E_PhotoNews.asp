<%@ Language=VBScript%>
<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="function.asp"-->
<!--#include file="char.inc"-->

<%
request_newsid=Request.QueryString("NewsID")

PageShowSize = 10            'ÿҳ��ʾ���ٸ�ҳ
MyPageSize   = 20           'ÿҳ��ʾ����������

If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
MyPage=1
Else
MyPage=Int(Abs(Request("page")))
End if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ͼƬ����__<%=eChSys%></title>
<LINK href="news.css" rel="stylesheet">
</head>
<body>
<!--#include file="other_Top.asp"-->
<table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="F1F7FC">
   <tr>
      <td height="25" background="Images/dh_bg.gif" class="daohang">&nbsp;&nbsp;&nbsp;&nbsp;<img src="Images/arrow_dh_1.gif" width="7" height="7">&nbsp;&nbsp;&nbsp;<span class="daohang">��Ŀ����&nbsp;&nbsp;��ǰλ�ã�<a class="daohang" href="./" >��վ��ҳ</a>&gt;&gt;ͼƬ����</span>
	  </td>
   </tr>
   <tr>
     <td height="25">
                      <%
set rs3=server.CreateObject("ADODB.RecordSet")
if uselevel=1 then
if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then
rs3.Source ="select * from "& db_EC_News_Table &" where picnews=1 and checkked=1  and picname is not null order by NewsID DESC"
end if
if Request.cookies(eChsys)("key")="" then
rs3.Source ="select * from "& db_EC_News_Table &" where picnews=1 and checkked=1   and picname is not null order by NewsID DESC"
end if
if Request.cookies(eChsys)("key")="selfreg" then
if Request.cookies(eChsys)("reglevel")=3 then
rs3.Source ="select * from "& db_EC_News_Table &" where picnews=1 and checkked=1   and picname is not null order by NewsID DESC"
end if
if Request.cookies(eChsys)("reglevel")=2 then
rs3.Source ="select * from "& db_EC_News_Table &" where picnews=1 and checkked=1   and picname is not null order by NewsID DESC"
end if
if Request.cookies(eChsys)("reglevel")=1 then
rs3.Source ="select * from "& db_EC_News_Table &" where picnews=1 and checkked=1   and picname is not null order by NewsID DESC"
end if
end if
else
rs3.Source ="select * from "& db_EC_News_Table &" where picnews=1 and checkked=1  and picname is not null order by NewsID DESC"
end if
rs3.Open rs3.Source,conn,1,1

if not rs3.EOF then
Rs3.PageSize     = MyPageSize
MaxPages         = Rs3.PageCount
Rs3.absolutepage = MyPage
total            = Rs3.RecordCount
i = 0

%>
                      <table width="99%" border="0" align="center" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
                        <!--ͼƬ������ʾ1-->
                        <% do until rs3.Eof or i = rs3.PageSize %>
                        <tr>
                          <%
fileExt=lcase(getFileExtName(rs3("picname")))
Content=htmlencode4(rs3("Content"))
%>
                          <% if not rs3.EOF then %>
                          <td width=25% align=center valign="top">
                              <% content=htmlencode4(rs3("Content"))
content=replace(content,"[---��ҳ---]","")%>
                              <a class=middle href="E_ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=CutStr(nohtml(Content),150)%>...">
                              <%if   rs3("picname")=("") then%>
                              <img  src="IMAGES/flashorno.gif" width="110" height="80" border=1 style=border-color:#000000 align=top>
                              <%else%>
                              <%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
                              <img  src="<%=FileUploadPath & rs3("picname")%>" width="110" height="80" border=1 style=border-color:#000000 align=top>
                              <%end if%>
                              <%if fileext="swf" then%>
                              <object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="110" height="80" border=0 >
                                <param name=movie value="<%=FileUploadPath & rs3("picname")%>">
                                <param name=quality value=high>
                                <param name='Play' value='-1'>
                                <param name='Loop' value='0'>
                                <param name='Menu' value='-1'>
                                <embed src="<%=FileUploadPath & rs3("picname")%>" width="110" height="80" border=0  pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
                              </object>
                              <%end if%>
                              <%end if%>
                              </a><br>
                              <br>
                              <a class=middle href="E_ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=htmlencode4(rs3("title"))%>"> <%=CutStr(htmlencode4(rs3("title")),18)%> </a> </td>
                          <%rs3.movenext
i = i + 1
end if
%>
                          <!--ͼƬ������ʾ2-->
                          <td width=25% align=center valign="top">
                            <% if not rs3.EOF then %>
                              <%
fileExt=lcase(getFileExtName(rs3("picname")))
Content=htmlencode4(rs3("Content"))
content=replace(content,"[---��ҳ---]","")%>
                              <a class=middle href="E_ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=CutStr(nohtml(Content),150)%>...">
                              <%if rs3("picname")=("") then%>
                              <img  src="IMAGES/flashorno.gif" width="110" height="80" border=1 style=border-color:#000000 align=top>
                              <%else%>
                              <%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
                              <img  src="<%=FileUploadPath & rs3("picname")%>" width="110" height="80" border=1 style=border-color:#000000 align=top>
                              <%end if%>
                              <%if fileext="swf" then%>
                              <object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="110" height="80" border=0 >
                                <param name=movie value="<%=FileUploadPath & rs3("picname")%>">
                                <param name=quality value=high>
                                <param name='Play' value='-1'>
                                <param name='Loop' value='0'>
                                <param name='Menu' value='-1'>
                                <embed src="<%=FileUploadPath & rs3("picname")%>" width="110" height="80" border=0  pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
                              </object>
                              <%end if%>
                              <%end if%>
                              </a><br>
                              <br>
                              <a class=middle href="E_ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=htmlencode4(rs3("title"))%>"> <%=CutStr(htmlencode4(rs3("title")),18)%> </a></td>
                          <%rs3.movenext
i = i + 1
end if
%>
                          <!--ͼƬ������ʾ3-->
                          <td width=25% align=center valign="top">
                            <% if not rs3.EOF then %>
                              <%
fileExt=lcase(getFileExtName(rs3("picname")))
Content=htmlencode4(rs3("Content"))
content=replace(content,"[---��ҳ---]","")%>
                              <a class=middle href="E_ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=CutStr(nohtml(Content),150)%>...">
                              <%if rs3("picname")=("") then%>
                              <img  src="IMAGES/flashorno.gif" width="110" height="80" border=1 style=border-color:#000000 align=top>
                              <%else%>
                              <%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
                              <img  src="<%=FileUploadPath & rs3("picname")%>" width="110" height="80" border=1 style=border-color:#000000 align=top>
                              <%end if%>
                              <%if fileext="swf" then%>
                              <object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="110" height="80" border=0 >
                                <param name=movie value="<%=FileUploadPath & rs3("picname")%>">
                                <param name=quality value=high>
                                <param name='Play' value='-1'>
                                <param name='Loop' value='0'>
                                <param name='Menu' value='-1'>
                                <embed src="<%=FileUploadPath & rs3("picname")%>" width="110" height="80" border=0  pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
                              </object>
                              <%end if%>
                              <%end if%>
                              </a><br>
                              <br>
                              <a class=middle href="E_ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=htmlencode4(rs3("title"))%>"> <%=CutStr(htmlencode4(rs3("title")),18)%> </a></td>
                          <%rs3.movenext
i = i + 1
end if
%>
                          <!--ͼƬ������ʾ4-->
                          <td width=25% align=center valign="top">
                            <% if not rs3.EOF then %>
                              <%
fileExt=lcase(getFileExtName(rs3("picname")))
Content=htmlencode4(rs3("Content"))
content=replace(content,"[---��ҳ---]","")%>
                              <a class=middle href="E_ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=CutStr(nohtml(Content),150)%>...">
                              <%if rs3("picname")=("") then%>
                              <img  src="IMAGES/flashorno.gif" width="110" height="80" border=1 style=border-color:#000000 align=top>
                              <%else%>
                              <%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
                              <img  src="<%=FileUploadPath & rs3("picname")%>" width="110" height="80" border=1 style=border-color:#000000 align=top>
                              <%end if%>
                              <%if fileext="swf" then%>
                              <object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="110" height="80" border=0 >
                                <param name=movie value="<%=FileUploadPath & rs3("picname")%>">
                                <param name=quality value=high>
                                <param name='Play' value='-1'>
                                <param name='Loop' value='0'>
                                <param name='Menu' value='-1'>
                                <embed src="<%=FileUploadPath & rs3("picname")%>" width="110" height="80" border=0  pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
                              </object>
                              <%end if%>
                              <%end if%>
                              </a><br>
                              <br>
                              <a class=middle href="E_ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=htmlencode4(rs3("title"))%>"> <%=CutStr(htmlencode4(rs3("title")),18)%> </a> </td>
                          <%rs3.movenext
i = i + 1
end if
%>
                        </tr>
                        <!--ͼƬ���н���-->
                        <%loop%>
                    </table>
                </tr>
                <tr>
                  <td>
                    <TABLE cellSpacing=0 cellPadding=3 width="100%" align=center border=0 >
                      <tr>
                        <td width="100%" align=center valign="middle">�� <%=total%> ������ǰ�� <%=Mypage%>/<%=Maxpages%> ҳ��ÿҳ <%=MyPageSize%> ��
                            <%
url="E_PhotoNews.asp?"& "&newsid=" & request_newsid										
PageNextSize=int((MyPage-1)/PageShowSize)+1
Pagetpage=int((total-1)/rs3.PageSize)+1

if PageNextSize >1 then
PagePrev=PageShowSize*(PageNextSize-1)
Response.write "<a class=black href='" & Url & "&page=" & PagePrev & "' title='��" & PageShowSize & "ҳ'>��һ��ҳ</a> "
Response.write "<a class=black href='" & Url & "&page=1' title='��1ҳ'>ҳ��</a> "
end if
if MyPage-1 > 0 then
Prev_Page = MyPage - 1
Response.write "<a class=black href='" & Url & "&page=" & Prev_Page & "' title='��" & Prev_Page & "ҳ'>��һҳ</a> "
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
Response.write "<a class=black href='" & Url & "&page=" & Next_Page & "' title='��" & Next_Page & "ҳ'>��һҳ</A>"
end if

if MaxPages > PageShowSize*PageNextSize then
PageNext = PageShowSize * PageNextSize + 1
Response.write " <A class=black href='" & Url & "&page=" & Pagetpage & "' title='��"& Pagetpage &"ҳ'>ҳβ</A>"
Response.write " <a class=black href='" & Url & "&page=" & PageNext & "' title='��" & PageShowSize & "ҳ'>��һ��ҳ</a>"
End if
else
Response.write "<tr><td align=center>&nbsp;����������Ϣ</td></tr>"
				
End If
Rs3.close
set rs3=nothing
%>
                        </td>
                      </tr>
                  </table></td>
                </tr>
</table> 
<!--#include file="other_bottom.asp" -->
</body>
</html>
