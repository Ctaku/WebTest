<%@ Language=VBScript%>
<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="function.asp"-->
<!--#include file="char.inc"-->
<%
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
<title>ר��__<%=eChSys%></title>
<LINK href="news.css" rel="stylesheet">
</head>

<body>
<!--#include file="other_Top.asp"-->
<table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr valign="top">
    <td width="210" rowspan="2" bgcolor="EFE8E8">
        <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="F1F7FC">
                <tr>
                  <td width="100%" height="50" valign="middle" background="Images/type_dh.gif" class="yellow_title">
                    <div align="center"><img src="Images/arrow_dh_1.gif" width="7" height="7"><font color="#ffffff" class="yellow_title">&nbsp;&nbsp;</font>�����Ķ�����</div></td>
                </tr>
                <tr>
                  <td width="100%">
                            <%
set rs=server.CreateObject("ADODB.RecordSet")
if uselevel=1 then
if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then
rs.Source="select top "& top_txt &" * from "& db_EC_News_Table &" where  checkked=1 order by click DESC,newsid desc"
end if
if Request.cookies(eChsys)("key")="" then
rs.Source="select top "& top_txt &" * from "& db_EC_News_Table &" where (checkked=1 and newslevel=0) order by click DESC,newsid desc"
end if
if Request.cookies(eChsys)("key")="selfreg" then
	if Request.cookies(eChsys)("reglevel")=3 then
rs.Source="select top "& top_txt &" * from "& db_EC_News_Table &" where (checkked=1 ) order by click DESC,newsid desc"
	end if
	if Request.cookies(eChsys)("reglevel")=2 then
rs.Source="select top "& top_txt &" * from "& db_EC_News_Table &" where (checkked=1 ) order by click DESC,newsid desc"
	end if
	if Request.cookies(eChsys)("reglevel")=1 then
rs.Source="select top "& top_txt &" * from "& db_EC_News_Table &" where (checkked=1 ) order by click DESC,newsid desc"
	end if
end if
else
rs.Source="select top "& top_txt &" * from "& db_EC_News_Table &" where  checkked=1 order by click DESC,newsid desc"
end if
rs.Open rs.Source,conn,1,1
while not rs.EOF
title=htmlencode4(trim(rs("title")))

%>
&nbsp;�� <a class=middle href="E_ReadNews.asp?NewsID=<%=rs("NewsID")%>" target="_blank" title="<%=title%>"><font color="<%=rs("titlecolor")%>"> <%=CutStr(title,22)%> </font></a>&nbsp;<font class=click_word>[<%=rs("click")%>]</font><br>
                    <%
rs.MoveNext
wend
rs.close
%>
</td>
                </tr>
          <tr>
            <td height="50" align="center" background="IMAGES/type_dh.gif"><FONT class="yellow_title"><img src="Images/arrow_dh_1.gif" width="7" height="7"><font color="#ffffff" class="yellow_title">&nbsp;&nbsp;</font>������������</FONT></td>
          </tr>
                  <% 
 dim ii
ii = 0
if uselevel=1 then
if Request.cookies(eChsys)("key")="" then
rs.Source="select top " & top_txt & " * from "& db_EC_News_Table &" where year(updatetime)=year(now()) and month(updatetime)=month(now()) and checkked=1 and newslevel=0 order by click DESC"   'ѡ����
end if
if Request.cookies(eChsys)("key")="selfreg" then
	if Request.cookies(eChsys)("reglevel")="3" then
	rs.Source="select top " & top_txt & " * from "& db_EC_News_Table &" where year(updatetime)=year(now()) and month(updatetime)=month(now()) and checkked=1 order by click DESC"   'ѡ����
	end if
	if Request.cookies(eChsys)("reglevel")="2" then
	rs.Source="select top " & top_txt & " * from "& db_EC_News_Table &" where year(updatetime)=year(now()) and month(updatetime)=month(now()) and checkked=1 order by click DESC"   'ѡ����
	end if
	if Request.cookies(eChsys)("reglevel")="1" then
	rs.Source="select top " & top_txt & " * from "& db_EC_News_Table &" where year(updatetime)=year(now()) and month(updatetime)=month(now()) and checkked=1 order by click DESC"   'ѡ����
	end if
end if
if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then
rs.Source="select top " & top_txt & " * from "& db_EC_News_Table &" where year(updatetime)=year(now()) and month(updatetime)=month(now()) and checkked=1 order by click DESC"   'ѡ����
end if
else
rs.Source="select top " & top_txt & " * from "& db_EC_News_Table &" where year(updatetime)=year(now()) and month(updatetime)=month(now()) and checkked=1 order by click DESC"   'ѡ����
end if
rs.Open rs.Source,conn,1,1
	if rs.bof and rs.eof then 
		response.write "<td align=center>��</td>" 
	else 
	do while not rs.eof 
%>
                  <tr>
                    <td height=18> ��
                        <%if rs("picnews")=1 then%>
                        <font class=pic_word>[ͼ]</font>
                        <%end if%>
                        <a class=middle href="E_ReadNews.asp?NewsID=<%=rs("NewsID")%>" title="<%=htmlencode4(rs("title"))%>" target="_blank" > <%=CutStr(rs("title"),20)%></a><font class=click_word>[<%=rs("click")%>]</font> </td>
                  </tr>
                  <%  ii = ii + 1
    if ii>20 then exit do
	rs.movenext     
	loop
	end if  
	rs.close   
	set rs=nothing
%>
      </table></td>
    <td width="570">
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td height="25" background="Images/dh_bg.gif" class="daohang">&nbsp;&nbsp;&nbsp;&nbsp;<img src="Images/arrow_dh_1.gif" width="7" height="7">&nbsp;&nbsp;&nbsp;<span class="daohang">��Ŀ����&nbsp;<a class="daohang" href="./" > ��վ��ҳ</a>&gt;&gt;ר��</span></td>
          </tr>
          <tr>
            <td>
                <div align="center">
                  <script language=javascript src=zongg/ad.asp?i=14></script>
                </div></td>
          </tr>
                      <%
set rs2=server.CreateObject("ADODB.RecordSet")  'ר��
rs2.Source="select * from "& db_EC_Special_Table &"  order by SpecialID DESC "
rs2.Open rs2.Source,conn,3,1

if rs2.eof and rs2.bof then
response.write "<p align=center>&nbsp;&nbsp;��û���κ�ר��</p>"
else
end if

if not rs2.eof then

Rs2.PageSize     = MyPageSize
MaxPages         = Rs2.PageCount
Rs2.absolutepage = MyPage
total            = Rs2.RecordCount
%>
                      <tr>
                        <td width=100% height="25" valign="middle" background="Images/dh_bg.gif" class="white_link"></td>
                      </tr>
                      <%

i = 0
do until Rs2.Eof or i = Rs2.PageSize
%>
                      <tr>
                        <td width="100%" height="20">&nbsp;&nbsp;<img src="IMAGES/point.gif" width="10" height="7" border="0">&nbsp;<a class=middle href="Special_News.asp?SpecialID=<%=rs2("SpecialID")%>"><%=htmlencode4(trim(rs2("E_SpecialName")))%></a></td>
                      </tr>
                      <%
Rs2.MoveNext
i = i + 1
loop
%>
                      <tr>
                        <td width="100%" height="24" align=center>�� <%=Mypage%>/<%=Maxpages%> ҳ��ÿҳ <%=MyPageSize%> ��
                            <%
url="E_Special.asp?"
PageNextSize=int((MyPage-1)/PageShowSize)+1
Pagetpage=int((total-1)/Rs2.PageSize)+1

if PageNextSize >1 then
PagePrev=PageShowSize*(PageNextSize-1)
Response.write "<a class=black href='" & Url & "page=" & PagePrev & "' title='��" & PageShowSize & "ҳ'>��һ��ҳ</a> "
Response.write "<a class=black href='" & Url & "page=1' title='��1ҳ'>ҳ��</a> "
end if
if MyPage-1 > 0 then
Prev_Page = MyPage - 1
Response.write "<a class=black href='" & Url & "page=" & Prev_Page & "' title='��" & Prev_Page & "ҳ'>��һҳ</a> "
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
Response.write "<a class=black href='" & Url & "page=" & PageLink & "'>[" & PageLink & "]</a> "
else
Response.Write "<B>["& PageLink &"]</B> "
end if
If PageLink = MaxPages Then Exit for
Next

if Mypage+1 <=Pagetpage  then
Next_Page = MyPage + 1
Response.write "<a class=black href='" & Url & "page=" & Next_Page & "' title='��" & Next_Page & "ҳ'>��һҳ</A>"
end if

if MaxPages > PageShowSize*PageNextSize then
PageNext = PageShowSize * PageNextSize + 1
Response.write " <A class=black href='" & Url & "page=" & Pagetpage & "' title='��"& Pagetpage &"ҳ'>ҳβ</A>"
Response.write " <a class=black href='" & Url & "page=" & PageNext & "' title='��" & PageShowSize & "ҳ'>��һ��ҳ</a>"
End if

				
End If
Rs2.close

%>
                        </td>
                      </tr>
                  </table></td>
                </tr>
               
</table>
<!--#include file="other_bottom.asp"-->

</body>

</html>