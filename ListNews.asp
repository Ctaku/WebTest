<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc" -->
<!--#include file="chkuser.asp" -->
<%
E_SmallClassID=ChkRequest(request("E_SmallClassID"),1)	'��ע��
if E_SmallClassID="" then
	Show_Err("δָ������<br><br><a href='javascript:history.back()'>��ȥ����</a>")
	response.end
else
	if not IsNumeric(E_SmallClassID) then
		Show_Err("�Ƿ�����<br><br><a href='javascript:history.back()'>��ȥ����</a>")
		response.end
	end if
end if
	dim E_smallclassname,E_BigClassName,E_typename,E_BigClassID,E_typeid
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_EC_SmallClass_Table &" where E_SmallClassID="&E_SmallClassID
	rs.open sql,conn,1,3
	E_smallclassname=rs("E_smallclassname")
	E_BigClassID=rs("E_BigClassID")
	master6=rs("smallclassma")
	if master6<>"" then
	smallmaster6=split(master6,"|")
	 for i=0 to ubound(smallmaster6)
		if request.cookies(eChsys)("username")=trim(smallmaster6(i)) then
		smallclassmaster6=true
		exit for
		else
		smallclassmaster6=false
		end if
	 next
	end if
	set rs1=server.createobject("adodb.recordset")
	sql1="select * from "& db_EC_BigClass_Table &" where E_BigClassID="&rs("E_BigClassID")
	rs1.open sql1,conn,1,3
	E_BigClassName=rs1("E_BigClassName")
	E_typeid=rs1("E_typeid")
	master5=rs1("bigclassmaster")
	if master5<>"" then
	bigmaster5=split(master5,"|")
	 for i=0 to ubound(bigmaster5)
		if request.cookies(eChsys)("username")=trim(bigmaster5(i)) then
		bigclassmaster5=true
		exit for
		else
		bigclassmaster5=false
		end if
	 next
	end if
	set rs2=server.createobject("adodb.recordset")
	sql2="select * from "& db_EC_Type_Table &" where E_typeid="&E_typeid
	rs2.open sql2,conn,1,3
	E_typename=rs2("E_typename")
	master4=rs2("typemaster")
	if master4<>"" then
	tmaster4=split(master4,"|")
	 for i=0 to ubound(tmaster4)
		if request.cookies(eChsys)("username")=trim(tmaster4(i)) then
		typemaster4=true
		exit for
		else
		typemaster4=false
		end if
	 next
	end if
	rs2.close
	set rs2=nothing
	rs1.close
	set rs1=nothing
	rs.close
	set rs=nothing
	if typemaster4=true or request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="check" or bigclassmaster5=true or smallclassmaster6=true or request.cookies(eChsys)("key")="selfreg" then
	
	dim jingyong
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_User_Table &" where  "& db_User_Name &"='"&request.cookies(eChsys)("username")&"'"
	rs.open sql,ConnUser,1,3
	jingyong=rs("jingyong")
	rs.close
	set rs=nothing
	
	if request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="check" or request.cookies(eChsys)("key")="bigmaster" or request.cookies(eChsys)("key")="smallmaster" or (request.cookies(eChsys)("key")="selfreg" and jingyong=0) or request.cookies(eChsys)("key")="typemaster" then%>
	<%
	PageShowSize = 10            'ÿҳ��ʾ���ٸ�ҳ
	MyPageSize   = 16           'ÿҳ��ʾ����������
		
	If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
	MyPage=1
	Else
	MyPage=Int(Abs(Request("page")))
	End if
	
	
	set rs=server.CreateObject("ADODB.RecordSet")
	if request.cookies(eChsys)("KEY")="super" or request.cookies(eChsys)("key")="check" or request.cookies(eChsys)("key")="bigmaster" or request.cookies(eChsys)("key")="typemaster" then
	rs.Source="select * from "& db_EC_News_Table &" where E_SmallClassID=" & E_SmallClassID & " order by NewsID desc"
	end if
	if request.cookies(eChsys)("key")="selfreg" then
	rs.Source="select * from "& db_EC_News_Table &" where E_SmallClassID=" & E_SmallClassID & " and editor='" & request.cookies(eChsys)("username") & "' order by NewsID desc"
	end if
	if request.cookies(eChsys)("key")="smallmaster" then
		if smallmana=1 then
	rs.Source="select * from "& db_EC_News_Table &" where E_SmallClassID=" & E_SmallClassID & " order by NewsID desc"
		else
	rs.Source="select * from "& db_EC_News_Table &" where E_SmallClassID=" & E_SmallClassID & " and editor='" & request.cookies(eChsys)("username") & "' order by NewsID desc"
		end if
	end if
	rs.Open rs.Source,conn,3,1
	
	If Not rs.eof then
	rs.PageSize     = MyPageSize
	MaxPages         = rs.PageCount
	rs.absolutepage = MyPage
	total            = rs.RecordCount
	%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - ��������</title>
<style type="text/css">
<!--
.STYLE2 {FONT-FAMILY: Verdana, Arial, ����; font-weight:bold; background-color: #FF9933; border: #FFFFFF; background-image:url(Images/admin_topbg.gif); font-size: 9pt;}
-->
</style>
</head>
<body topmargin="0">
<div align="center">
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFDFBF" width="100%" id="AutoNumber1">
<tr>
  <td height="20" colspan=8 align="left"  class="TDtop"><a href=E_TypeManage.asp>ȫ��������</a></td>
</tr>
<tr>
    <td colspan=8 width="100%" height="30" align="left"  class="TDtop1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src=images/menu1.gif><b><a href="E_BigClassManage.asp?E_typeid=<%=E_typeid%>"><font color=red><%=E_typename%></font></a></b>&nbsp;&nbsp;&nbsp;&nbsp;����������
<%
Set rs1 = Server.CreateObject("ADODB.Recordset")
sql1 ="SELECT  * From "& db_EC_Type_Table &" where E_typeid<>"&E_typeid&" order by E_typeorder"
RS1.open sql1,Conn,3,3
do while not rs1.EOF
master1=rs1("typemaster")
if master1<>"" then
tmaster1=split(master1,"|")
 for i=0 to ubound(tmaster1)
	if request.cookies(eChsys)("username")=trim(tmaster1(i)) then
	typemaster1=true
	exit for
	else
	typemaster1=false
	end if
 next
end if
if typemaster1=true or request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="check" or request.cookies(eChsys)("key")="selfreg" then

%><a href=E_BigClassManage.asp?E_typeid=<%=rs1("E_typeid")%>><%=rs1("E_typename")%></a>
<%end if
rs1.movenext
loop
rs1.close
set rs1=nothing%>
      <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src=images/menu1.gif><b><a href="E_SmallClassManage.asp?E_BigClassID=<%=E_BigClassID%>"><font color=red><%=E_BigClassName%></font></a></b>
      &nbsp;&nbsp;&nbsp;&nbsp;�������ࣺ
<%
Set rs1 = Server.CreateObject("ADODB.Recordset")
sql1 ="SELECT  * From "& db_EC_BigClass_Table &" where E_typeid="&E_typeid&" and E_BigClassID<>"&E_BigClassID&" order by E_bigclassorder"
RS1.open sql1,Conn,3,3
do while not rs1.EOF
master=rs1("bigclassmaster")
if master<>"" then
bigmaster=split(master,"|")
 for i=0 to ubound(bigmaster)
	if request.cookies(eChsys)("username")=trim(bigmaster(i)) then
	bigclassmaster=true
	exit for
	else
	bigclassmaster=false
	end if
 next
end if
if (bigclassmaster=true and request.cookies(eChsys)("key")="bigmaster") or request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="check" or (request.cookies(eChsys)("key")="typemaster" and typemaster4=true) or request.cookies(eChsys)("key")="selfreg" then
%><a href=E_SmallClassManage.asp?E_BigClassID=<%=rs1("E_BigClassID")%>><%=rs1("E_BigClassName")%></a>
<%end if
rs1.movenext
loop
rs1.close
set rs1=nothing%>
      <br>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src=images/menu1.gif><b><a href="listnews.asp?E_SmallClassID=<%=E_SmallClassID%>"><font color=red><%=E_smallclassname%></font></a></b>
      &nbsp;&nbsp;&nbsp;&nbsp;����С�ࣺ
<%if request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="check" or request.cookies(eChsys)("key")="typemaster" or request.cookies(eChsys)("key")="bigmaster" or request.cookies(eChsys)("key")="selfreg" then
Set rs1 = Server.CreateObject("ADODB.Recordset")
sql1 ="SELECT  * From "& db_EC_SmallClass_Table &" where E_BigClassID="&E_BigClassID&" and E_SmallClassID<>"&E_SmallClassID&" order by smallclassorder"
RS1.open sql1,Conn,3,3
do while not rs1.EOF
%><a href=ListNews.asp?E_SmallClassID=<%=rs1("E_SmallClassID")%>><%=rs1("E_smallclassname")%></a>
<%rs1.movenext
loop
rs1.close
set rs1=nothing
else
Set rs1 = Server.CreateObject("ADODB.Recordset")
sql1 ="SELECT * From "& db_EC_SmallClass_Table &" where E_SmallClassID<>"&E_SmallClassID&" order by SmallClassorder"
RS1.open sql1,Conn,3,3
do while not rs1.EOF
master2=rs1("smallclassma")
if master2<>"" then
smallmaster=split(master2,"|")
 for i=0 to ubound(smallmaster)
	if request.cookies(eChsys)("username")=trim(smallmaster(i)) then
	smallclassmaster=true
	exit for
	else
	smallclassmaster=false
	end if
 next
end if
if smallclassmaster=true then
%><a href=listnews.asp?E_SmallClassID=<%=rs1("E_SmallClassID")%>><%=rs1("E_smallclassname")%></a>
<%end if
rs1.movenext
loop
rs1.close
set rs1=nothing
end if
%></td>
</tr>
<tr>
    <td colspan=8 align=center height=25>�� <%=total%> ������ǰ�� <%=Mypage%>/<%=Maxpages%> 
            ҳ��ÿҳ <%=MyPageSize%> ����ID��Ϊ��ɫ����ʾ������δ������ˣ�&nbsp;&nbsp;<a href='<%if eWebEditor="1" then%>NewsAdd1_eWebEditor.asp<%else%>NewsAdd1.asp<%end if%>?E_SmallClassID=<%=E_SmallClassID%>'><U>����������</U></a></td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
<td width="7%">ID��</td>
<td width="39%">�� &nbsp;&nbsp;&nbsp;��</td>
<td width="16%">��&nbsp;&nbsp;&nbsp;��</td>
<td width="6%">�༭</td>
<td width="6%">���</td>
<td width="6%">�Ƽ�</td>
<td width="6%">�̶�</td>
<td width="6%">ɾ��</td>		
</tr>
<%
for i=1 to rs.PageSize
if not rs.EOF then
title=htmlencode4(trim(rs("title")))
%>
<%
if request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="bigmaster" or request.cookies(eChsys)("key")="smallmaster" or (request.cookies(eChsys)("KEY")="check" and checkmod="1") or request.cookies(eChsys)("key")="typemaster" or request.cookies(eChsys)("key")="selfreg" then
%>	
<tr>
<td align=center ><%if rs("checkked")<>1 then%><font color=red><%end if%><%=rs("NewsID")%></font>��</td>
<td>&nbsp;<a href='E_ReadNews.asp?NewsID=<%=rs("NewsID")%>' target=_blank title="�鿴���¡�<%=Title%>��"><font color="<%=rs("titlecolor")%>"><%=left(title,40)%></font></a><%if rs("picname")<>"" then%><font color=red>[ͼ]<%end if%><%if request.cookies(eChsys)("key")="super" and request.cookies(eChsys)("purview")="99999" then%><font color=red>(<%=rs("editor")%>)</font><%end if%></td>
<td align=center ><%=rs("UpdateTime")%>��</td>
<td align=center ><a href='E_NewsEdit.asp?Newsid=<%=rs("newsid")%>'>�༭</a></td>
<td align=center >
	<%IF not(request.cookies(eChsys)("KEY")="super" or request.cookies(eChsys)("KEY")="check" or request.cookies(eChsys)("key")="typemaster" or request.cookies(eChsys)("key")="bigmaster") THEN%>
		���
	<%else%>
		<a href='NewsCheck.asp?Newsid=<%=rs("newsid")%>'><%if rs("checkked")<>1 then%>���<%else%>����<%end if%></a>
	<%end if%>
</td>

<td align=center >
	<%IF not(request.cookies(eChsys)("KEY")="super" or request.cookies(eChsys)("KEY")="check" or request.cookies(eChsys)("key")="typemaster" or request.cookies(eChsys)("key")="bigmaster") THEN%>
		<%if rs("goodnews")<>1 then%>�Ƽ�<%else%>�Ѽ�<%end if%>
	<%else%>
		<a href='Newsgood.asp?Newsid=<%=rs("newsid")%>'><%if rs("goodnews")<>1 then%>�Ƽ�<%else%>�Ѽ�<%end if%></a>
	<%end if%>
</td>

<td align=center >
	<%IF not(request.cookies(eChsys)("KEY")="super" or request.cookies(eChsys)("KEY")="check" or request.cookies(eChsys)("key")="typemaster" or request.cookies(eChsys)("key")="bigmaster") THEN%>
		<%if rs("istop")<>1 then%>�̶�<%else%>�ѹ�<%end if%>
	<%else%>
		<a href='NewsGu.asp?Newsid=<%=rs("newsid")%>'><%if rs("istop")<>1 then%>�̶�<%else%>�ѹ�<%end if%></a>
	<%end if%>
</td>
<td align=center ><a href='NewsDel.asp?Newsid=<%=rs("newsid")%>'>ɾ��</a></td>
</tr>
<%
end if
rs.MoveNext
end if
next
%>
<tr>
    <td colspan=8  align=center height=25>�� <%=total%> ������ǰ�� <%=Mypage%>/<%=Maxpages%> 
            ҳ��ÿҳ <%=MyPageSize%> �� 
            <%
url="listnews.asp?E_SmallClassID=" & E_SmallClassID & "&"
							
PageNextSize=int((MyPage-1)/PageShowSize)+1
Pagetpage=int((total-1)/rs.PageSize)+1

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
%>
</td>
  </tr>
</table>
</div>
<%else%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - ��������</title>
</head>
<body topmargin="0">
<div align="center">
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFDFBF" width="100%" id="AutoNumber1">
<tr>
    <td width="100%" bgcolor="<%=m_top%>" height="30" align="left">
<a href=E_TypeManage.asp>ȫ������</a><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src=images/menu1.gif><b><a href="E_BigClassManage.asp?E_typeid=<%=E_typeid%>"><font color=red><%=E_typename%></font></a></b>&nbsp;&nbsp;&nbsp;&nbsp;����������
<%
Set rs1 = Server.CreateObject("ADODB.Recordset")
sql1 ="SELECT  * From "& db_EC_Type_Table &" where E_typeid<>"&E_typeid&" order by E_typeorder"
RS1.open sql1,Conn,3,3
do while not rs1.EOF
master1=rs1("typemaster")
if master1<>"" then
tmaster1=split(master1,"|")
 for i=0 to ubound(tmaster1)
	if request.cookies(eChsys)("username")=trim(tmaster1(i)) then
	typemaster1=true
	exit for
	else
	typemaster1=false
	end if
 next
end if
if typemaster1=true or request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="check" or request.cookies(eChsys)("key")="selfreg" then
%><a href=E_BigClassManage.asp?E_typeid=<%=rs1("E_typeid")%>><%=rs1("E_typename")%></a>
<%end if
rs1.movenext
loop
rs1.close
set rs1=nothing%>
      <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src=images/menu1.gif><b><a href="E_SmallClassManage.asp?E_BigClassID=<%=E_BigClassID%>"><font color=red><%=E_BigClassName%></font></a></b>
      &nbsp;&nbsp;&nbsp;&nbsp;�������ࣺ
<%
Set rs1 = Server.CreateObject("ADODB.Recordset")
sql1 ="SELECT  * From "& db_EC_BigClass_Table &" where E_typeid="&E_typeid&" and E_BigClassID<>"&E_BigClassID&" order by E_bigclassorder"
RS1.open sql1,Conn,3,3
do while not rs1.EOF
if (bigclassmaster=true and request.cookies(eChsys)("key")="bigmaster") or request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="check" or (request.cookies(eChsys)("key")="typemaster" and typemaster4=true) or request.cookies(eChsys)("key")="selfreg" then
%><a href=E_SmallClassManage.asp?E_BigClassID=<%=rs1("E_BigClassID")%>><%=rs1("E_BigClassName")%></a>
<%end if
rs1.movenext
loop
rs1.close
set rs1=nothing%>
      <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src=images/menu1.gif><b><a href="listnews.asp?E_SmallClassID=<%=E_SmallClassID%>"><font color=red><%=E_smallclassname%></font></a></b>
      &nbsp;&nbsp;&nbsp;&nbsp;����С�ࣺ
<%if request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="check" or request.cookies(eChsys)("key")="typemaster" or request.cookies(eChsys)("key")="bigmaster" or request.cookies(eChsys)("key")="selfreg" then
Set rs1 = Server.CreateObject("ADODB.Recordset")
sql1 ="SELECT  * From "& db_EC_SmallClass_Table &" where E_BigClassID="&E_BigClassID&" and E_SmallClassID<>"&E_SmallClassID&" order by smallclassorder"
RS1.open sql1,Conn,3,3
do while not rs1.EOF
%><a href=listnews.asp?E_SmallClassID=<%=rs1("E_SmallClassID")%>><%=rs1("E_smallclassname")%></a>
<%rs1.movenext
loop
rs1.close
set rs1=nothing
else
Set rs1 = Server.CreateObject("ADODB.Recordset")
sql1 ="SELECT * From "& db_EC_SmallClass_Table &" where E_SmallClassID<>"&E_SmallClassID&" order by SmallClassorder"
RS1.open sql1,Conn,3,3
do while not rs1.EOF
master2=rs1("smallclassma")
if master2<>"" then
smallmaster=split(master2,"|")
 for i=0 to ubound(smallmaster)
	if request.cookies(eChsys)("username")=trim(smallmaster(i)) then
	smallclassmaster=true
	exit for
	else
	smallclassmaster=false
	end if
 next
end if
if smallclassmaster=true then
%><a href=listnews.asp?E_SmallClassID=<%=rs1("E_SmallClassID")%>><%=rs1("E_smallclassname")%></a>
<%end if
rs1.movenext
loop
rs1.close
set rs1=nothing
end if
%>
</td>
</tr>
<tr>
    <td  align=center height=25>��С���»�δ������<input type=button value="����������" class=button onClick="javascript:window.open('newsadd1.asp?E_SmallClassID=<%=E_SmallClassID%>','_self','')"  ></td>
</tr>
</table>
</body>
</html>
<%end if
rs.close
set rs=nothing
conn.close
set conn=nothing
end if
else
response.write "<script>alert('�Բ�������Ȩ������');history.go(-1);</Script>"
Response.End 
end if
%>
<!--#include file="Admin_Bottom.asp"-->