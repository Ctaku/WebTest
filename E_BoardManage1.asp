<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--��#include file="char.inc"-->
<!--#include file="ChkUser.asp" -->

<%
PageShowSize = 10            'ÿҳ��ʾ���ٸ�ҳ
MyPageSize   = 10           'ÿҳ��ʾ����������
	
If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
MyPage=1
Else
MyPage=Int(Abs(Request("page")))
End if

%>
<%

	
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_EC_Board_Table &" order by id desc"
	rs.open sql,conn,1,1
	%>
<head><meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - �������</title>
</head>
<body topmargin="0">

<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFDFBF" width="100%" id="AutoNumber1">
<form method="POST" action="E_BoardSet.asp">
<tr class="TDtop">
	<td width="100%" height="25" colspan=5 align=center>�� <strong>�������</strong> ��</td>
</tr>
<tr class="TDtop1" height=20>
	<td width="10%" align="center">ѡ��</td>
	<td width="10%" align="center">ID</td>
	<td width="60%" align="center">����</td>
	<td width="10%" align="center">�޸�</td>
	<td width="10%" align="center">ɾ��</td>
</tr>
	<%
		
	If Not rs.eof then
rs.PageSize     = MyPageSize
MaxPages         = rs.PageCount
rs.absolutepage = MyPage
total            = rs.RecordCount
i = 0
do until rs.Eof or i = rs.PageSize
		%>
<tr>
	<td width="10%" align="center">
		<input type="radio" value=<%=rs("ID")%><%if rs("inuse")=1 then%> checked<%end if%> name="inuse">
	</td>
	<td width="10%" align="center"><%=rs("ID")%></td>
	<td width="60%"><%=CheckStr(rs("Title"))%></td>
	<td width="10%" align="center">
		<a href='E_BoardModify.asp?id=<%=rs("ID")%>' style="font-size: 9pt;  color: #000000; background-color: #ffffff; solid #FEEBE4" onMouseOver ="this.style.backgroundColor='#FEEBE4'" onMouseOut ="this.style.backgroundColor='#ffffff'">�޸�</a>
	</td>
	<td width="10%" align="center">
		<a href='E_BoardDel.asp?id=<%=rs("ID")%>' style="font-size: 9pt;  color: #000000; background-color: #ffffff; solid #FEEBE4" onMouseOver ="this.style.backgroundColor='#FEEBE4'" onMouseOut ="this.style.backgroundColor='#ffffff'">ɾ��</a>
	</td>
</tr>
<%
rs.MoveNext
i = i + 1
loop
%>
   <tr>
     <td width="100%" align=center colspan="5">
        �� <%=total%> ������ǰ�� <%=Mypage%>/<%=Maxpages%> ҳ��ÿҳ <%=MyPageSize%> ��
 <%
url="E_BoardManage.asp?"				
PageNextSize=int((MyPage-1)/PageShowSize)+1
Pagetpage=int((total-1)/rs2.PageSize)+1

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

else
Response.write "<tr><td valign=top height=80 colspan=5 align=center>&nbsp;���޹���</td></tr>"				
End If
rs.close

%>
               </td>
             </tr>

<tr>
	<td colspan=5 align=center>
		<input type="submit" value="��Ϊ�����Ƽ�����" name="submit" > &nbsp;
		<input onClick="javascript:window.open('E_BoardAdd.asp','_self','')" type="button" value="����¹���" name="button" >
</td>
</tr>
</form>
</table>
</div>
</body>
</html>
<!--#include file="Admin_Bottom.asp"-->
	