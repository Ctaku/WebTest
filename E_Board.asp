<%@ Language=VBScript%>
<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="function.asp"-->
<%
PageShowSize = 10            'ÿҳ��ʾ���ٸ�ҳ
MyPageSize   = 30           'ÿҳ��ʾ����������
	
If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
MyPage=1
Else
MyPage=Int(Abs(Request("page")))
End if

set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from "& db_EC_Type_Table &" order by E_typeorder"
rs.Open rs.Source,conn,1,1

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ȫ������__<%=eChSys%></title>
<LINK href="news.css" rel="stylesheet">
</head>
<body>
<!--#include file="other_Top.asp"-->
<table width="780" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
<tr>
     <td width="780" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
       <tr>
         <td height="25" background="Images/dh_bg.gif" class="white_link">&nbsp;&nbsp;&nbsp;&nbsp;<img src="Images/arrow_dh_1.gif" width="7" height="7">&nbsp;&nbsp;<span class="daohang">��Ŀ����</span><b>&nbsp;&nbsp; </b><span class="daohang">��ǰλ�ã�<a class="daohang" href="./" >��վ��ҳ</a>&gt;&gt;ȫ������</span></td>
       </tr>
       <tr>
         <td><table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
           <tr>
             <td height="27" class="yellow_title"> ���¹��� </td>
           </tr>
           <tr>
             <td><img src="Images/list_line.gif" width="760" height="10"></td>
           </tr>
           <tr>
             <td><%
set rs2=server.CreateObject("ADODB.RecordSet")
rs2.Source="select * from "& db_EC_Board_Table &" order by ID DESC "
rs2.Open rs2.Source,conn,3,1

If Not rs2.eof then
rs2.PageSize     = MyPageSize
MaxPages         = rs2.PageCount
rs2.absolutepage = MyPage
total            = rs2.RecordCount
i = 0
do until rs2.Eof or i = rs2.PageSize
%>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
  <tr>
     <td width="100%" height="20" colspan="2" >
      <img src="Images/list_dot.gif" width="10" height="10">&nbsp;<A HREF="E_Board_News.asp?ID=<%=rs2("ID")%>" class="middle"><%=rs2("title")%></A>&nbsp;<%=rs2("dateandtime")%>     </td>
  </tr>
</table>
<%
rs2.MoveNext
i = i + 1
loop
%>	</td>
  </tr>
</table>
           <br>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
   <tr>
     <td width="100%" align=center colspan="2">
        �� <%=total%> ������ǰ�� <%=Mypage%>/<%=Maxpages%> ҳ��ÿҳ <%=MyPageSize%> ��
 <%
url="E_Board.asp?"				
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
Response.write "<tr><td valign=top height=80>&nbsp;������Ϣ</td></tr>"				
End If
rs2.close

%>               </td>
             </tr>
         </table></td>
       </tr>
     </table></td>
  </tr>
  
  <tr><td colspan="2" height="5"></td></tr>
</table> 
<!--#include file="other_bottom.asp"-->
</body>
</html>
