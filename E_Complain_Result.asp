<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<%

keyword=trim(checkstr(Request("keyword")))

PageShowSize = 10	'ÿҳ��ʾ���ٸ�ҳ
MyPageSize   = 10	'ÿҳ��ʾ������
	
If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
	MyPage=1
Else
	MyPage=Int(Abs(Request("page")))
End if

if keyword="" or  keyword="���������ؼ���" then
	%>
	<script language=javascript>
		alert("�������ѯ�ؼ��֣�")
		history.back()
	</script>
	<body onload=javascript:window.close()></body> 
	<%
	Response.End
end if

set rs=server.CreateObject("ADODB.RecordSet")
sql1="select * from "& db_EC_Complain_Table &" where topic like '%"&keyword&"%'  order by ComplainID DESC"
rs.Open sql1,conn,1,1

%>
<html>
<head>
<title>Ͷ�߾ٱ��������__<%=eChSys%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="news.css" rel=stylesheet type=text/css>
</head>
<body bgcolor="#FFFFFF" text="#000000" topmargin="0">
<!--#include file="other_Top.asp"-->

<table width="780" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
        <tr valign="top">
          <td height="163"><table width="780" height="194" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td width="190" valign="top" background="images\left_bg.gif"><table width="169" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="169" height="43" background="images\title_bg.gif"><div align="center" class="yellow_title">�ֳ�����</div></td>
                </tr>
              </table>
                <table width="169" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td><img src="Images/btn_box_top.gif" width="169" height="20"></td>
                  </tr>
                  <tr>
                    <td width="169" height="121" valign="top" background="images\bg_content_169px.gif"><div align="center" class="yellow_title">
                      <table width="98%" height="127" border="0" cellpadding="0" cellspacing="0">
                        <tr>
                          <td width="31%" height="33" align="center" valign="middle"><img src="Images/Sicon_QYTS.gif" width="22" height="22"></td>
                          <td width="69%" align="left" valign="middle" ><a href="E_LeadMail.asp" class="top1">�ֳ�����</a></td>
                        </tr>
                        <tr>
                          <td height="32" align="center" valign="middle"><img src="Images/Sicon_XZBSZX.gif" width="22" height="22"></td>
                          <td align="left" valign="middle"><a href="E_GuestBook.Asp" class="top1">��������</a></td>
                        </tr>
                        <tr>
                          <td height="31" align="center" valign="middle"><img src="Images/Sicon_XZTS.gif" width="22" height="22"></td>
                          <td align="left" valign="middle"><a href="Complain.asp" class="top1">Ͷ�߾ٱ�</a></td>
                        </tr>
                        <tr>
                          <td align="center" valign="middle"><img src="Images/Sicon_MYZJ.gif" width="22" height="22"></td>
                          <td align="left" valign="middle"><a href="E_Opinion.asp" class="top1">�����ײ�</a></td>
                        </tr>
                      </table>
                    </div></td>
                  </tr>
                  <tr>
                    <td><img src="Images/btn_box_bottom.gif" width="169" height="22"></td>
                  </tr>
                </table></td>
              <td width="588" valign="top"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="9%" height="53" valign="middle" background="Images/location_bg.gif" align="center" >&nbsp;<img src="Images/point_location.gif" width="24" height="24"></td>
                  <td width="91%" valign="middle" background="Images/location_bg.gif" class="daohang">&nbsp;����λ�ã�&nbsp;<a class=daohang href="./" >��վ��ҳ</a>&gt;&gt;<a href="Complain.asp" class="daohang">Ͷ�߾ٱ�</a>&gt;&gt;Ͷ�߾ٱ�����</td>
                </tr>
              </table>
                <table width="560" height="300" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td height="42" background="images\topic_bg.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="yellow_title">Ͷ�߾ٱ��������</span></td>
                  </tr>
                  <tr>
                    <td height="300" valign="top" background="images\bg_content_560px.gif">
                      <br>
                      <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#E6D9A4" id="AutoNumber1" style="border-collapse: collapse">	
                        <tr class="TDtop1" align="center">
                          <td width="7%" height="25" bgcolor="#FCFDEB"><span class="top1">������</span></td>
                          <td width="43%" bgcolor="#FCFDEB" class="top1">Ͷ�߱���</td>
                          <td width="12%" bgcolor="#FCFDEB" class="top1">Ͷ������</td>
                        </tr>
                        <%
			
if not rs.EOF then
rs.PageSize     = MyPageSize
MaxPages         = rs.PageCount
rs.absolutepage = MyPage
total            = rs.RecordCount
dim i
i = 1
do until rs.Eof or i = rs.PageSize
%> 
				<% 	topic=trim(rs("topic"))
				DisplayTopic=mid(Topic,1,28)
				%>
                          <tr bgcolor=ffffff>
                            <td height="23" align=center bgcolor="#FCFDEB" >
					        <%if trim(rs("ReplyContent"))<>"" then %>
					        <font color=red >�������</font>
					        <%else%>
					        <font color=red >������</font>
					        <%end if%>                            </td>
<td align="left" bgcolor="#FCFDEB">
��<a  class=middle href='E_ReadComplain.asp?ComplainID=<%=rs("ComplainID")%>' target="_black"><%=htmlencode(DisplayTopic)%>...</a></td>
                            <td align=center bgcolor="#FCFDEB"><%=year(rs("UpdateTime")) %>-<%=Month(rs("UpdateTime")) %>-<%=Day(rs("UpdateTime")) %></td>
                          </tr>
					    <%
		rs.MoveNext
		i = i + 1
		loop
		%>
		<tr> 
	<td colspan="5" width="100%" align=center height=25 bgcolor="#FCFDEB">�������� <%=total%> ������ǰ�� <%=Mypage%>/<%=Maxpages%> ҳ��ÿҳ <%=MyPageSize%> ��
	<%
	url="E_Complain_Result.asp?keyword=" & keyword
PageNextSize=int((MyPage-1)/PageShowSize)+1
Pagetpage=int((total-1)/rs.PageSize)+1

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
Response.write "<tr><td align=center bgcolor=#FCFDEB colspan=3 height=100><font color=red>�Բ���û���ҵ���Ҫ������Ͷ�߾ٱ���</font></td></tr>"
			
End If
rs.close
set rs=nothing
%>
</td>
</tr>			
</table>

                      <br>
				    </td>
                  </tr>
                  <tr>
                    <td height="13"><img src="Images/shadow_690px.gif" width="560" height="17"></td>
                  </tr>
                </table></td>
            </tr>
          </table>
          </td>
        </tr>
</table>
<!--#include file="other_bottom.asp"-->
