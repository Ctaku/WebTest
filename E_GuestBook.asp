<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="function.asp"-->
<!--#include file="char.inc"-->
<%if request.cookies(eChsys)("key")<>"" then%>
<!--#include file="ChkUser.asp"-->
<%end if
const MaxPerPage=10
dim totalPut
dim CurrentPage
dim TotalPages
dim a,j

if not isempty(request("page")) then
	currentPage=cint(request("page"))
else
	currentPage=1
end if

if request.cookies(eChsys)("key")="super" then
	aaas=1
	set urs=server.createobject("adodb.recordset")
	sql="select * from "& db_User_Table &" where "& db_User_Name &"='"&Request.cookies(eChsys)("username")&"'"
	urs.open sql,ConnUser,1,3
	if urs.bof or urs.eof then
		aaas=0
	end if
	IF Request.cookies(eChsys)("passwd")<>urs(db_User_Password) THEN
		aaas=0
	END IF
	urs.close
	set urs=nothing
end if
%>
<html>
<head> 
<title>���Ա�_<%=eChSys%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="news.css" rel=stylesheet type=text/css>
</head>
<!--#include file="other_Top.asp"-->
<body>
<table width="780" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="F1F7FC">
  <tr>
    <td width="100%" height="25" background="Images/dh_bg.gif" class="daohang" >&nbsp;&nbsp;&nbsp;<img src="Images/arrow_dh_1.gif" width="7" height="7">&nbsp;&nbsp;&nbsp;��Ŀ����&nbsp;&nbsp;<a class="daohang" href="./" >��վ��ҳ</a>&gt;&gt;<a href="E_GuestBook.asp" class="daohang">��������</a>&gt;&gt;�鿴����</td>
  </tr>
							<tr>
								<td height="48" align="center" background="IMAGES/bookBG.gif"  class="white_link">
									<a href="guestadd.asp" class=top1>��Ҫ����</a> | <a class=top1 href="javascript:window.location.reload()">ˢ��</a> </td>
							</tr>
<tr valign="top"> 
		<td>
<%set rs=server.createobject("adodb.recordset")
if not isempty(request("ReViewID")) and IsNumeric(request("ReViewID")) then
	sql="Select * From "& db_EC_Review_Table &" where ReViewID="& CheckStr(trim(request("ReViewID"))) &" Order by ReviewId DESC"
else
	sql="Select * From "& db_EC_Review_Table &" where NewsId is NULL or NewsId < 1 Order by ReviewId DESC"
end if
rs.open sql,conn,1,1
if rs.eof and rs.bof then
	if not isempty(request("ReViewID")) and IsNumeric(request("ReViewID")) then
		response.write "<br><br><br><p align=center><font color=red>�� �� �� �� ָ �� �� �� ��</font></p><br><br>"
	else
		response.write "<br><br><br><p align=center><font color=red>�� û �� �� �� �� ��</font></p><br><br>"
	end if
else
	totalPut=rs.recordcount
	if currentpage<1 then
		currentpage=1
	end if
	if (currentpage-1)*MaxPerPage>totalput then
		if (totalPut mod MaxPerPage)=0 then
			currentpage=totalPut \ MaxPerPage
		else
			currentpage=totalPut \ MaxPerPage + 1
		end if
	end if
	if currentPage=1 then
		showContent
		showpage totalput,MaxPerPage,"E_GuestBook.asp"
	else
		if (currentPage-1)*MaxPerPage<totalPut then
			rs.move (currentPage-1)*MaxPerPage
			dim bookmark
			bookmark=rs.bookmark
			showContent
			showpage totalput,MaxPerPage,"E_GuestBook.asp"
		else
			currentPage=1
			showContent
			showpage totalput,MaxPerPage,"E_GuestBook.asp"
		end if
	end if
	rs.close
end if
set rs=nothing
conn.close
set conn=nothing

sub showContent
do while not rs.eof
	author=rs("author")
	set urs=server.createobject("adodb.recordset")
	sql="select * from "& db_User_Table &" where "& db_User_Name &"='"&author&"'" 
	urs.open sql,ConnUser,1,3
	if not urs.bof or not urs.eof then
		if UserTableType = "Dvbbs" then
			photowidth=urs(db_User_FaceWidth)	''ȡע���û���̳�趨��ͼƬƬ���
			photoheight=urs(db_User_FaceHeight)	''ȡע���û���̳�趨��ͼƬ�߶� 
		end if
		user=true			''����ע���û�
	else 
		user=notreg			''�Ƿ�ע���û�
	end if
	urs.close
	set urs=nothing%>
						<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" >
							<tr>
								<td width="20%" rowspan="2" valign="top">
									<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
										<tr>
											<td width="10" height="105"></td>
											<td height="105" valign="top"><br>
											<img src="<%if trim(rs("Face"))<>"" then%>
													<%=rs("Face")%>
												    <%else%>
													images/Image1.gif 
												    <%end if%>" boder="0" 
												<%if UserTableType = "Dvbbs" then	'ȡ�û�ͷ��,��Ϊע���û�,����ͼƬ��ʾ���߷ֱ���%>
													<%if user=true then%>
														width="<%=photowidth%>" height="<%=photoheight%>" 
													<%end if%>
												<%end if%>
												  >
												<br>
												<br>
												��ݣ�<font color="red">
												<%if rs("shengfen")<>"" then%>
													<%=rs("shengfen")%> 
												<%else%>
													����
												<%end if%>
												</font>
												<br>
												������<%=rs("author")%>
												<br> 
												<%if rs("editor")<>"" then %>
													���ԣ�<%=rs("editor")%> 
												<%end if%>
												<br> 
												<%if aaas=1 then %>
													IP��<%=rs("reviewip")%> 
												<%end if%>
											</td>
										</tr>
									</table>
								</td>
								<td valign="top">
									<table width="100%" border="0" cellspacing="0" cellpadding="0" >
										<tr>
											<td width="10"></td>
											<td> 
											<%if rs("email")<>"" then%>
												<a href="mailto:<%=rs("email")%>"><img src="images/EMAIL.GIF" border="0" title='�����ʼ�'></a>
											<%else%>
												<img src="images/EMAIL.GIF" border="0">
											<%end if%>
											<%if rs("homepage")<>"" then%>
												<a href="<%=rs("homepage")%>" target="_blank"><img src="images/homepage.gif" border="0" title='��ҳ'></a>
											<%else%>
												<img src="images/homepage.gif" border="0">
											<%end if%>
											<% if rs("oicq")<>"" then%>
												<a href="http://search.tencent.com/cgi-bin/friend/user_show_info?ln=<%=rs("oicq")%>" ><img src="images/OICQ.gif" border="0"  title='��ʾQQ��Ϣ'></a>
											<%else%>
												<img src="images/OICQ.gif" border="0">
											<%end if%>
											
											    <%if aaas=1 then %>
												<a href="GuestEdit.asp?ReviewId=<%=rs("reviewid")%>"><img src="images/EDIT.GIF" border="0" title='�༭'></a> 
												<a href="guestreply.asp?reviewid=<%=rs("reviewid")%>"><img src="images/reply.GIF" border="0" title='�ظ�'></a> 
												<a href='CheckReView.asp?ReviewID=<%=rs("ReviewID")%>&Guest=2'><%if rs("passed")=0 then%><img src="images/pass0.GIF" border="0" title='ͨ�����'><%else%><img src="images/pass1.GIF" border="0" title='ȡ�����'><%end if%></a>
												<a href="guestdel.asp?reviewid=<%=rs("reviewid")%>"><img src="images/guestbook_delete.GIF" border="0"></a> 
											<%end if%>
										
											<%if rs("newsid")<>"" then%>
												<a target="_blank" href="E_ReadNews.asp?newsid=<%=rs("newsid")%>"><img src="images/reply1.GIF" width="0" height="0" border="0"></a> 
											<%end if%>
											&nbsp;
											</td>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td valign="top">
									<table width="98%" border="0" cellspacing="0" cellpadding="0" style="table-layout:fixed; word-break:break-all;border-collapse: collapse">
										<tr>
											<td >
												<%=rs("title")%> <p><b>�������ݣ�</b><br>
											<%if rs("passed")=1 then
												if rs("NewsID")=-1 then	'-1 ��ʾΪ���Ļ�
													if aaas=1 then
														response.write rs("content")
													else
														response.write "<center><BR><font color=red>����<B>���Ļ�</B>,ֻ�й���Ա�ſ��ļ�...</font></center>"
													end if
												else
													response.write rs("content")
												end if
											else
												response.write "<center><BR><font color=red><B>����δͨ�����</B></font></center>"
												if aaas=1 then
													response.write "<font color=green>�������������˶���</font><br><br>"
													response.write rs("content")
												end if
											end if%>
												</p>
												<%if rs("reply")<>"" then %>
													<p>
													<font color="#FF0000"><b><img src="Images/m_reply.gif" width="20" height="20" align="absmiddle">&nbsp;վ���ظ���</b><br>
											  <%=rs("reply")%></font></p>
												<%end if%>
											</td>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td height="25" colspan="2" align="right" valign="top">&nbsp;&nbsp;�����ڣ�<%=rs("updatetime")%>&nbsp;&nbsp;</td>
							</tr>
							<tr>
								<td  background="Images/dh_bg.gif" height="25" colspan="2">&nbsp;</td>
						  </tr>
		  </table>
<%a=a+1
	if a>=MaxPerPage then exit do
	rs.movenext
loop
end sub%>
		</td>
  </tr>
</table>
<table width="780" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr> 
		<td align="center" background="IMAGES/bookBG.gif" height="50" class="white_link">
			<span class="yellow_title"><a href="guestadd.asp" class="yellow_title">��Ҫ����</a> | <a  class="yellow_title" href="javascript:window.location.reload()">ˢ��</a>	        </span>
	  <%function showpage(totalnumber,maxperpage,filename)
			dim n
			if totalnumber mod maxperpage=0 then
				n=totalnumber \ maxperpage
			else
				n=totalnumber \ maxperpage + 1
			end if
			response.write "<form method=post action="&filename&">"
			response.write "<p align=center>"
			if CurrentPage<2 then
				response.write "<font color='#b9b9b9' style='font-family: ����; font-size: 9pt'>��ҳ ��һҳ</font> "
			else
				response.write "<a class=class href="&filename&"?page=1>��ҳ</a> "
				response.write "<a class=class href="&filename&"?page="&CurrentPage-1&">��һҳ</a> "
			end if
			if n-currentpage<1 then
				response.write "<font color='#b9b9b9' style='font-family: ����; font-size: 9pt'>��һҳ βҳ</font>"
			else
				response.write "<a class=class href="&filename&"?page="&(CurrentPage+1)&">��һҳ</a> "
				response.write "<a class=class href="&filename&"?page="&n&">βҳ</a>"
			end if
				response.write "<font color='#000080' style='font-family: ����; font-size: 9pt'> ��<b>"&totalnumber&"</b>������ <b>"&maxperpage&"</b> ������/ҳ</font> "
				response.write "<font color='#000080' style='font-family: ����; font-size: 9pt'>ת����</font><input class=smallInput type='text' name='page' size=4 maxlength=10 value="&Currentpage&">"
				response.write "<input class=buttonface type='submit' value='Go' name='cndok'></p></form>"
			end function%>	  </td>
	</tr>
	<tr>
		<td height="19" align="center" bgcolor="#FFFFFF"></td>
	</tr>
</table>
<!--#include file="other_bottom.asp"-->
</body>
</html>