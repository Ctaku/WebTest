<%@ Language=VBScript%>
<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="function.asp"-->
<%
username1=checkstr(trim(Request("user")))

PageShowSize = 10            'ÿҳ��ʾ���ٸ�ҳ
MyPageSize   = 10           'ÿҳ��ʾ������
	
If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
	MyPage=1
Else
	MyPage=Int(Abs(Request("page")))
End if

if username1="" then
	Response.Write "<script>alert('δָ������');history.back()</script>"
	Response.End
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
set rs=nothing
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=username1%>�Ĳ�����־__<%=eChSys%></title>
<LINK href=news.css rel=stylesheet>
</head>
<BODY ONLOAD="newCalendar()" OnUnload="window.returnValue = document.all.ret.value;" topmargin="0"> 
<!--#include file="other_Top.asp"-->
<table width="780" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
<tr valign="top"> 
	<td width="210" bgcolor="F1F7FC">
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
		<tr>
			<td width="100%" height="50" valign="middle" background="Images/type_dh.gif" class="white_link"> 
		      <div align="center" class="yellow_title"><img src="Images/arrow_dh_1.gif" width="7" height="7">&nbsp;&nbsp;��������</div></td>
		</tr>
<%
set rs2=server.CreateObject("ADODB.RecordSet")
rs2.Source="select * from "& db_User_Table &" where "& db_User_Name &"='"&username1&"'"
rs2.Open rs2.Source,ConnUser,1,1
if not rs2.eof then
	%>
		<tr> 
			<td width="100%" height="24" align=center valign="middle"><br>
	<%if rs2(db_User_Face)<>"" then%>
		<%if UserTableType = "EChuang" then%>
			<%if rs2(db_User_Face)<>"images/nopic.gif" then %>
				<img src="<%=rs2(db_User_Face)%>" border="0" width="120">
			<%else%>
				<%if rs2(db_User_sex)="����" then %>
				<img src="images/nopic.gif" border="0" >
				<%else%> 
					<%if  rs2(db_User_sex)="����" then %>
				   <img src="images/Image9.gif" border="0" >
				<%else%>
					<%if  rs2(db_User_sex)="Ůʿ" then %>
				   <img src="images/Image18.gif" border="0" >
					<%end if%>
				<%end if%>
			<%end if%>
		<%end if%>
	<%else%>
		<%if UserTableType="Dvbbs" then%>
			<img src="<%=BbsPath%><%=rs2(db_User_Face)%>" border="0" width="<%=rs2(db_User_FaceWidth)%>" height="<%=rs2(db_User_FaceHeight)%>"> 
			<%''��ʾ�û�ͷ�񣬼�'bbs/'ǰ׺·��,ʹͼ��ϵͳֱ����ʾ������̳ͷ��%><br>
			<a href="<%=BbsPath%>dispuser.asp?name=<%=username1%>" target="_blank"><font color=blue>��̳�еĸ�������</font></a><br><br>
		<%end if%>
	<%end if%>
<%else%>
	<% if rs2(db_User_sex)="����" then %>
				<img src="images/nopic.gif" border="0" >
	<%else%> 
		<%if  rs2(db_User_sex)="����" then %>
				<img src="images/Image9.gif" border="0" >
		<%else%>
			<%if  rs2(db_User_sex)="Ůʿ" then %>
				<img src="images/Image18.gif" border="0" >
			<%else%> 
			<%end if%>
		<%end if%>
	<%end if%> 
<%end if%>			
		<tr> 
			<td width="100%" height="24" valign="middle" style="WORD-WRAP: break-word">
&nbsp;&nbsp;��&nbsp;�ǳ�:<%=username1%><br>
&nbsp;&nbsp;��&nbsp;��ʵ����:<%=rs2("fullname")%><br>
<%if db_Birthday_Select = "EChuang" then%>
&nbsp;&nbsp;��&nbsp;��λ:<%=rs2("E_DepName")%> <br>
&nbsp;&nbsp;��&nbsp;����:<%=rs2("E_DepType")%><br>
&nbsp;&nbsp;��&nbsp;����:<%=rs2(db_User_Birthyear)%>-<%=rs2(db_User_Birthmonth)%>-<%=rs2(db_User_Birthday)%><br>
<%else%>
	<%if db_Birthday_Select = "Text" then%>
&nbsp;&nbsp;��&nbsp;����:<%=year(rs2(db_User_birthday))%>-<%=month(rs2(db_User_birthday))%>-<%=day(rs2(db_User_birthday))%><br>
	<%end if%>
<%end if%>
<%if db_Sex_Select = "EChuang" then%>
&nbsp;&nbsp;��&nbsp;�Ա�:<%=rs2(db_User_sex)%><br>
<%else%>
	<%if db_Sex_Select = "Number" then%>
&nbsp;&nbsp;��&nbsp;�Ա�:<%if rs2(db_User_Sex)="0" then%>Ů<%else%>��<%end if%><br>
	<%end if%>
<%end if%>
&nbsp;&nbsp;��&nbsp;EMAIL:<A HREF="mailto:<%=rs2(db_User_Email)%>"><%=CutStr(rs2(db_User_Email),8)%></A><br>
&nbsp;&nbsp;��&nbsp;ע��ʱ��:<%=year(rs2(db_User_RegDate))%>-<%=month(rs2(db_User_RegDate))%>-<%=day(rs2(db_User_RegDate))%><br><br></td>
		  </tr>
		<tr> 
		  <td height="50" background="IMAGES/type_dh.gif" class="yellow_title"><div align="center"><img src="Images/arrow_dh_1.gif" width="7" height="7">&nbsp;&nbsp;��������</div></td>
		</tr>
		<tr>
			<td>
<!--������ʼ-->
<br>
<STYLE TYPE="text/css"> 
.today {font-weight:bold;BACKGROUND: #6699cc} 
.satday{color:green} 
.sunday{color:red} 
.days {font-weight:bold} 
</STYLE>
<script language="JavaScript"> 
//�����·�,�������ʾӢ���·ݣ��޸������ע�� 
/*var months = new Array("January?, "February?, "March", 
"April", "May", "June", "July", "August", "September", 
"October", "November", "December");*/ 
var months = new Array("һ��", "����", "����", 
"����", "����", "����", "����", "����", "����", 
"ʮ��", "ʮһ��", "ʮ����"); 
var daysInMonth = new Array(31, 28, 31, 30, 31, 30, 31, 31, 
30, 31, 30, 31); 
//������ �������ʾ Ӣ�ĵģ��޸������ע�� 
/*var days = new Array("Sunday", "Monday", "Tuesday", 
"Wednesday", "Thursday", "Friday", "Saturday");*/ 
var days = new Array("��","һ", "��", "��", 
"��", "��", "��"); 
function getDays(month, year) { 
//�������δ������жϵ�ǰ�Ƿ�������� 
if (1 == month) 
return ((0 == year % 4) && (0 != (year % 100))) || 
(0 == year % 400) ? 29 : 28; 
else 
return daysInMonth[month]; 
} 

function getToday() { 
//�õ��������,��,�� 
this.now = new Date(); 
this.year = this.now.getFullYear(); 
this.month = this.now.getMonth(); 
this.day = this.now.getDate(); 
} 

today = new getToday(); 

function newCalendar() { 

today = new getToday(); 
var parseYear = parseInt(document.all.year 
[document.all.year.selectedIndex].text); 
var newCal = new Date(parseYear, 
document.all.month.selectedIndex, 1); 
var day = -1; 
var startDay = newCal.getDay(); 
var daily = 0; 
if ((today.year == newCal.getFullYear()) &&(today.month == newCal.getMonth())) 
day = today.day; 
var tableCal = document.all.calendar.tBodies.dayList; 
var intDaysInMonth =getDays(newCal.getMonth(), newCal.getFullYear()); 
for (var intWeek = 0; intWeek < tableCal.rows.length;intWeek++) 
for (var intDay = 0;intDay < tableCal.rows[intWeek].cells.length;intDay++) 
{ 
var cell = tableCal.rows[intWeek].cells[intDay]; 
if ((intDay == startDay) && (0 == daily)) 
daily = 1; 
if(day==daily) 
//���죬���ý����Class 
cell.className = "today"; 
else if(intDay==6) 
//���� 
cell.className = "sunday"; 
else if (intDay==0) 
//���� 
cell.className ="satday"; 
else 
//ƽ�� 
cell.className="normal"; 

if ((daily > 0) && (daily <= intDaysInMonth)) 
{ 
cell.innerText = daily; 
daily++; 
} 
else 
cell.innerText = ""; 
} 
} 

function getDate() { 
var sDate; 
//��δ��봦������������ 
if ("TD" == event.srcElement.tagName) 
if ("" != event.srcElement.innerText) 
{ 
sDate = document.all.year.value + "��" + document.all.month.value + "��" + event.srcElement.innerText + "��"; 
//alert(sDate); 
window.open ("BlogResult.asp?keyword="+document.all.year.value+"-"+document.all.month.value+"-"+event.srcElement.innerText+"&action=UpdateTime") 
} 
} 
			</script>
			<input type="hidden" name="ret">
			<table id="calendar" border="0" cellpadding="3" cellspacing="0"  bgcolor="F1F7FC">
			<thead>
			<tr>
				<TD COLSPAN=7 ALIGN=CENTER>
<SELECT ID="month" ONCHANGE="newCalendar()"> 
	<SCRIPT LANGUAGE="JavaScript"> 
	for (var intLoop = 0; intLoop < months.length; 
	intLoop++) 
	document.write("<OPTION VALUE= " + (intLoop + 1) + " " + 
	(today.month == intLoop ? 
	"Selected" : "") + ">" + 
	months[intLoop]); 
	</SCRIPT> 
</SELECT>

<SELECT ID="year" ONCHANGE="newCalendar()"> 
	<SCRIPT LANGUAGE="JavaScript"> 
	for (var intLoop = today.year-50; intLoop < (today.year + 10); 
	intLoop++) 
	document.write("<OPTION VALUE= " + intLoop + " " + 
	(today.year == intLoop ? 
	"Selected" : "") + ">" + 
	intLoop); 
	</script>
</select>
<br><br>					</td>
				</tr>
				<tr class="days"> 
<script language="JavaScript"> 
document.write("<TD class=satday>" + days[0] + "</TD>"); 
for (var intLoop = 1; intLoop < days.length-1; 
intLoop++) 
document.write("<TD>" + days[intLoop] + "</TD>"); 
document.write("<TD class=sunday>" + days[intLoop] + "</TD>"); 
</script>
				</tr>
			  </thead>
				<tbody border=1 cellspacing="0" cellpadding="0" id="dayList"align=CENTER onClick="getDate()">
<script language="JavaScript"> 
for (var intWeeks = 0; intWeeks < 6; intWeeks++) { 
document.write("<TR style='cursor:hand'>"); 
for (var intDays = 0; intDays < days.length; 
intDays++) 
document.write("<TD></TD>"); 
document.write("</TR>"); 
} 
</script>
				</tbody>
			  </table>
				<br>
<!--��������--></td>
		</tr>
		</table>
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
		<tr> 
		  <td height="50" background="IMAGES/type_dh.gif" class="yellow_title"><div align="center"><img src="Images/arrow_dh_1.gif" width="7" height="7">&nbsp;&nbsp;��վ����</div></td>
		</tr>
		<tr> 
			<td>
				<div align="center">
<!--�࿪ʼ-->
				<TABLE width="90%">
				<TR>
			  	<TD>
<%
dim menuid
dim menuname
'dim menucontent
dim menuview
set rs11=server.CreateObject("ADODB.RecordSet")
rs11.Source="select * from "& db_EC_BigClass_Table &" order by E_bigclassorder"
rs11.Open rs11.Source,conn,1,1
i=1
Dim ArraymenuID(1000),ArraymenuName(1000)
while not rs11.EOF
	RecordCount=rs11.RecordCount
	menuid=rs11("E_BigClassID")
	menuname=rs11("E_BigClassName")
	'menucontent=rs11("bigclasscontent")
	menuview=rs11("E_BigClassView")
	TdString="<br><font color='#000000'>��</font>&nbsp;<a href='blogResult.asp?keyword="& username1 &"&action=editor&E_BigClassID="&menuid&"' class=middle >"& menuName &"</a><br>" 
	Response.Write TdString
	ArraymenuID(i)=menuID
	ArraymenuName(i)=menuName
	i=i+1
	rs11.MoveNext
wend
rs11.close
set rs11=nothing
%><br>				  </TD>
				</TR>
				</TABLE>
<!--�����-->
				</div>			</td>
		</tr>
	  </table>
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
		<tr> 
		  <td height="50" background="IMAGES/type_dh.gif" class="yellow_title"><div align="center"><img src="Images/arrow_dh_1.gif" width="7" height="7">&nbsp;&nbsp;���ͽ���</div></td>
		</tr>
		<tr> 
			<td>
				<div align="center">
				<TABLE width="95%" align="center">
			  <TR>
					<TD>
<br>
��&nbsp;blog����:<%=CutStr(nohtml(rs2("content")),10)%><br>
��&nbsp;��־����:<%=rs2("number")%><br>
��&nbsp;��½����:<%=rs2(db_User_LoginTimes)%><br>
��&nbsp;����ʱ��:<%=year(rs2(db_User_RegDate))%>-<%=month(rs2(db_User_RegDate))%>-<%=day(rs2(db_User_RegDate))%>
<br><br>					</TD>
				</TR>
				</TABLE>
				</div>			</td>
		</tr>
	  </table>
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
		<tr> 
		  <td height="50" background="IMAGES/type_dh.gif" class="yellow_title"><div align="center"><img src="Images/arrow_dh_1.gif" width="7" height="7">&nbsp;&nbsp;��������</div></td>
		</tr>
		<tr> 
			<td>
				<div align="center">
			  <TABLE width="90%">
			  <TR>
			  	<TD><!--#include file="ReviewIndex.asp"--></TD>
			  </TR>
			  </TABLE>
				</div>	    </td>
		</tr>
	  </table>
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
		<tr> 
		  <td height="50" background="IMAGES/type_dh.gif" class="yellow_title"><div align="center"><img src="Images/arrow_dh_1.gif" width="7" height="7">&nbsp;&nbsp;��������</div></td>
		</tr>
		<tr> 
			<td>
				<TABLE width="90%">
				<TR>
					<TD><!--#include file="BlogtopUser.asp" --></TD>
				</TR>
				</TABLE>
			</td>
		</tr>
	  </table>
	</TD>
	<td  width="570">
	<table  width="570" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="25" align="center" background="Images/dh_bg.gif" class="daohang" ><span class="daohang"><A HREF="index.asp" class="daohang">[��վ��ҳ]</A>&nbsp;</span>&nbsp;<%=CutStr(nohtml(rs2("content")),40)%>&nbsp;&nbsp;<A HREF="Bloguser.asp?user=<%=username1%>" class="daohang">[������ҳ]</A> </td>
      </tr>
      <%rs2.close
		set rs2=nothing%>
                    <%
			set rs=server.CreateObject("ADODB.RecordSet")
			if session("key")="" then
				rs.Source="select * from "& db_EC_News_Table &" where editor='" & username1 & "' and checkked=1 order by NewsID DESC"
			end if
			if session("key")="super" or session("key")="typemaster" or session("key")="bigmaster" or session("key")="smallmaster" or session("key")="check" then
				rs.Source="select * from "& db_EC_News_Table &" where editor='" & username1 & "' and checkked=1 order by NewsID DESC"
			end if
			if session("key")="selfreg" then
				if session("reglevel")=3 then
					rs.Source="select * from "& db_EC_News_Table &" where editor='" & username1 & "' and checkked=1 order by NewsID DESC"
				end if
				if session("reglevel")=2 then
					rs.Source="select * from "& db_EC_News_Table &" where editor='" & username1 & "' and checkked=1 order by NewsID DESC"
				end if
				if session("reglevel")=1 then
					rs.Source="select * from "& db_EC_News_Table &" where editor='" & username1 & "' and checkked=1 order by NewsID DESC"
				end if
			end if

			rs.Open rs.Source,conn,1,1
			rs.PageSize=20
			rs.CacheSize = RS.PageSize
			for i=1 to rs.PageSize *(page-1)
				if not rs.EOF then
					rs.MoveNext
				end if
			next

			Response.Write "<tr><td width=100% height=28 align=center>&nbsp;"
			if rs.EOF then
				Response.Write "<font color=red>�û�Ա���ͻ�û�����£�</font>"
			else
				rs.PageSize     = MyPageSize
				MaxPages         = rs.PageCount
				rs.absolutepage = MyPage
				total            = rs.RecordCount
				Response.Write "�û�Ա���͹���" & total & "ƪ���£���ǰ��"& myPage &"/"& MaxPages &"ҳ��ÿҳ"& rs.PageSize &"��"
			end if
			Response.Write "</td></tr>"

			If Not rs.eof then
				i = 0
				do until rs.Eof or i = rs.PageSize
				if rs("picname")<>"" then
					img="<font color=red>[ͼ]</font>"
				else
					img=""
				end if

				title=htmlencode4(trim(rs("title")))
				dim E_BigClassID
				E_BigClassID=rs("E_BigClassID")
				set rs11=server.CreateObject("ADODB.RecordSet")
				rs11.Source="select * from "& db_EC_BigClass_Table &" where E_BigClassID="&E_BigClassID&""

				If  E_BigClassID <>"" Then
				
				rs11.Open rs11.Source,conn,1,1
				E_typeid=rs11("E_typeid")
				E_BigClassName=rs11("E_BigClassName")
				rs11.Close
	
				fileExt=lcase(getFileExtName(rs("picname")))
				Content=htmlencode4(rs("Content"))
				content=replace(content,"[---��ҳ---]","")
				
				End if

				%>
                    <tr>
                      <td width="570" height="20" bgcolor="F1F7FC">&nbsp;&nbsp; <img src="images/blog.gif" border="0" align="absmiddle">&nbsp;
                          <% If  E_BigClassID <>"" Then %>
                        [<%=E_BigClassName%>]
                        <% else %>
                        [�޴�����Ŀ]
                        <% End if %>
                      &nbsp;<a class=middle href="E_ReadNews.asp?NewsID=<%=rs("NewsID")%>" target=_blank title="<%=title%>"><%=CutStr(title,50)%></a> </td>
                    </tr>
                    <tr>
                      <td align="right"><font class=noline><%=username1%>������(<%=rs("UpdateTime")%>)&nbsp;&nbsp;�����[<font color="#ff0000"><%=rs("click")%></font>]</font> </td>
                    </tr>
                    <tr>
                      <td><p><font class=middle><br>
                              <%=CutStr(nohtml(rs("Content")),300)%></font><br>
                      </p>
                          <%if rs("picname")=("") then%>
                          <%else%>
                          <%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
                          <p align="center"><img  src="<%=FileUploadPath & rs("picname")%>" onload='javascript:if(this.width>500) this.width=500' border=0 ></p>
                        <%end if%>
                          <%if fileext="swf" then%>
                          <p align="center">
                            <object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="500"  border=0 >
                              <param name=movie value="<%=FileUploadPath & rs("picname")%>">
                              <param name=quality value=high>
                              <param name='Play' value='-1'>
                              <param name='Loop' value='0'>
                              <param name='Menu' value='-1'>
                              <embed src="<%=FileUploadPath & rs("picname")%>" width="500" border=0  pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
                            </object>
                          </p>
                        <%end if%>
                          <%end if%>
                      </td>
                    </tr>
                    <tr>
                      <td align="right"> | <a class=middle href="E_ReadNews.asp?NewsID=<%=rs("NewsID")%>" target=_blank >�Ķ�ȫ��</a> | <a class=middle href="Review.asp?NewsID=<%=rs("NewsID")%>" target=_blank >�ظ�</a> | </td>
                    </tr>
                    <%
			rs.MoveNext
			i = i + 1
		loop
			%>
                    <tr>
                      <td width="100%" align=center> �� <%=Mypage%>/<%=Maxpages%>ҳ��ÿҳ <%=MyPageSize%> ��
                        <%
			url="Bloguser.asp?user=" & username1
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
			end if
			rs.close					
			%>
                      </td>
                    </tr>

    </table></td>
</tr>
</table>
<!--#include file="other_bottom.asp" -->
<%
else

Response.Write "<script>alert('�޴��û���');history.back()</script>"
Response.End

end if
%>