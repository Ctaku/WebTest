<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="function.asp"-->
<!--'#include file="char.inc"-->
<%
IF request.cookies(eChsys)("KEY")="" THEN

else
	usernamecookie=CheckStr(request.cookies(eChsys)("UserName"))
	passwdcookie=CheckStr(trim(Request.cookies(eChsys)("passwd")))
	KEYcookie=CheckStr(trim(request.cookies(eChsys)("KEY")))
	if usernamecookie="" or passwdcookie="" then
		response.cookies(eChsys)("UserName")=""
		response.cookies(eChsys)("KEY")=""
		response.cookies(eChsys)("purview")=""
		response.cookies(eChsys)("fullname")=""
		response.cookies(eChsys)("reglevel")=""
	else
		'�ж��û��ĺϷ���
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_User_Table &" where "& db_User_Name &"='"&usernamecookie&"'"
		rs.open sql,ConnUser,1,1
		if rs.eof and rs.bof then
			response.cookies(eChsys)("UserName")=""
			response.cookies(eChsys)("KEY")=""
			response.cookies(eChsys)("purview")=""
			response.cookies(eChsys)("fullname")=""
			response.cookies(eChsys)("reglevel")=""
		end if
		IF passwdcookie<>rs(db_User_Password) THEN
			response.cookies(eChsys)("UserName")=""
			response.cookies(eChsys)("KEY")=""
			response.cookies(eChsys)("purview")=""
			response.cookies(eChsys)("fullname")=""
			response.cookies(eChsys)("reglevel")=""
		END IF
		'�����ж��û�����ʵ�������û������Ƕ�Ӧ���ж�
		if KEYcookie<>rs("OSKEY") then
			response.cookies(eChsys)("UserName")=""
			response.cookies(eChsys)("KEY")=""
			response.cookies(eChsys)("purview")=""
			response.cookies(eChsys)("fullname")=""
			response.cookies(eChsys)("reglevel")=""
		end if
		rs.close
		set rs=nothing
	END IF
END IF



'���ļ���Ҫ���е���������
dim E_typename
NewsID=Request.QueryString("NewsID")
Page=Request.QueryString("page")

if page="" then
	page=1
	elseif not IsNumeric(page) then
		Show_Err("�Ƿ�������<br><br><a href='javascript:history.back()'>����</a>")
		response.end
	end if
	page=int(page)
	if newsid="" then
		Show_Err("δָ��������<br><br><a href='javascript:history.back()'>����</a>")
		response.end
	elseif not IsNumeric(newsid) then
		Show_Err("�Ƿ�������<br><br><a href='javascript:history.back()'>����</a>")
		response.end
	else
		'�жϸ�ƪ�����Ƿ����
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_EC_News_Table &" where NewsId="&NewsId
		rs.open sql,conn,3,3
		if rs.eof and rs.bof then
			rs.close
			set rs=nothing
			Show_Err("ָ�������²����ڣ�<br><br><a href='javascript:history.back()'>����</a>")
			response.end
		else
			checked=rs("checkked")
			if checked=1 or Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" then	'��������˲�������ع���Ա
				Click=rs("Click")
				if isnull(rs("Click"))  then
					conn.execute("update "& db_EC_News_Table &" Set Click=1 where NewsID=" & NewsID )
				else 
					conn.execute("update "& db_EC_News_Table &" Set Click=click+1 where NewsID=" & NewsID )
				end if
			end if
		
			rs.close
			set rs=nothing
		end if

		set rs=server.CreateObject("ADODB.RecordSet")
		if uselevel=1 then
			if Request.cookies(eChsys)("key")="" then
				rs.Source="select * from "& db_EC_News_Table &" where checkked=1 and newslevel=0 and newsid="&newsid
			end if
			if Request.cookies(eChsys)("key")="selfreg" then
				if Request.cookies(eChsys)("reglevel")=3 then
					rs.Source="select * from "& db_EC_News_Table &" where checkked=1 and newslevel<=3 and newsid="&newsid
				end if
				if Request.cookies(eChsys)("reglevel")=2 then
					rs.Source="select * from "& db_EC_News_Table &" where checkked=1 and newslevel<=2 and newsid="&newsid
				end if
				if Request.cookies(eChsys)("reglevel")=1 then
					rs.Source="select * from "& db_EC_News_Table &" where checkked=1 and newslevel<=1 and newsid="&newsid
				end if
			end if
		end if
			if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then
				rs.Source="select * from "& db_EC_News_Table &" where newsid="&newsid
			else
				rs.Source="select * from "& db_EC_News_Table &" where newsid="&newsid
			end if
			rs.Open rs.Source,conn,1,1
			E_BigClassID=rs("E_BigClassID")
			E_SmallClassID=rs("E_SmallClassID")
			title=htmlencode4(trim(rs("title")))
			title1=htmlencode4(trim(rs("title")))
			about=htmlencode4(trim(rs("about")))
			Author=htmlencode4(trim(rs("Author")))
			editor=htmlencode4(trim(rs("editor")))
			Original=htmlencode4(trim(rs("Original")))
			image=rs("image")
			UpdateTime=trim(rs("UpdateTime"))
			News_Content=rs("Content")
			SpecialID=rs("SpecialID")
			SpecialID2=rs("SpecialID2")
			click=rs("click")
			EnCode=trim(rs("EnCode"))
			E_typeid=rs("E_typeid")
			titletype=rs("titletype")
			titlecolor=rs("titlecolor")
			titleface=rs("titleface")
			editor=rs("editor")
			wzdj=rs("newslevel")
			backtype=rs("backtype")
			backtype1=rs("backtype1")
			rs.Close
			set rs=nothing

			set rs=server.CreateObject("ADODB.RecordSet")
			rs.Source="select * from "& db_EC_Type_Table &" where E_typeid=" & E_typeid
			rs.Open rs.Source,conn,1,1
			E_typename=rs("E_typename")
			rs.Close
			set rs=nothing

			set rs=server.CreateObject("ADODB.RecordSet")
			If(isnull(E_BigClassID)) Then
				E_BigClassName=" "
				E_BigClassID = 0
			Else
				rsSource="select * from "& db_EC_BigClass_Table &" Where E_BigClassID="& E_BigClassID
				rs.Open rsSource,conn,1,3
				If (rs.eof or rs.bof) Then
					E_BigClassName=" "
				else
					E_BigClassName=rs("E_BigClassName")
				End if
				rs.close
				set rs=nothing
			End if
			
			if E_SmallClassID<>"" then
				set rs=server.CreateObject("ADODB.RecordSet")
				rs.Source="select * from "& db_EC_SmallClass_Table &" Where E_SmallClassID=" & E_SmallClassID
				rs.Open rs.Source,conn,1,1
				E_smallclassname=rs("E_smallclassname")
				rs.close
				set rs=nothing
			end if%>
<%
'����ip����
	if cINT(backtype1)=1 then
	dim byte2
	byte2=split(newsipType,"|")
	for i=0 to ubound(byte2)
		if Request.serverVariables("REMOTE_ADDR")<>byte2(i) then
		Show_Err("�Բ��𣬴�ƪ����ֻ���ڲ����������ϵ����Ա����<br><br><a href='javascript:history.back()'>����</a>")
		Response.End
		end if
	next
	end if
%>

<html><head><meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=title%><%if E_SmallClassID<>"" then%>_<%=E_smallclassname%><%end if%>_<%=E_BigClassName%>_<%=E_typename%>_<%=eChSys%></title>
<LINK href="news.css" rel="stylesheet">
<script language="JavaScript" type="text/JavaScript">

function validateMode()
{
  if (!bTextMode) return true;
  alert("����ȡ���鿴HTMLԴ���룬���롰�༭��״̬��Ȼ����ʹ��ϵͳ�༭����!");
  message.focus();
  return false;
}
function validateModea()
{
  if (!bTextMode) return true;
  alert("����ȡ���鿴HTMLԴ����!");
  message.focus();
  return false;
}

function sign()
{if (!validateMode()) return;
message.focus();

var range =message.document.selection.createRange();
str1="<font color='red'><b>|ǩ��|</b>�ļ��Ѿ��Ķ�</font>"
   range.pasteHTML(str1);
}

function img_zoom(e, o)		//ͼƬ����������
{
	var zoom = parseInt(o.style.zoom, 10) || 100;
	zoom += event.wheelDelta / 12;
	if (zoom > 0) o.style.zoom = zoom + '%';
		return false;
}
</script>

<script language=javascript>
function user_login()
{
	document.UserLogin.UserName.focus();
	return false;
}
</script>

<script language="JavaScript" type="text/JavaScript">
<!--
function MM_preloadImages() { //v3.0
	var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
	var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
	if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}
//-->
</script>

<script language="JavaScript" type="text/JavaScript">

var size=14;			//�����С

function fontZoom(fontsize){	//��������
	size=fontsize;
	document.getElementById('fontZoom').style.fontSize=size+'px';
}

function fontMax(){	//����Ŵ�
	size=size+2;
	document.getElementById('fontZoom').style.fontSize=size+'px';
}

function fontMin(){	//������С
	size=size-2;
	if (size < 2 ){
		size = 2
	}
	document.getElementById('fontZoom').style.fontSize=size+'px';
}

</SCRIPT>

<script language="javascript">
<!--
function changedata() {
	if ( document.AddReview.none.checked ) {
		document.AddReview.Author.value = "�ο�";
	} 
}
function changedata1() {
	if ( document.AddReview.none1.checked ) {
		document.AddReview.email.value = "guest@echsys.com";
	} 
}

//-->
</script>

<script language=javascript>
function CheckFormAddReview()	//������۷�������д��Ŀ�Ƿ�Ϊ��
{
	if(document.AddReview.Author.value=="")
	{
		alert("�������û�����");
		document.AddReview.Author.focus();
		return false;
	}
	if(document.AddReview.email.value == "")
	{
		alert("����������EMAIL��");
		document.AddReview.email.focus();
		return false;
	}
	if(document.AddReview.content.value == "")
	{
		alert("�������������ݣ�");
		return false;
	}
}
</script>
<style type="text/css">
<!--
.STYLE1 {color: #FF3333}
-->
</style>
</head>
<%if ScrollClick = "double" then%>
	<SCRIPT language=JavaScript>
	//˫����������
	var currentpos,timer;
	
	function initialize()
	{
	timer=setInterval("scrollwindow()",50);
	}
	
	function sc(){
	clearInterval(timer);
	}
	
	function scrollwindow()
	{
	currentpos=document.body.scrollTop;
	window.scroll(0,++currentpos);
	if (currentpos != document.body.scrollTop)
		sc();
	}
	document.onmousedown=sc
	document.ondblclick=initialize
	</script>
<%else%>
	<SCRIPT>
	//�����϶���Ļ��ʽ
	var old_y=0;  //��¼��갴��ʱ�ģ���λ��
	var yn=0;  //���ڼ�¼���״̬
	function moveit()
	{
	if(yn==1 &&  event.button==1)  //������������¾��ƶ�ҳ��
	document.body.scrollTop=(old_y-event.clientY); //�ƶ�ҳ��
	}
	function downit()
	{old_y=event.clientY+document.body.scrollTop; //��ס��갴��ʱ�ģ���λ��
	yn=1; //���ڼ�ס����Ѱ���
	}
	
	function upit()
	{yn=0;}  //��ס����ѷſ�
	
	document.onmouseup=upit; //���ſ�ʱִ��upit()����
	document.onmousedown=downit; //��갴��ʱִ��downit()����
	document.onmousemove =moveit; //����ƶ�ʱִ��moveit()����
	</SCRIPT>
<%end if%>
<body topmargin="0" marginheight="0">
<SCRIPT language=javascript>
<!--
  function do_color(vobject,vvar)
  { document.getElementById(vobject).style.color=vvar; }
  function do_zooms(vobject,vvar)
  { document.getElementById(vobject).style.fontsize=vvar+'px'; }
-->
</SCRIPT>
<br><br>
<table width="80%" border="0" align="center" cellpadding="0" cellspacing="0"  bgcolor="#F1F7FC">
<tr><td width="80%" height="124" valign="top"  class="daohang" ><table width="780" height="94" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="212" height="31"><div align="center"></div></td>
    <!--
	<td width="303"><div align="center"><a href=send.asp?NewsID=<%=NewsID%> target=_blank  class="bottom">������Ϣ��������</a></div></td>
	-->
    <td width="97"><div align="center"><a href="javascript:window.print()" class="bottom">��ӡ��ҳ</a></div></td>
    <td width="78"><div align="center"><a href="javascript:window.close()" class="bottom">�ر�</a></div></td>
    <td width="90"><div align="center"></div></td>
  </tr>
  <tr>
    <td colspan="5"><div align="left">&nbsp;&nbsp;�������ڣ�<%=year(updateTime)%>��<%=month(updateTime)%>��<%=day(updateTime)%>��&nbsp; �����<font color=red><%=click%></font> �� &nbsp;
        <%if Original<>"" then%>
������<%=Original%>
<%end if%>

  </tr>
</table></td>
</tr>
<tr>
<td >
 <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="100%" align=center> <br>
                      <font color="<%=titlecolor%>" size="+3" face="����_GB2312"><strong><%=title1%></strong></font>
                  <HR align="center" width="95%" noshade style="color:#EFF5FA"></td>
                </tr>
				<tr>
               <td height="20" ><div align="center">�����ˣ�<%=editor%>&nbsp;&nbsp;
			    </td>
              </tr>
              <tr>
             <td>
                <%	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;" & Content%><br>
              </td>
             </tr>
                <%
set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from "& db_EC_News_Table &" where NewsID=" & NewsID
rs.Open rs.source,conn,1,1
E_typeid=rs("E_typeid")
Title=trim(rs("Title"))
image=rs("image")

dim mode
    
set rs11=server.CreateObject("ADODB.RecordSet")
rs11.Source="select * from "& db_EC_Type_Table &" where E_typeid="&E_typeid&" order by E_typeid"
rs11.Open rs11.Source,conn,1,1
mode=rs11("mode")
rs11.close
set rs11=nothing

''���ͼƬ����������Ч��
if mouse_wheel_zoom="on" then
	News_Content=replace(News_Content,"<IMG","<IMG onmousewheel='return img_zoom(event,this)' onload='javascript:if(this.width>screen.width-333)this.width=screen.width-333'",1,-1,1) 
end if
''ͼƬ�ϴ�·����ԭΪ config.asp ���趨�� [FileUploadPath] ֵ
News_Content=replace(News_Content,"="&chr(34)&"uploadfile/","="&chr(34)&FileUploadPath,1,-1,1)

arr_Content=split(News_Content,"[---��ҳ---]")
MaxPages=ubound(arr_Content)
%>
                <tr>
                  <td width="100%"  align="center">
                    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" style="TABLE-LAYOUT: fixed">
                      <tr>
                        <td width="100%" align=center></td>
                      </tr>
                      <tr>
                        <TD class=newstitle id=fontzoom vAlign=top><br>
                            <%if (checked<>1) and ((Request.cookies(eChsys)("key")<>"super") and (Request.cookies(eChsys)("key")<>"typemaster") and (Request.cookies(eChsys)("key")<>"bigmaster") and (Request.cookies(eChsys)("key")<>"smallmaster")) then	'����δ���,�����Ƿ���ع���Ա
	response.write "<P><CENTER><strong><font color='ff0000' size='+2' face='����'>���»�δ�������<br>��ȴ�����֪ͨ����Ա��˲���������</font></strong></CENTER>"
	response.write "<P><CENTER><IMG SRC='" &ReadNews_CopyRight_Logo& "' BORDER='0' ALT=''></CENTER>"
else	'���������
	if checked<>1 then
		response.write "<P><CENTER><strong><font color='ff0000' size='+2' face='����'>���ѣ������»�δͨ�����</font></strong></CENTER>"
	end if



if uselevel=1 then

	if cINT(wzdj)<1 then
		Response.Write arr_Content(Page-1)%>
                            <% else %>
                            <%if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then %>
                            <%Response.Write arr_Content(Page-1)%>
                            <% else %>
                            <%if  Request.cookies(eChsys)("key")="" then%>
                            <br>
                            <font color="#cc0000"><b>���ݼ�飺</b></font><br>
                            <br>
                            <%=CutStr(nohtml(rs("Content")),10)%>... <br>
                            <br>
                            <br>
                            <font color="#cc0000"><b>�������ѣ�</b></font><br>
                            <br>
                  �㻹ûע�᣿����û�е�¼������Ȩ�޲�������ƪ����Ҫ���Ǳ�վ����Ȩ��Ҫ���ע���û������Ķ���<br>
                  <br>
                  <%
				response.write "���¼���"
				response.write cINT(wzdj)
				response.write "��"
				%>
                  <br>
                  <%
				response.write "����Ȩ�ޣ�"
				response.write "δע��"
				%>
                  <br>
                  <br>
                  ����㻹ûע�ᣬ��Ͻ����<a href="#" OnClick="adduser()"  class=bottom><font color="#cc0000">ע��</font></a>�ɣ�<br>
                  <br>
                  ������Ѿ�ע�ᵫ��û��¼����Ͻ����<a href="#" OnClick="user_login()"  class=bottom><font color="#cc0000">��¼</font></a>�ɣ�<br>
                  <% else %>
                  <%if  Request.cookies(eChsys)("key")="selfreg" then
					if cINT(Request.cookies(eChsys)("reglevel"))>=cINT(wzdj) then%>
                  <%Response.Write arr_Content(Page-1)%>
                  <% else %>
                  <br>
                  <font color="#cc0000"><b>���ݼ�飺</b></font> <br>
                  <br>
                  <%=CutStr(nohtml(rs("Content")),10)%>... <br>
                  <br>
                  <br>
                  <font color="#cc0000"><b>�������ѣ�</b></font><br>
                  <br>
                  �㻹ûע�᣿����û�е�¼������Ȩ�޲�������ƪ����Ҫ���Ǳ�վ����Ȩ��Ҫ���ע���û������Ķ���<br>
                  <br>
                  <%
						response.write "���¼���"
						response.write cINT(wzdj)
						response.write "��"
						%>
                  <br>
                  <%
						response.write "����Ȩ�ޣ�"
						response.write (Request.cookies(eChsys)("reglevel"))
						response.write "��"
						%>
                  <br>
                  <br>
                  ����㻹ûע�ᣬ��Ͻ����<a href="#" OnClick="adduser()" class=bottom><font color="#cc0000">ע��</font></a>�ɣ�<br>
                  <br>
                  ������Ѿ�ע�ᵫ��û��¼����Ͻ����<a href="#" OnClick="user_login()"  class=bottom><font color="#cc0000">��¼</font></a>�ɣ�<br>
                  <br>
                  <%end if%>
                  <%end if%>
                  <%end if%>
                  <%end if%>
                  <%end if%>
                  <%else%>
                  <%Response.Write arr_Content(Page-1)%>
                  <%end if%>
                  <%end if%>
                  <br>
                  <div align="right">
                    <%
url="E_ReadNews.asp?NewsId="&newsid
if MaxPages >0 then
	Response.write "<a class=black href='"& Url &"&page=1' title='��1ҳ'>��ҳ</a> "
	if Page-1 > 0 then
		Prev_Page = Page - 1
		Response.write "<a class=black href='"& Url &"&page="& Prev_Page &"' title='��"& Prev_Page &"ҳ'>��һҳ</a> "
	end if

	for PageCounter=0 to MaxPages
		PageLink = PageCounter+1
		if PageLink <> Page Then
			Response.write "<a class=black href='"& Url &"&page="& PageLink &"'>["& PageLink &"]</a> "
		else
			Response.Write "<font color='#FF0000'><B>["& PageLink &"]</B></font> "
		end if
		If PageLink = MaxPages+1 Then Exit for
	Next
	if page <= Maxpages then
		bdd_Page = Page + 1
		Response.write "<a class=black href='" & Url & "&page=" & bdd_Page & "' title='��" & bdd_Page & "ҳ'>��һҳ</A>"
	end if
	Response.write " <A class=black href='" & Url & "&page=" & Maxpages+1 & "' title='��"& Maxpages+1 &"ҳ'>βҳ</A>"
end if
%>
                </div></td>
                      </tr>
                  </table></td>
                </tr>
                <tr>
                  <td width="100%"  height="25">
                    <div align="center">
                      <!--#include file=attach.asp -->
                  </div></td>
                </tr>
                <tr>
                  <td height="22"> <b><IMG SRC="IMAGES/006.gif"> ��һƪ��</b><a class=blacklink href="E_ReadNews.asp?NewsID=<%=getSideNewsID(E_BigClassID,newsid,-1)%>"><%=getnewstitle(getSideNewsID(E_BigClassID,newsid,-1))%></a><br>                  </td>
                </tr>
                <tr>
                  <td height="22"> <b><IMG SRC="IMAGES/006.gif"> ��һƪ��</b><a  class=blacklink href="E_ReadNews.asp?NewsID=<%=getSideNewsID(E_BigClassID,newsid,1)%>"><%=getnewstitle(getSideNewsID(E_BigClassID,newsid,1))%></a> </td>
                </tr>
                
<!--
              <tr>
                  <td width=100% ><hr size=1 noshade style="color:#EFF5FA"></td>
                </tr>
                <tr>
                  <td  height=25>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="32%" valign="top" bgcolor="#EFF5FA">
                          <table width="100%" border="0" cellpadding="3" cellspacing="0" bgcolor="#EFF5FA">
                            <tr>
                              <td height="25" bgcolor="#E2EDF5" class="white_link">&nbsp;<B class="yellow_title">���ר�⣺</b></td>
                            </tr>
                            <tr>
                              <td bgcolor="#EFF5FA"><%set rs4=server.CreateObject("ADODB.RecordSet")
if SpecialID<>0 then
	rs4.Source="select * from "& db_EC_Special_Table &" where SpecialID=" & SpecialID
	rs4.Open,conn,1,3
	if not rs4.eof then%>
                                  <br>
                                  <a class=class href='Special_News.asp?SpecialID=<%=SpecialID%>'>��<%=CutStr(trim(rs4("E_SpecialName")),28)%></a>
                                  <%			
    end if
	rs4.Close
	set rs4=nothing

else

	Response.Write "<br><font color=red >��ר��1��Ϣ��</font>"

	end if%>
                                  <br>
                                  <br>
                                  <%
	set rs4=server.CreateObject("ADODB.RecordSet")
	if specialid2<>0 then
		rs4.Source="select * from "& db_EC_Special_Table &" where SpecialID=" & SpecialID2
		rs4.Open,conn,1,3
		if not rs4.eof then %>
                                  <a class=class href='Special_News.asp?SpecialID=<%=SpecialID2%>'>��<%=CutStr(trim(rs4("E_SpecialName")),28)%></a>
                              <%end if
		rs4.Close
		set rs4=nothing

else

	Response.Write "<font color=red >��ר��2��Ϣ��</font>"

	end if
	%>                              </td>
                            </tr>
                        </table></td>
                        <td width="1%">&nbsp;</td>
                        <td width="34%" valign="top" bgcolor="#EFF5FA"><table width="100%" border="0" cellspacing="0" cellpadding="3">
                            <tr>
                              <td height="25" bgcolor="#E2EDF5" class="white_link">&nbsp;<strong class="yellow_title"> �������£�</strong></td>
                            </tr>
                            <%
set rs=server.CreateObject("ADODB.RecordSet")
if uselevel=1 then
if Request.cookies("key")="super" or Request.cookies("key")="typemaster" or Request.cookies("key")="bigmaster" or Request.cookies("key")="smallmaster" or Request.cookies("key")="check" then
rs.Source="select top 4 * from "& db_EC_News_Table &" where  checkked=1 order by click DESC,newsid desc"
end if
if Request.cookies("key")="" then
rs.Source="select top 4 * from "& db_EC_News_Table &" where (checkked=1 and newslevel=0) order by click DESC,newsid desc"
end if
if Request.cookies("key")="selfreg" then
	if Request.cookies("reglevel")=3 then
rs.Source="select top 4 * from "& db_EC_News_Table &" where (checkked=1 ) order by click DESC,newsid desc"
	end if
	if Request.cookies("reglevel")=2 then
rs.Source="select top 4 * from "& db_EC_News_Table &" where (checkked=1 ) order by click DESC,newsid desc"
	end if
	if Request.cookies("reglevel")=1 then
rs.Source="select top 4 * from "& db_EC_News_Table &" where (checkked=1 ) order by click DESC,newsid desc"
	end if
end if
else
rs.Source="select top 4 * from "& db_EC_News_Table &" where  checkked=1 order by click DESC,newsid desc"
end if
rs.Open rs.Source,conn,1,1
while not rs.EOF
title=trim(rs("title"))
title=replace(title,"<br>","")
%>
                            <tr>
                              <td> &nbsp;�� <a class=middle href="E_ReadNews.asp?NewsID=<%=rs("NewsID")%>" target="_blank" title="<%=title%>"><font color="<%=rs("titlecolor")%>">
                                <%if len(title)>13 then%>
                                <%=left(title,13)%>
                                <%else%>
                                <%=title%>
                                <%end if%>
                              </font></a>[<%=rs("click")%>] </td>
                            </tr>
                            <%
rs.MoveNext
wend
rs.close
%>
                        </table></td>
                        <td width="1%">&nbsp;</td>
                        <td width="32%" valign="top" bgcolor="#EFF5FA">
                          <table width="100%" border="0" cellspacing="0" cellpadding="3">
                            <tr>
                              <td height="25" bgcolor="#E2EDF5" class="yellow_title">&nbsp;<B class="yellow_title">������£�</b></td>
                            </tr>
                            <%
Response.Write ""
Response.Write ""
if about<>""  then
	sql="select top 4 * from "& db_EC_News_Table &" where (about like '%" & about & "%' or title like '%" & about & "%') and checkked=1 order by newsid desc"
	set rs=conn.execute(sql)

	do while not rs.eof
		title=htmlencode4(trim(rs("title")))
		%>
                            <tr>
                              <td> <img src=images/goxp.gif> <a class=middle target=_top href='E_ReadNews.asp?NewsID=<%=rs("NewsID")%>'><%=Title%></a>
                                  <!--<font color=#666666>(<%=month(trim(rs("updateTime")))%>��<%=day(trim(rs("updateTime")))%>��)-->
                                  [<font color=#666666><%=rs("click")%></font>] </td>
                            </tr>
                            <% rs.movenext
	loop
	Response.Write "<tr><td align='right'><a class=lift  href='Result.asp?keyword=" & about &"'><img src='images/more.gif' border='0'></a></td></tr>"
	rs.close   
	set rs=nothing  
else
	Response.write "<tr><td><font color=red ><br>��û���������</font></td></tr>"
end if
%>
                        </table></td>
                      </tr>
                  </table></td>
                </tr>

-->
                <%
set rs1=server.CreateObject("ADODB.RecordSet")
rs1.Source="select top "&reviewnum&" * from "& db_EC_Review_Table &" where NewsID=" & NewsID & " and passed=1 order by reviewid desc"
rs1.Open rs1.Source,conn,1,1
if rs1.EOF then  NoReview=1
	Response.Write "<tr><td width=100% ><hr size=1 noshade style=color:#EFF5FA></td></tr><tr><td height=8></td></tr>"
	Response.Write "<tr><td width=100% ><B></b></td></tr><tr><td height=8></td></tr>"
	%>
                <tr>
                  <td width="100%">
                    <%
	if NoReview then 						
		Response.Write "<font color=red ><b>���������</b></font>"
	end if
	%>                  </td>
                </tr>
                <%
	if not NoReview then
		while not rs1.EOF
			author=server.HTMLEncode(trim(rs1("author")))
			email=server.HTMLEncode(trim(rs1("email")))
			reviewip=rs1("reviewip")
			content=trim(rs1("content"))
            content=replace(content,"table","�������")
			content=replace(content,"TABLE","�������")
			ContentLen=len(Content)
			%>
                <tr>
                  <td>
                    <table width="100%" border="0" cellspacing="0" cellpadding="5" style="table-layout:fixed; word-break:break-all">
                      <tr bgcolor="#EFEFEF">
                        <td width="322" bgcolor="#EFF5FA">�����ˣ�<%=author%></td>
                        <td width="270" bgcolor="#EFF5FA">
                          <p align="right">
                            <%if Request.cookies(eChsys)("key")="super" or showip="1" then%>
                    IP��<%=reviewip%>
                    <%end if%>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="2">
                          <table width="100%" border="0" cellspacing="0" cellpadding="0" style="table-layout:fixed; word-break:break-all">
                            <tr>
                              <td>�������ʼ���<a href='mailto:<%=email%>'><%=email%></a></td>
                              <td align="right">����ʱ�䣺<%=rs1("updatetime")%></td>
                            </tr>
                        </table></td>
                      </tr>
                      <tr bgcolor="#FFFFFF">
                        <td colspan="2"><%
			If ContentLen=<50  then	
				DisplayContent=Content
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;" & displaycontent
										%></td>
                      </tr>
                      <%
			else
									%>
                      <tr bgcolor="#FFFFFF">
                        <td colspan="2"><%
				DisplayContent=replace(nohtml(trim(Content)),"&nbsp;","",1,-1,1) '��ȡ����������ֶ����ݲ��滻��ʽ��
				DisplayContent=replace(DisplayContent,vbcrlf,"",1,-1,1)
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"& CutStr(displaycontent,50) &"<a href='ReadView.asp?ReviewID=" & rs1("ReviewID") &"&NewsID=" & NewsID &"' target=_blank class=class>��ϸ����</a>"
			end if
										%></td>
                      </tr>
                  </table></td>
                </tr>
                <%
			rs1.MoveNext
		wend
	end if

	rs1.Close
	set rs1=nothing	
						%>
                <tr>
                  <td width="98%" height="28"></td>
                </tr>
    </table></td>
</tr>
<tr valign="top">
    <td background="Images/table_bg.gif">
        <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
          <%if reviewable="1" then%>
          <form method="POST" action="AddReview.asp" name=AddReview  onSubmit="return CheckFormAddReview();">
            <%if Request.cookies(eChsys)("username")<>"" then
		
				%>
            <tr>
              <td width="100%" align="center">
                <table border="0" cellspacing="0" width="100%" cellpadding="4">
                  <input type=hidden name=NewsID value=<%=NewsID%>>
                  <tr>
                    <td width="15%" align="left"><a name="send"><font color="#FF0000">*</font>��&nbsp;��&nbsp;����</a></td>
                    <td width="35%"><input type="text" name="Author" size="30" value="<%=Request.cookies(eChsys)("UserName")%>" readonly></td>
                    <td width="15%" align="left"><font color="#FF0000">*</font>�����ʼ���</td>
                    <td width="35%"><input type="text" name="email" size="30" value="<%=Request.cookies(eChsys)("UserEmail")%>" ></td>
                  </tr>
                  <tr>
                    <td colspan="4" align="left" valign="top"><div align="center"><font color="#FF0000">*</font>��100�����ڣ�
                            <% if M_MAIN=1 then %>
                            <font color=red >ǩ���ļ�</font>�� <img class=None  src="images/watermark.gif" align="absmiddle" border="0" style="cursor:hand;" alt="ǩ���ļ�" lANGUAGE="javascript" onClick="sign()"></div></td>
                    <% end if%>
                  </tr>
                  <tr>
                    <td width="70%" colspan="4" align="center" valign="top">
                      <script language="javascript">
									var bTextMode=false;
									document.write ('<iframe src="guestbox.asp?action=new" id="message" width="500" height="100" frameborder=yes scrolling=no align=center></iframe>')
									frames.message.document.designMode = "On";
									</script>                    </td>
                  </tr>
                  <tr>
                    <td width="70%" colspan="4" align="center" height="50">
                      <input type="hidden" name="editor" value="<%=editor%>">
                      <input name="passed" type="hidden" value="<%if reviewcheck="1" then%>1<%else%>0<%end if%>">
                      <input type="submit" value="��������" name="Submit3" OnClick="document.AddReview.content.value = frames.message.document.body.innerHTML;">
                      <input name="reset" type=reset value="������д">
                      <input type="hidden" name="content" value="">
                      <input type="hidden" name="title" value="���ۣ�<%=title%>">                    </td>
                  </tr>
              </table></td>
            </tr>
            <%else%>
            <tr>
              <td width="100%" align="center">
                <table border="0" cellspacing="0" width="100%" cellpadding="4">
                  <input type=hidden name=NewsID value=<%=NewsID%>>
                  <tr>
                    <td width="15%" align="left"><a name="send"><font color="#FF0000">*</font>��&nbsp;��&nbsp;����</a></td>
                    <td width="35%">
                      <input type="text" name="Author" size="18">
        &nbsp;�οͣ�
                      <INPUT onclick=changedata() type=checkbox value=true  name=none></td>
                    <td width="15%" align="left"><font color="#FF0000">*</font>�����ʼ���</td>
                    <td width="35%"><input name="email" type="text" size="18" value="">
&nbsp;�οͣ�
              <INPUT onclick=changedata1() type=checkbox value=true  name=none1></td>
                  </tr>
                  <tr>
                    <td colspan="4" align="left" valign="top" ><div align="center"><font color="#FF0000">*</font></div></td>
                  </tr>
                  <tr>
                    <td width="70%" colspan="4" align="center" valign="top">
                      <script language="javascript">
									document.write ('<iframe src="guestbox.asp?action=new" id="message" width="500" height="100" frameborder=yes scrolling=no align=center></iframe>')
									frames.message.document.designMode = "On";
									</script>                    </td>
                  </tr>
                  <tr>
                    <td width="70%" colspan="4" align="center" height="34">
                      <input type="hidden" name="editor" value="<%=editor%>">
                      <input name="passed" type="hidden" value="<%if reviewcheck="1" then%>1<%else%>0<%end if%>">
                      <input type="submit" value="��������" name="Submit3" OnClick="document.AddReview.content.value = frames.message.document.body.innerHTML;">
                      <input name="reset2" type=reset value="������д">
                      <input type="hidden" name="content" value="">
                      <input type="hidden" name="title" value="���ۣ�<%=title%>">                    </td>
                  </tr>
              </table></td>
            </tr>
            <%end if%>
          </form>
          <%end if%>
      </table>
        </td>
  </tr>
<tr valign="top">
  <td height="47" valign="bottom" background="Images/news_bottom.gif"><div align="center">
    <table width="780" height="33" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="210">&nbsp;</td>
		<!--
        <td width="304"><div align="center"><a href=Review.asp?NewsID=<%=NewsID%> target="_blank" class="bottom"><img src="images/icon1.gif"  border="0">�����鿴������ڸ���Ϣ������</a></div></td>
		-->
        <td width="95"><div align="center"><a href="javascript:window.print()" class="bottom">��ӡ��ҳ</a></div></td>
        <td width="82"><div align="center"><a href="javascript:window.close()" class="bottom">�ر�</a></div></td>
        <td width="89">&nbsp;</td>
      </tr>
    </table>
  </div></td>
</tr>
</table>
<br><br>
</body>
</html>
<%end if%>
<%
conn.close
set conn=nothing
%> 