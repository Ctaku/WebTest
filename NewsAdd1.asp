<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="chkuser.asp"-->
<%dim jingyong
E_SmallClassID=ChkRequest(request("E_SmallClassID"),1)	'��ע��
E_BigClassID=ChkRequest(request("E_BigClassID"),1)	'��ע��
E_typeid=ChkRequest(request("E_typeid"),1)	'��ע��
set rs=server.createobject("adodb.recordset")
sql="select * from "& db_User_Table &" where  "& db_User_Name &"='"&request.cookies(eChsys)("username")&"'"
rs.open sql,ConnUser,1,3
jingyong=rs("jingyong")
rs.close
set rs=nothing

set rs8=server.createobject("adodb.recordset")
sql8="select * from "& db_EC_System_Table
rs8.open sql8,conn,1,3
name=rs8("name")

randomize
rannum=int(90000*rnd)+10000
Response.cookies(eChsys)("newsrelated")=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&rannum

if request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="check" or request.cookies(eChsys)("KEY")="bigmaster" or request.cookies(eChsys)("KEY")="smallmaster" or (request.cookies(eChsys)("key")="selfreg" and jingyong=0) or request.cookies(eChsys)("KEY")="typemaster" then%>

<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - �������</title>
<script src="ckeditor/ckeditor.js"></script>
<script language="javascript">
<!--
function checkdata()
{
if (document.form1.title.value=="")
	{
	  alert("�Բ������������±��⣡")
	  document.form1.title.focus()
	  return false
	 }
}

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>

<link rel="stylesheet" type="text/css" href="site.css">
<link rel="STYLESHEET" type="text/css" href="editor.css">
</head>
<body topmargin="0">

<div id=menuDiv style='Z-INDEX: 1000; VISIBILITY: hidden; WIDTH: 1px; POSITION: absolute; HEIGHT: 1px; BACKGROUND-COLOR: #9cc5f8'></div>
<%set rs5=server.CreateObject ("ADODB.RecordSet")
rs5.Source="select * from "& db_EC_System_Table &""
rs5.Open rs5.source,conn,1,1
%>
<div align="center"><center>
<table border="0" cellspacing="0" width="100%" cellpadding="0">
<form name="form1" method="post" action="NewsAdd2.asp" OnSubmit="return checkdata()" onReset="return ResetForm();">
<tr align="center"> 
  <td width="100%"> 
	<table border="1" cellspacing="0" width="100%" cellpadding="0"  bordercolor="<%=border%>" style="border-collapse: collapse" height="463">
	<tr align="center">
<%if E_SmallClassID<>"" then
dim E_typename,E_BigClassName,E_smallclassname
set rs1=server.createobject("adodb.recordset")
sql1="select * from "& db_EC_SmallClass_Table &" where E_SmallClassID="&E_SmallClassID
rs1.open sql1,conn,1,3
E_smallclassname=rs1("E_smallclassname")
set rs2=server.createobject("adodb.recordset")
sql2="select * from "& db_EC_BigClass_Table &" where E_BigClassID="&rs1("E_BigClassID")
rs2.open sql2,conn,1,3
E_BigClassName=rs2("E_BigClassName")
set rs3=server.createobject("adodb.recordset")
sql3="select * from "& db_EC_Type_Table &" where E_typeid="&rs2("E_typeid")
rs3.open sql3,conn,1,3
E_typename=rs3("E_typename")
%>
	<input name="E_typeid" type="hidden" value="<%=rs2("E_typeid")%>">
	<input name="E_BigClassID" type="hidden" value="<%=rs1("E_BigClassID")%>">
	<input name="E_SmallClassID" type="hidden" value="<%=E_SmallClassID%>">
	<td colspan="2" height="25"  valign="middle" class="TDtop">
		<div align="center" >�� <B>��<b class="unnamed2">[<a href="E_BigClassManage.asp?E_typeid=<%=rs2("E_typeid")%>"><%=E_typename%></a>]-[<a href="E_SmallClassManage.asp?E_BigClassID=<%=rs1("E_BigClassID")%>"><%=E_BigClassName%></a>]-[<a href="ListNews.asp?E_SmallClassID=<%=E_SmallClassID%>"><%=E_smallclassname%></a>]</b>���� �� �� ��</B> ��</div>
	</td>
<%rs3.close
set rs3=nothing
rs2.close
set rs2=nothing
rs1.close
set rs1=nothing
else%>
<%if E_BigClassID<>"" then
set rs1=server.createobject("adodb.recordset")
sql1="select * from "& db_EC_BigClass_Table &" where E_BigClassID="&E_BigClassID
rs1.open sql1,conn,1,3
E_BigClassName=rs1("E_BigClassName")
bigclasszs=rs1("bigclasszs")
set rs2=server.createobject("adodb.recordset")
sql2="select * from "& db_EC_Type_Table &" where E_typeid="&rs1("E_typeid")
rs2.open sql2,conn,1,3
E_typename=rs2("E_typename")
%>
	<input name="E_typeid" type="hidden" value="<%=rs1("E_typeid")%>">
	<input name="E_BigClassID" type="hidden" value="<%=E_BigClassID%>">
	<td colspan="2"  height="29"  class="TDtop" valign="middle" bgcolor="<%=m_top%>" width="100%">��<b class="unnamed2">
		<a href="E_BigClassManage.asp?E_typeid=<%=rs2("E_typeid")%>"><%=E_typename%></a>-<a href="E_SmallClassManage.asp?E_BigClassID=<%=rs1("E_BigClassID")%>"><%=E_BigClassName%></a></b>���� �� �� ��
	</td>
<%rs2.close
set rs2=nothing
rs1.close
set rs1=nothing
else
%>

<%
if E_typeid="" then
response.write "<script>alert('������ѡ��������');history.go(-1);</Script>"
response.end
else
set rs1=server.createobject("adodb.recordset")
sql1="select * from "& db_EC_Type_Table &" where E_typeid="&E_typeid
rs1.open sql1,conn,1,3
E_typename=rs1("E_typename")
mode=rs1("mode")
%>
	<input name="E_typeid" type="hidden" value="<%=E_typeid%>">
	
	<td colspan="2"  height="29"  class="TDtop" valign="middle" bgcolor="<%=m_top%>" width="100%">��<b class="unnamed2">
		<a href="E_BigClassManage.asp?E_typeid=<%=rs1("E_typeid")%>"><%=E_typename%></a></b>���� �� �� ��
	</td>
<%
rs1.close
set rs1=nothing
end if 
end if
end if

%>

	</tr>
<%if mode<>5 and bigclasszs<>5 then%>
	<tr> 
	<td width="10%" align="right" class="unnamed2" valign="middle" bgcolor="#FFFFFF" >����ר�⣺</td>
	<td width="90%" bgcolor="#FFFFFF" height="25"> &nbsp; 
	<select name="SpecialID" size="1" onMouseOver="window.status='�������ѡ������ר��';return true;" onMouseOut="window.status='';return true;">
                    <%			
set rs3=server.CreateObject ("ADODB.RecordSet")
rs3.Source="select * from "& db_EC_Special_Table &""
rs3.Open rs3.source,conn,1,1			
%>
	<option value=0>�������κ�ר��</option>
                    <%if rs3.EOF then %>
	<option value=0>�����κ�ר��</option>
                    <%else
while not rs3.EOF
%>
	<option value="<%=rs3("SpecialID")%>"><%=trim(rs3("E_SpecialName"))%></option>
                    <%
rs3.MoveNext
wend
end if
rs3.close	
%>
	</select>
&nbsp;�ڶ�ר�⣺<select name="SpecialID2" size="1" onMouseOver="window.status='�������ѡ������ר��';return true;" onMouseOut="window.status='';return true;">
                    <%			
set rs3=server.CreateObject ("ADODB.RecordSet")
rs3.Source="select * from "& db_EC_Special_Table &""
rs3.Open rs3.source,conn,1,1			
%>
	<option value=0>�������κ�ר��</option>
                    <%if rs3.EOF then %>
	<option value=0>�����κ�ר��</option>
                    <%else
while not rs3.EOF
%>
	<option value="<%=rs3("SpecialID")%>"><%=trim(rs3("E_SpecialName"))%></option>
                    <%
rs3.MoveNext
wend
end if
rs3.close	
%>
	</select>
	</td>
	</tr>
	<%end if%>
	<tr height="25"> 
	<td align="right" class="unnamed2" valign="middle" bgcolor="#FFFFFF">���±��⣺</td>
	<td bgcolor="#FFFFFF"> <span style='cursor:hand' title='���̶Ի���' onClick='if (me.size>10)me.size=me.size-2'>-</span> 
	<input name="title" id=me type="TEXT" size=60 maxlength=100 style=" height:22;background-color:ffffff;border: 1 double; border-color:#BDD0FD;"  onMouseOver="window.status='����������Ҫ��ӵ����±��⣬����';return true;" onMouseOut="window.status='';return true;" title='����������Ҫ��ӵ����±��⣬����'>
	<span style='cursor:hand' title='�ӳ��Ի���' onClick='if (me.size<102)me.size=me.size+2'>+</span> <font color="#FF0000">*</font>
	<SELECT size=1 id=title_color name=title_color LANGUAGE=javascript onChange="return title_color_onchange()" style="width:55">
	<OPTION value="" selected>��ɫ</OPTION>
	<OPTION value="">ȱʡ</OPTION>
	<OPTION value="#000000" style="background-color:#000000"></OPTION>
	<OPTION value="#FFFFFF" style="background-color:#FFFFFF"></OPTION>
	<OPTION value="#008000" style="background-color:#008000"></OPTION>
	<OPTION value="#800000" style="background-color:#800000"></OPTION>
	<OPTION value="#808000" style="background-color:#808000"></OPTION>
	<OPTION value="#000080" style="background-color:#000080"></OPTION>
	<OPTION value="#800080" style="background-color:#800080"></OPTION>
	<OPTION value="#808080" style="background-color:#808080"></OPTION>
	<OPTION value="#FFFF00" style="background-color:#FFFF00"></OPTION>
	<OPTION value="#00FF00" style="background-color:#00FF00"></OPTION>
	<OPTION value="#00FFFF" style="background-color:#00FFFF"></OPTION>
	<OPTION value="#FF00FF" style="background-color:#FF00FF"></OPTION>
	<OPTION value="#FF0000" style="background-color:#FF0000"></OPTION>
	<OPTION value="#0000FF" style="background-color:#0000FF"></OPTION>
	<OPTION value="#008080" style="background-color:#008080"></OPTION>
	</SELECT>

	<SELECT size=1 id=title_type name=title_type LANGUAGE=javascript onChange="return title_type_onchange()" style="width:55">
	<OPTION value="0" selected>����</OPTION>
	<OPTION value="0">��ͨ</OPTION>
	<OPTION value="b">�Ӵ�</OPTION>
	<OPTION value="i">��б</OPTION>
	<OPTION value="u">�»���</OPTION>
	</SELECT>

	</SELECT>
	<SELECT size=1 id=title_size name=title_size LANGUAGE=javascript onChange="return title_size_onchange()" style="width:55">
	<OPTION value="0" selected>����</OPTION>
	<OPTION value="0">����</OPTION>
	<OPTION value="1">Ҫ��</OPTION>
	</SELECT>

	</td>
	</tr>
<%if mode<>5 and bigclasszs<>5 then%>
	<tr height="25"> 
	<td align="right" class="unnamed2" valign="middle" bgcolor="#FFFFFF" >�������£�</td>
	<td bgcolor="#FFFFFF" > <span style='cursor:hand' title='���̶Ի���' onClick='if (title_face.size>10)title_face.size=title_face.size-2'>-</span> 
	<input name="title_face" id=title_face type="TEXT" size=30 maxlength=100 style=" height:22;background-color:ffffff;border: 1 double; border-color:#BDD0FD;" onMouseOver="window.status='����Ϊֱ����������ʱ������д��ַ';return true;" onMouseOut="window.status='';return true;" title='����Ϊֱ����������ʱ������д��ַ'>
	<span style='cursor:hand' title='�ӳ��Ի���' onClick='if (title_face.size<102)title_face.size=title_face.size+2'>+</span> 
	(����Ϊֱ����������ʱ������д��ַ)</td>
	</tr>
	<tr height="25"> 
                <td align="right" class="unnamed2" valign="middle" bgcolor="#FFFFFF" >������Դ��</td>
                <td bgcolor="#FFFFFF" > <span style='cursor:hand' title='���̶Ի���' onClick='if (message.size>10)message.size=message.size-2'>-</span> 
                  <input name="Original" id=message type="TEXT" size=30 maxlength=100 style=" height:22;background-color:ffffff;border: 1 double; border-color:#BDD0FD;" onMouseOver="window.status='����У�����������������Դ';return true;" onMouseOut="window.status='';return true;" title='����У�����������������Դ'>
                  <span style='cursor:hand' title='�ӳ��Ի���' onClick='if (message.size<102)message.size=message.size+2'>+</span> 
                (��ҳģ��ʱ������д��ַ)</td>
              </tr>
              <tr height="25"> 
                <td align="right" class="unnamed2" valign="middle" bgcolor="#FFFFFF" >�������ߣ�</td>
                <td bgcolor="#FFFFFF" ><span style='cursor:hand' title='���̶Ի���' onClick='if (mess.size>10)mess.size=mess.size-2'>-</span> 
                  <input value="" name="Author" id=mess type="TEXT" size=30 maxlength=100 style=" height:22;background-color:ffffff;border: 1 double; border-color:#BDD0FD;" onMouseOver="window.status='����У�������������������';return true;" onMouseOut="window.status='';return true;" title='����У�������������������'>
                  <span style='cursor:hand' title='�ӳ��Ի���' onClick='if (mess.size<102)mess.size=mess.size+2'>+</span> 
                </td>
              </tr>
<%end if%>
              <tr> 
                <%
    '���߼������ TextBox.asp
    Dim strContentForEditor
    Dim p_newsID
    p_newsID = ChkRequest(Request.QueryString("newsID"),1)

    If p_newsID <> "" Then
        set rs_content=server.createobject("adodb.recordset")
        sql_content="select * from "& db_EC_News_Table &" where newsid=" & p_newsID
        rs_content.open sql_content,conn,1,1
        If Not rs_content.Eof Then
            strContentForEditor=rs_content("Content")
            'ͼƬ·���滻����ԭʼ TextBox.asp �е��߼���ͬ
            strContentForEditor=replace(strContentForEditor,"="&chr(34)&"uploadfile/","="&chr(34)&FileUploadPath,1,-1,1) 
        End If
        rs_content.close
        set rs_content = nothing
    Else
        ' ���������£������Ҫ�����Դ� cookie ���أ�������
        strContentForEditor = request.cookies(eChsys)("content") 
    End If
%>

<td align="right" valign="top" class="unnamed2" bgcolor="#FFFFFF" height="145">
    �������ݣ�<br><font color="#FF0000">*</font>
</td>
<td bgcolor="#FFFFFF">
    <%-- ���� CKEditor �����ı��� --%>
    <textarea name="content" id="content" rows="10" cols="80">
        <%=strContentForEditor%>
    </textarea>
</td>
              </tr>
			  <%if mode<>5 and bigclasszs<>5 then%>
              <tr height="25"> 
                <td align="right" class="unnamed2" valign="middle" bgcolor="#FFFFFF" >������£�</td>
                <td bgcolor="#FFFFFF" ><span style='cursor:hand' title='���̶Ի���' onclick='if (ss.size>10)ss.size=ss.size-2'>-</span> 
                  <INPUT NAME="about" id=ss TYPE="TEXT" SIZE=30 maxlength=100 style=" height:22;background-color:ffffff;border: 1 double; border-color:#BDD0FD;" onMouseOver="window.status='����У�����������������µĹؼ��ֻ���������';return true;" onMouseOut="window.status='';return true;" title='����У�����������������µĹؼ��ֻ���������'>
                  <span style='cursor:hand' title='�ӳ��Ի���' onclick='if (ss.size<102)ss.size=ss.size+2'>+</span> 
                  ����ؼ��ֻ��������� </td>
              </tr><tr> 
                <td align="right" class="unnamed2" valign="middle" bgcolor="#FFFFFF">��ҳͼƬ��</td>
                <td bgcolor="#FFFFFF">&nbsp;&nbsp;<input name="PicUrl" type="text" id="PicUrl" size="30" maxlength="200" style=" height:22;background-color:ffffff;border: 1 double; border-color:#BDD0FD;">
              ��������ҳ��ͼƬ���´���ʾ����ֱ�Ӵ��ϴ�ͼƬ��ѡ��<select name="PicList" id="PicList" onChange="PicUrl.value=this.value;">
                <option selected>��ָ����ҳͼƬ</option>
              </select>
              </tr>
			  <tr height="25"> 
                <td align="right" valign="middle" class="unnamed2" bgcolor="#FFFFFF" >����ͼƬ��</td>
                <td valign="middle" bgcolor="#FFFFFF" >&nbsp;
				<input type="radio" value="1" name="getphoto" onMouseOver="window.status='������������Զ��ͼƬ';return true;" onMouseOut="window.status='';return true;" title='������������Զ��ͼƬ'>
                  �� 
                  <input type="radio" value="0" checked name="getphoto">
                  ��&nbsp;&nbsp;����Զ��ͼƬ&nbsp;&nbsp;(��������վ�ϸ������ݣ����������а�����ͼƬ)</td>
              </tr>

              <%if request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="bigmaster" or request.cookies(eChsys)("key")="typemaster" then%>
              <tr height="25"> 
                <td align="right" valign="middle" class="unnamed2" bgcolor="#FFFFFF" >�Ƽ����£�</td>
                <td valign="middle" bgcolor="#FFFFFF" >&nbsp; 
                  <input type="radio" value="1" name="goodnews" onMouseOver="window.status='�������������Ƽ�����';return true;" onMouseOut="window.status='';return true;" title='�������������Ƽ�����'>
                  �� 
                  <input type="radio" value="0" checked name="goodnews">
                  ��&nbsp;&nbsp;�����Ƽ�����</td>
              </tr>
			    <tr height="25"> 
                <td align="right" valign="middle" class="unnamed2" bgcolor="#FFFFFF" >�����ͷ��</td>
                <td valign="middle" bgcolor="#FFFFFF" >&nbsp; 
                  <input type="radio" value="1" name="daodu" onMouseOver="window.status='�����������ɴ���⵼��';return true;" onMouseOut="window.status='';return true;" title='�����������ɴ���⵼��'>
                  �� 
                  <input type="radio" value="0" checked name="daodu">
                  ��&nbsp;&nbsp;���ɴ����ͷ</td>
              </tr>
<tr height="25"> 
                <td align="right" valign="middle" class="unnamed2" bgcolor="#FFFFFF" >���¹̶���</td>
                <td valign="middle" bgcolor="#FFFFFF" >&nbsp; 
                  <input type="radio" value="1" name="istop" onMouseOver="window.status='�����������ɹ̶�����';return true;" onMouseOut="window.status='';return true;" title='�����������ɹ̶�����'>
                  �� 
                  <input type="radio" value="0" checked name="istop">
                  ��&nbsp;&nbsp;���ɹ̶�����</td>
              </tr>
<tr height="25"> 
                <td align="right" valign="middle" class="unnamed2" bgcolor="#FFFFFF" >�Զ�������</td>
                <td valign="middle" bgcolor="#FFFFFF" >&nbsp; 
                  <input type="radio" value="1" name="backtype" onMouseOver="window.status='�������������������Ű�';return true;" onMouseOut="window.status='';return true;" title='�������������������Ű�'>
                  �� 
                  <input type="radio" value="0" checked name="backtype">
                  ��&nbsp;&nbsp;�������������Ű�</td>
</tr>
<tr height="25"> 
                <td align="right" valign="middle" class="unnamed2" bgcolor="#FFFFFF" >�ڲ����£�</td>
                <td valign="middle" bgcolor="#FFFFFF" >&nbsp; 
                  <input type="radio" value="1" name="backtype1" onMouseOver="window.status='��������Ϊ�ڲ��������';return true;" onMouseOut="window.status='';return true;" title='��������Ϊ�ڲ��������'>
                  �� 
                  <input type="radio" value="0" checked name="backtype1">
                  ��&nbsp;&nbsp;�ڲ��������</td>
</tr>

<%if uselevel=1 then%>
<tr height="25"> 
                <td align="right" bgcolor="#FFFFFF" height="30">���¼���</td>
                <td bgcolor="#FFFFFF" height="30"> &nbsp; <select size="1" name="newslevel" onMouseOver="window.status='������ѡ����Ҫ��ӵ����·ּ��������';return true;" onMouseOut="window.status='';return true;">
<%for i=0 to 3%>
                    <option value="<%=i%>"><%=i%></option>
<%next%>
                    </select>���¼���ָ����������»�Ա�ļ���0Ϊ�οͣ�1Ϊ��ͨ��Ա��2Ϊ�߼���Ա��3Ϊ�ؼ���Ա��
                  </td>
              </tr>
              <%end if%><%end if%><%end if%>
<tr align="center" height="25"> 
<td colspan="2" height="25" width="100%" clacc="TDtop">
	<input name="newsrelated" type="hidden" value="<%=Request.cookies(eChsys)("newsrelated")%>">
	<input name="checkked" type="hidden" value="<%if request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="check" or request.cookies(eChsys)("key")="bigmaster" or request.cookies(eChsys)("key")="typemaster" then%>1<%else%><%if request.cookies(eChsys)("key")="selfreg" and fabiaocheck="1" then%>1<%else%><%if request.cookies(eChsys)("key")="smallmaster" and rs5("usecheck")="1" then%>1<%else%>0<%end if%><%end if%><%end if%>"><%rs5.close%>
	<input name="encode" type="hidden" value="html">
	<input name="editor" type="hidden" value="<%=request.cookies(eChsys)("username")%>">
	<input type="button" value=" �� �� " onClick="javascript:history.go(-1)" class="unnamed5" >&nbsp;&nbsp;
	<input type="submit" value=" �� �� " name="Submit" class="unnamed5" >&nbsp; 
	<input type="reset" value=" �� �� " name="Reset" class="unnamed5" >
 <%rs8.close
set rs8=nothing%>
</td>
</tr>
</table>
</td>
</tr>

</form>
</table>
</center>
</div>
<script>
    CKEDITOR.replace('content', {
        height: 350 // ��������������ø߶�
    });
</script>
</body>
</html>
<%end if%>
<!--#include file="Admin_Bottom.asp"-->