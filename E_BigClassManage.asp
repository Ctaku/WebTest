<!--#include file="Conn.ASP"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<%
if request.cookies(eChsys)("KEY")="smallmaster" then
	response.redirect "E_SmallClassManage.asp"
	response.end
else
	if request.cookies(eChsys)("key")<>"bigmaster" then
		dim E_typeid,E_typename
		E_typeid=ChkRequest(request("E_typeid"),1)	'��ע��
	
		if E_typeid="" then
			Show_Err("δָ��������<br><br><a href='javascript:history.back()'>����</a>")
			response.end
		else
			if not IsNumeric(E_typeid) then
				Show_Err("�Ƿ�������<br><br><a href='javascript:history.back()'>����</a>")
				response.end
			end if
		end if
	
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql ="SELECT  * From "& db_EC_Type_Table &" where E_typeid="&E_typeid
		RS.open sql,Conn,3,3
		E_typename=rs("E_typename")
		mode=rs("mode")
		master=rs("typemaster")
		if master<>"" then
			tmaster=split(master,"|")
			for i=0 to ubound(tmaster)
				if request.cookies(eChsys)("username")=trim(tmaster(i)) then
					typemaster=true
					exit for
				else
					typemaster=false
				end if
			next
		end if
		rs.close
		set rs=nothing
	end if
	if typemaster=true or request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="check" or request.cookies(eChsys)("key")="bigmaster" or request.cookies(eChsys)("key")="selfreg" then%>
<html>
<head>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - �������</title>
<script>
function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
</script>
</head>
<LINK href=site.css rel=stylesheet>
</head>
<body topmargin="0">

<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFDFBF" width="100%" id="AutoNumber1">
	<%Set rs6 = Server.CreateObject("ADODB.Recordset")
	if request.cookies(eChsys)("key")<>"bigmaster" then
		sql6 ="SELECT  * From "& db_EC_BigClass_Table &" where E_typeid="&E_typeid&" order by E_bigclassorder"
	else
		sql6 ="SELECT  * From "& db_EC_BigClass_Table &" order by E_bigclassorder"
	end if
	RS6.open sql6,Conn,3,3
	%>
<form method="post" action="E_BigClassSet.asp?action=update">
	<%if request.cookies(eChsys)("key")<>"bigmaster" then%>
<tr class="TDtop">
<td colspan="9"><a href=E_TypeManage.asp >ȫ������</a></td>
</tr>
<tr class="TDtop1">
<td colspan="9">&nbsp;&nbsp;&nbsp;&nbsp;<img src=images/menu1.gif>
<b>
<a href=E_BigClassManage.asp?E_typeid=<%=E_typeid%>><font color=red><%=E_typename%></font></a>
</b>&nbsp;&nbsp;&nbsp;&nbsp;����������
		<%
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql ="SELECT  * From "& db_EC_Type_Table &" where E_typeid<>"&E_typeid&" order by E_typeorder"
		RS.open sql,Conn,3,3
		do while not rs.EOF
			master1=rs("typemaster")
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
			if (typemaster1=true and request.cookies(eChsys)("key")="typemaster") or request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="check" or request.cookies(eChsys)("key")="selfreg" then%>
<a href=E_BigClassManage.asp?E_typeid=<%=rs("E_typeid")%>><%=rs("E_typename")%></a>
			<%end if
			rs.movenext
		loop
		rs.close
		set rs=nothing%>
	</td></tr>
	<%end if%>
<%if mode<>5 then%>
<tr align=center >
<td width="25%"><b>��ѡ�����Ŀ¼</b></td>
<td width="12%">��������</td>
<td width="12%">����ģ��</td>
<td width="9%">������ʾ</td>
<td width="9%">��������</td>
<td width="13%">��������<br>(���������)</td>
<td width="8%">�ⲿ����</td>
<td width="8%">�������Ա<br>(���Զ������|�ָ�)</td>
<td width="5%">ɾ��</td>
</tr>
	<%	
	do while not rs6.eof
		dim bigclassmaster,bigmaster,master
		master=rs6("bigclassmaster")
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
	
		set rstype=createobject("adodb.recordset")
		sql="select * from "& db_EC_Type_Table &" order by E_typeorder"
		rstype.Open sql,conn,1,3
		if bigclassmaster=true or request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="check" or (typemaster=true and request.cookies(eChsys)("key")="typemaster") or request.cookies(eChsys)("key")="selfreg" then
			%>
<tr >
<td>
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src=images/menu1.gif>
	<a href="E_SmallClassManage.asp?E_BigClassID=<%=rs6("E_BigClassID")%>" ><%=rs6("E_BigClassName")%></a></td>
<td align=center>
	<input type="hidden" name="E_BigClassID" value="<%=rs6("E_BigClassID")%>"  >
	<input class=text type="text" name="E_BigClassName" size="10" value="<%=rs6("E_BigClassName")%>" <%if not(request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="typemaster") then%>disabled<%end if%>  >
</td>

<!--
<td align=center>
	<input class=text type="text" name="bigclasszs" size="10" value="<%=rs6("bigclasszs")%>" <%if not(request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="typemaster") then%>disabled<%end if%>  >
</td>
-->

<td align="center" bgcolor="#FFFFFF">
	<select size="1" name="bigclasszs" style="font-family: ����; font-size: 9pt" <%if not(request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="typemaster") then%>disabled<%end if%>>
	  <option <%if rs6("bigclasszs")="1" then%>selected<%end if%> value="1">ͼƬ��Ʒģ��</option>
	  <option <%if rs6("bigclasszs")="2" then%>selected<%end if%> value="2">����ý��ģ��</option>
	<option <%if rs6("bigclasszs")="3" then%>selected<%end if%> value="3">��ַ�Ƽ�ģ��</option>
	<option <%if rs6("bigclasszs")="4" then%>selected<%end if%> value="4">�������ģ��</option>
	<option <%if rs6("bigclasszs")="5" then%>selected<%end if%> value="5">����ҳ��ģ��</option>
	</select>
</td>

<td align=center>
	<select size="1" name="E_BigClassView" style="font-family: ����; font-size: 9pt" <%if not(request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="typemaster") then%>disabled<%end if%>>
	<option <%if rs6("E_BigClassView")="1" then%> selected <%end if%>value="1">��ʾ</option>
	<option <%if rs6("E_BigClassView")="0" then%> selected <%end if%>value="0">����ʾ</option>                             
	</select>
</td>
<td align=center>
	<select name="E_typeid" style="font-family: ����; font-size: 9pt" <%if not(request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="typemaster") then%>disabled<%end if%>>
			<%if request.cookies(eChsys)("key")<>"bigmaster" then
				do while not rstype.EOF
					master2=rstype("typemaster")
					if master2<>"" then
						tmaster2=split(master2,"|")
						for i=0 to ubound(tmaster2)
							if request.cookies(eChsys)("username")=trim(tmaster2(i)) then
								typemaster2=true
								exit for
							else
								typemaster2=false
							end if
						next
					end if
					if (typemaster2=true and request.cookies(eChsys)("key")="typemaster") or request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="check" or request.cookies(eChsys)("key")="selfreg" then%>
	<option value="<%=rstype("E_typeid")%>" <%if rs6("E_typeid")=cint(rstype("E_typeid")) then%>selected<%end if%>><%=rstype("E_typename")%></option>
					<%end if
					rstype.MoveNext
				loop
			else
				Set rs = Server.CreateObject("ADODB.Recordset")
				sql ="SELECT  * From "& db_EC_Type_Table &" where E_typeid="&rs6("E_typeid")
				RS.open sql,Conn,3,3%>
	<option value="<%=rs6("E_typeid")%>"><%=rs("E_typename")%></option>
				<%
				rs.close
				set rs=nothing
			end if%>
	</select>
</td>
<td align=center><input class=text type="text" name="E_bigclassorder" size="10" value="<%=rs6("E_bigclassorder")%>" ONKEYPRESS="event.returnValue=IsDigit();" <%if not(request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="typemaster") then%>disabled<%end if%>  ></td>
<td align=center><input class=text type="text" name="url" size="14" value="<%=rs6("url")%>"  <%if not(request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="typemaster") then%>disabled<%end if%>  ></td>
<td align=center><input class=text type="text" name="bigclassmaster" size="8" value="<%=rs6("bigclassmaster")%>" <%if not(request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="typemaster") then%>disabled<%end if%>  ></td>
<td align=center>
			<%if request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="typemaster" then%>
	<a href="E_BigClassKill.asp?E_BigClassID=<%=rs6("E_BigClassID")%>">ɾ��</a>
			<%end if%>
</td>
</tr>
		<%end if
		RS6.MoveNext
	Loop
	set rstype=nothing
	rs6.close
	set rs6=nothing
	%>
<tr>
<td height="25" colspan="1" align="center" >
<a href="javascript:window.location.reload()" target=mainmenu title="�����Ŀ��������������˵���" style="font-family: ����; font-size: 9pt">ˢ������</a>
</td> 
<td colspan="8"  align="center" >
	<input type="submit" name="Submit2" value="����" style="font-family: ����; font-size: 9pt" <%if not(request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="typemaster") then%>disabled<%end if%>  >&nbsp;
	<!--
	<input type="button" value="�������" style="font-family: ����; font-size: 9pt" onclick="javascript:window.open('NewsAddd1.asp','_self','')" <%if not(request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="typemaster" or request.cookies(eChsys)("key")="check" or request.cookies(eChsys)("key")="selfreg") then%>disabled<%end if%>  >
	-->
</td>
</tr>
</form>
	<%if request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="typemaster" then%>
<form method="post" action="E_BigClassSet.asp?action=add">
<tr >
<td align="center">��Ӵ���</td>
<td align=center><input class=text type="text" name="E_BigClassName" size="10"  ></td>

<td align="center" bgcolor="#FFFFFF">
	  <select size="1" name="bigclasszs" style="font-family: ����; font-size: 9pt" <%if not(request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="typemaster") then%>disabled<%end if%>>
	  <option value="2">����ý��ģ��</option>
	  <option value="1">ͼƬ��Ʒģ��</option>
	  <option value="3">��ַ�Ƽ�ģ��</option>
	  <option value="4">�������ģ��</option>
	  <option value="5">����ҳ��ģ��</option>
	  </select>
</td>

<td align=center>
	<select size="1" name="E_BigClassView" style="font-family: ����; font-size: 9pt">
	<option selected value="1">��ʾ</option>
	<option value="0">����ʾ</option>                             
	</select>
</td>
<td align=center>
		<%set rstype=createobject("adodb.recordset")
		sql="select * from "& db_EC_Type_Table &" order by E_typeorder"
		rstype.Open sql,conn,1,3%>
	<select name="E_typeid" style="font-family: ����; font-size: 9pt">
		<%do while not rstype.EOF
			master3=rstype("typemaster")
			if master3<>"" then
				tmaster3=split(master3,"|")
				for i=0 to ubound(tmaster3)
					if request.cookies(eChsys)("username")=trim(tmaster3(i)) then
						typemaster3=true
						exit for
					else
						typemaster3=false
					end if
				next
			end if
			if request.cookies(eChsys)("key")="super" or (typemaster3=true and request.cookies(eChsys)("key")="typemaster") then%>
	<option <%if E_typeid=trim(rstype("E_typeid")) then%>selected<%end if%> value="<%=rstype("E_typeid")%>"><%=rstype("E_typename")%></option>
			<%end if
			rstype.MoveNext
		loop
		set rstype=nothing
		%>
	</select>
</td>
<td align=center><input class=text type="text" name="E_bigclassorder" size="10" ONKEYPRESS="event.returnValue=IsDigit();"  ></td>
<td align=center><input class=text type="text" name="url" size="14" value="0" ></td>
<td align=center><input class=text type="text" name="bigclassmaster" size="8"  ></td>
<td align=center><input type="submit" name="Submit2" value="���" style="font-family: ����; font-size: 9pt"  ></td>
</tr>
</form>
	<%
	set rstype=nothing
end if
end if
%>
<%
if request.cookies(eChsys)("key")="check" or request.cookies(eChsys)("key")="typemaster" or request.cookies(eChsys)("key")="super" then
	%>
<tr >
<td colspan="8" align=center>
	<input type=button value="�����޴�������" class=button onClick="javascript:window.open('newsadd1.asp?E_typeid=<%=E_typeid%>','_self','')"  >
	<input type=button value="�༭�޴�������" class=button onClick="javascript:window.open('E_bigclassNO.asp?E_typeid=<%=E_typeid%>','_self','')"  >
</td>
</tr>
	<%end if%>
</table>
</body>
</html>
<!--#include file="Admin_Bottom.asp"-->
	<%else
		response.write "<script>alert('�Բ�������Ȩ������');history.go(-1);</Script>"
		Response.End 
	end if
end if
conn.close
set conn=nothing%>