<!--#include file="Conn.ASP"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.ASP"-->
<!--'#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<%
if request.cookies(eChsys)("key")<>"smallmaster" then
	dim E_BigClassID,E_BigClassName,E_typename,E_typeid
	E_BigClassID=ChkRequest(request("E_BigClassID"),1)	'防注入
             
	if E_BigClassID="" then
		Show_Err("未指定参数！<br><br><a href='javascript:history.back()'>返回</a>")
		response.end
	else
		if not IsNumeric(E_BigClassID) then
			Show_Err("非法参数！<br><br><a href='javascript:history.back()'>返回</a>")
			response.end
		end if
	end if

	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="SELECT * From "& db_EC_BigClass_Table &" where E_BigClassID="&E_BigClassID
	RS.open sql,Conn,3,3
	E_BigClassName=rs("E_BigClassName")
	E_typeid=rs("E_typeid")
	master5=rs("bigclassmaster")
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
	
	Set rs1 = Server.CreateObject("ADODB.Recordset")
	sql1="SELECT * From "& db_EC_Type_Table &" where E_typeid="&E_typeid
	RS1.open sql1,Conn,3,3
	E_typename=rs1("E_typename")
	master4=rs1("typemaster")
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
	rs1.close
	set rs1=nothing
	rs.close
set rs=nothing
end if
if typemaster4=true or request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="check" or bigclassmaster5=true or request.cookies(eChsys)("key")="smallmaster" or request.cookies(eChsys)("key")="selfreg" then
	%>

<html>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 文章管理</title>
<script>
function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
</script>
</head>
<body topmargin="0">


<center>
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFDFBF" width="100%" id="AutoNumber1">
<form method="post" action="E_SmallClassSet.asp?action=update">
<%if request.cookies(eChsys)("key")<>"smallmaster" then%>
<tr class="TDtop">
<td colspan="9">
<a href=E_TypeManage.asp >全部总栏</a>
</td>
</tr>
<tr >
<td colspan="9"class="TDtop1">
&nbsp;&nbsp;&nbsp;&nbsp;<img src=images/menu1.gif>
<b>
<a href="E_BigClassManage.asp?E_typeid=<%=E_typeid%>"><font color=red><%=E_typename%></font></a>
</b>
&nbsp;&nbsp;&nbsp;&nbsp;其他总栏：
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
		if (request.cookies(eChsys)("key")="typemaster" and typemaster1=true) or request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="check" or request.cookies(eChsys)("key")="selfreg" then%>
<a href=E_BigClassManage.asp?E_typeid=<%=rs("E_typeid")%>><%=rs("E_typename")%></a>
		<%end if
		rs.movenext
	loop
	rs.close
	set rs=nothing%>
</td>
</tr>
<tr >
<td colspan="9" class="TDtop1">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src=images/menu1.gif>
<b>
<a href="E_SmallClassManage.asp?E_BigClassID=<%=E_BigClassID%>"><font color=red><%=E_BigClassName%></font></a>
</b>
&nbsp;&nbsp;&nbsp;&nbsp;其他大类：
	<%if request.cookies(eChsys)("key")<>"bigmaster" then
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql ="SELECT  * From "& db_EC_BigClass_Table &" where E_typeid="&E_typeid&" and E_BigClassID<>"&E_BigClassID&" order by E_bigclassorder"
		RS.open sql,Conn,3,3
		do while not rs.EOF%>
<a href=E_SmallClassManage.asp?E_BigClassID=<%=rs("E_BigClassID")%>><%=rs("E_BigClassName")%></a>
			<%rs.movenext
		loop
		rs.close
		set rs=nothing
	else
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql ="SELECT  * From "& db_EC_BigClass_Table &" where E_BigClassID<>"&E_BigClassID&" order by E_bigclassorder"
		RS.open sql,Conn,3,3
		do while not rs.EOF
			master=rs("bigclassmaster")
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
			if (bigclassmaster=true and request.cookies(eChsys)("key")="bigmaster") then%>
<a href=E_SmallClassManage.asp?E_BigClassID=<%=rs("E_BigClassID")%>><%=rs("E_BigClassName")%></a>
			<%end if
			rs.movenext
		loop
		rs.close
		set rs=nothing
	end if%>
</td>
</tr>
<%end if
Set rs6 = Server.CreateObject("ADODB.Recordset")
if request.cookies(eChsys)("key")<>"smallmaster" then
	sql6 ="SELECT * From "& db_EC_SmallClass_Table &" where E_BigClassID="&E_BigClassID&" order by SmallClassorder"
else
	sql6 ="SELECT * From "& db_EC_SmallClass_Table &" order by SmallClassorder"
end if
RS6.open sql6,Conn,3,3
%>
<%if bigclasszs<>5 then%>
<tr align=center >
<td width="25%"><b>请选择小类目录</b></td>
<td width="12%">小类名称</td>
<td width="12%">小类注释</td>
<td width="9%">小类显示</td>
<td width="9%">属于大类</td>
<td width="12%">小类排序<br>(必填项，数字)</td>
<td width="8%">外部链接</td>
<td width="8%">小类管理员<br>(可以多个，用|分隔)</td>
<td width="5%">删除</td>
</tr>
<%
do while not rs6.eof
	if request.cookies(eChsys)("key")="smallmaster" then
		dim smallclassmaster,smallmaster,master2
		master2=rs6("smallclassma")
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
	end if
	set rstype=createobject("adodb.recordset")
	if request.cookies(eChsys)("key")<>"smallmaster" then
		sql="select * from "& db_EC_BigClass_Table &" where E_typeid="&E_typeid&" order by E_bigclassorder "
	else
		sql="select * from "& db_EC_BigClass_Table &" where E_BigClassID="&rs6("E_BigClassID")&" order by E_bigclassorder "
	end if
	rstype.Open sql,conn,1,3
	if request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="check" or (request.cookies(eChsys)("key")="typemaster" and typemaster4=true) or (request.cookies(eChsys)("key")="bigmaster") or (smallclassmaster=true and request.cookies(eChsys)("key")="smallmaster") or request.cookies(eChsys)("key")="selfreg" then%>
<tr>
<td>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src=images/menu1.gif>
<a href="listnews.asp?E_SmallClassID=<%=rs6("E_SmallClassID")%>" title="<%if rs6("smallClasszs")<>"" then%><%=rs6("smallClasszs")%><%else%>无<%end if%>"><%=rs6("E_smallclassname")%></a></td>
<td align=center>
	<input type="hidden" name="E_SmallClassID" value="<%=rs6("E_SmallClassID")%>">
	<input class=text type="text" name="E_smallclassname" size="10" value="<%=rs6("E_smallclassname")%>" <%if not(request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="typemaster" or request.cookies(eChsys)("key")="bigmaster") then%>disabled<%end if%>  >
</td>
<td align=center>
	<input class=text type="text" name="smallclasszs" size="10" value="<%=rs6("smallclasszs")%>" <%if not(request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="typemaster" or request.cookies(eChsys)("key")="bigmaster") then%>disabled<%end if%>  >
</td>
<td align=center>
	<select size="1" name="E_smallclassview" style="font-family: 宋体; font-size: 9pt" <%if not(request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="typemaster" or request.cookies(eChsys)("key")="bigmaster") then%>disabled<%end if%>>
	<option <%if rs6("E_smallclassview")="1" then%> selected <%end if%> value="1">显示</option>
	<option <%if rs6("E_smallclassview")="0" then%> <%end if%> value="0">不显示</option>                             
    </select>
</td>
<td align=center>
	<select name="E_BigClassID" style="font-family: 宋体; font-size: 9pt" <%if not(request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="typemaster" or request.cookies(eChsys)("key")="bigmaster") then%>disabled<%end if%>>
		<%if (request.cookies(eChsys)("key")="typemaster" and typemaster1=true) or request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="selfreg" then
			do while not rstype.EOF%>
	<option value="<%=rstype("E_BigClassID")%>" <%if rs6("E_BigClassID")=cint(rstype("E_BigClassID")) then%>selected<%end if%>><%=rstype("E_BigClassName")%></option>
				<%
				rstype.MoveNext
			loop
		end if
		if request.cookies(eChsys)("key")="bigmaster" then%>
	<option value="<%=E_BigClassID%>"><%=E_BigClassName%></option>
		<%end if
		if request.cookies(eChsys)("key")="smallmaster" and smallclassmaster=true then%>
	<option value="<%=rstype("E_BigClassID")%>"><%=rstype("E_BigClassName")%></option>
		<%end if%>
	</select>
</td>
<td align=center><input class=text type="text" name="smallclassorder" size="10" value="<%=rs6("smallclassorder")%>" ONKEYPRESS="event.returnValue=IsDigit();" <%if not(request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="typemaster" or request.cookies(eChsys)("key")="bigmaster") then%>disabled<%end if%>  ></td>
<td align=center><input class=text type="text" name="url" size="14" value="<%=rs6("url")%>"  <%if not(request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="typemaster" or request.cookies(eChsys)("key")="bigmaster") then%>disabled<%end if%> ></td>
<td align=center><input class=text type="text" name="smallclassma" size="8" value="<%=rs6("smallclassma")%>" <%if not(request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="typemaster" or request.cookies(eChsys)("key")="bigmaster") then%>disabled<%end if%>  ></td>
<td align=center><%if request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="typemaster" or request.cookies(eChsys)("key")="bigmaster" then%><a href="E_SmallClassKill.asp?E_SmallClassID=<%=rs6("E_SmallClassID")%>">删除</a><%end if%></td>
</tr>
	<%end if
	RS6.MoveNext
Loop
set rstype=nothing
rs6.close
set rs6=nothing
%>
<tr> 
<td colspan="1" height="25" align="center" bgcolor="#FFFFFF">
	<a href="javascript:window.location.reload()" target=mainmenu title="添加栏目类别后请更新左栏菜单项" style="font-family: 宋体; font-size: 9pt">刷新左栏</a>
</td>
<td colspan="8"  align="center" bgcolor="#FFFFFF">
	<input type="submit" name="Submit2" value="保存" style="font-family: 宋体; font-size: 9pt" <%if not(request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="typemaster" or request.cookies(eChsys)("key")="bigmaster") then%>disabled<%end if%>  >
</td>
</tr>
</form>
<%if (request.cookies(eChsys)("key")="bigmaster") or request.cookies(eChsys)("key")="super" or (request.cookies(eChsys)("key")="typemaster" and typemaster1=true) then%>
<form method="post" action="E_SmallClassSet.asp?action=add">
<tr >
<td align="center">
添加小类
</td>
<td align=center><input class=text type="text" name="E_smallclassname" size="10"  ></td>
<td align=center><input class=text type="text" name="smallclasszs" size="10"  ></td>
<td align=center>
	<select size="1" name="E_smallclassview" style="font-family: 宋体; font-size: 9pt">
	<option selected value="1">显示</option>
	<option value="0">不显示</option>                             
	</select></td>
<td align=center>
	<% 
	set rstype=createobject("adodb.recordset")
	if request.cookies(eChsys)("key")="typemaster" or request.cookies(eChsys)("key")="super" then
		sql="select * from "& db_EC_BigClass_Table &" where E_typeid="&E_typeid&" order by E_bigclassorder "
	else
		sql="select * from "& db_EC_BigClass_Table &" order by E_bigclassorder "
	end if
	rstype.Open sql,conn,1,3%>
	<select name="E_BigClassID" style="font-family: 宋体; font-size: 9pt">
	<%if request.cookies(eChsys)("key")<>"bigmaster" then
		do while not rstype.EOF%>
	<option <%if E_BigClassID=trim(rstype("E_BigClassID")) then%>selected<%end if%> value="<%=rstype("E_BigClassID")%>"><%=rstype("E_BigClassName")%></option>
			<%rstype.MoveNext
		loop
	else%>
	<option value="<%=E_BigClassID%>"><%=E_BigClassName%></option>
	<%end if%>
	</select>
</td>
<td align=center>
	<input class=text type="text" name="smallclassorder" size="10" ONKEYPRESS="event.returnValue=IsDigit();"  >
</td>
<td align=center><input class=text type="text" name="url" size="14" value="0" ></td>
<td align=center>
	<input class=text type="text" name="smallclassma" size="8"  >
</td>
<td align=center>
	<input type="hidden" name="E_typeid" value="<%=E_typeid%>"  >
	<input type="submit" name="Submit2" value="添加" style="font-family: 宋体; font-size: 9pt"  >
</td>
</tr>
</form>
	<%
	set rstype=nothing
end if 
end if%><%
if request.cookies(eChsys)("key")<>"smallmaster" then
	%>
<tr >
<td colspan="9" align=center>
	<input type=button value="发表无小类文章" class=button onClick="javascript:window.open('newsadd1.asp?E_BigClassID=<%=E_BigClassID%>','_self','')"  >
	<input type=button value="编辑无小类文章" class=button onClick="javascript:window.open('E_SmallNO.asp?E_BigClassID=<%=E_BigClassID%>','_self','')"  >
</td>
</tr>
<%end if%>
</table>
</center>
</body>
</html>
<!--#include file="Admin_Bottom.asp"-->
<%else
	Show_Err("对不起，你无权操作！<br><br><a href='javascript:history.back()'>返回</a>")
	Response.End 
end if
%>