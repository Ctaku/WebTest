<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkURL.asp"-->

<%
IF request.cookies(eChsys)("KEY")="bigmaster" THEN
	response.redirect "NewsAddd2.asp"
	response.end
else
IF request.cookies(eChsys)("KEY")="typemaster" THEN
	response.redirect "NewsAddd1.asp"
	response.end
else
dim jingyong
set rs=server.createobject("adodb.recordset")
sql="select * from "& db_User_Table &" where  "& db_User_Name &"='"&request.cookies(eChsys)("username")&"'"
rs.open sql,ConnUser,1,3
jingyong=rs("jingyong")
rs.close
set rs=nothing

NewsID=Request.QueryString("NewsID")
E_BigClassID=ChkRequest(Request.Form("E_BigClassID"),1)	'防注入

set rs2=server.CreateObject ("ADODB.RecordSet")
if request.cookies(eChsys)("key")="smallmaster" then
	rs2.Source="select * from "& db_EC_SmallClass_Table &""
else
	rs2.Source="select * from "& db_EC_SmallClass_Table &" where E_BigClassID="&E_BigClassID
end if
rs2.Open rs2.source,conn,1,1

dim rs
dim sql
dim count
set rs=server.createobject("adodb.recordset")
sql = "select * from "& db_EC_SmallClass_Table &" order by smallclassorder"
rs.open sql,conn,1,1
%>
<script language="JavaScript">
var onecount;
onecount=0;
subcat = new Array();
<%
count = 0
do while not rs.eof
%>
subcat[<%=count%>] = new Array("<%= trim(rs("E_smallclassname"))%>","<%= trim(rs("E_BigClassID"))%>","<%= trim(rs("E_SmallClassID"))%>");
<%
	count = count + 1
	rs.movenext
loop
rs.close
%>
onecount=<%=count%>;
function changelocation(locationid)
{
document.form1.E_SmallClassID.length = 0;
document.form1.E_SmallClassID.options[document.form1.E_SmallClassID.length] = new Option("请选择小类", "");
var locationid=locationid;
var i;
for (i=0;i < onecount; i++)
{
if (subcat[i][1] == locationid)
{
document.form1.E_SmallClassID.options[document.form1.E_SmallClassID.length] = new Option(subcat[i][0], subcat[i][2]);
}
}
}
</script>
<script language="JavaScript">
<!--
function na_select_form (fname, type_name) 
{
  document.forms[fname].elements[type_name].select()
  document.forms[fname].elements[type_name].focus()
}
// -->
</script>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 添加文章</title>
</head>
<body topmargin="0">

<table border="1" width="100%" align=center cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="<%=border%>">
<form name=form1 method="post" action="<%if eWebEditor="1" then%>NewsAdd1_eWebEditor.asp<%else%>NewsAdd1.asp<%end if%>">
<tr align="center">
<td colspan="1" class="TDtop" height=25>
	<div align="center" >┊ <B>添加文章--类别选择</B> ┊</div>
</td>
</tr>
<input type=hidden name="E_typeid" value=<%=E_typeid%>>
<tr bgcolor="#FFFFFF">
<td align="center">
<p>　</p>
<p>所属小类：
<select name="E_SmallClassID" size="1">
<option selected value="">请选择小类</option>
<%dim E_SmallClassID,master,smallmaster,smallclassma
do while not rs2.eof
E_SmallClassID=rs2("E_SmallClassID")
master=rs2("smallclassma")
if master<>"" then
smallmaster=split(master,"|")
 for i=0 to ubound(smallmaster)
	if request.cookies(eChsys)("username")=trim(smallmaster(i)) then
	smallclassma=true
	exit for
	else
	smallclassma=false
	end if
 next
end if
if (request.cookies(eChsys)("key")="typemaster") or (request.cookies(eChsys)("key")="bigmaster") or (request.cookies(eChsys)("key")="smallmaster" and smallclassma=true) or request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="check" or request.cookies(eChsys)("key")="selfreg" then
%>
<option value="<%=trim(rs2("E_SmallClassID"))%>"><%=trim(rs2("E_smallclassname"))%></option>
<%
end if
rs2.movenext
loop
rs2.close
%>
</select>

</p>

</td>
</tr>
<tr>
<td align="center" height=25 class="TDtop">
<input type="button" value=" 返 回 " name="B1" onclick=javascript:history.go(-1) >
<input type="submit" value=" 继 续 " name="B1" >
</td>
</tr>
</form>
</table>
</body>
</html>
<!--#include file="Admin_Bottom.asp"-->
<%
end if
end if
%>