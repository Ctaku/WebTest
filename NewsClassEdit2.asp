<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp" -->
<!--#include file="chkuser.asp" -->
<!--#include file="char.inc"-->
<!--#include file="ChkURL.asp"-->
<%dim jingyong
set rs=server.createobject("adodb.recordset")
sql="select * from "& db_User_Table &" where  "& db_User_Name &"='"&request.cookies(eChsys)("username")&"'"
rs.open sql,ConnUser,1,3
jingyong=rs("jingyong")
rs.close
set rs=nothing

if request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="check" or (request.cookies(eChsys)("key")="selfreg" and jingyong=0) or request.cookies(eChsys)("key")="typemaster" then%>
<%
NewsID=Request.QueryString("NewsID")
E_typeid=ChkRequest(Request.Form("E_typeid"),1)	'防注入

set rs2=server.CreateObject ("ADODB.RecordSet")
rs2.Source="select * from "& db_EC_BigClass_Table &" where E_typeid="&E_typeid
rs2.Open rs2.source,conn,1,1

%>
<%
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
document.form1.E_SmallClassID.options[document.form1.E_SmallClassID.length] = new Option("请选择小类(不选为无小类文章)", "");

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
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 修改文章类别</title>
</head>
<body topmargin="0">

<table border="1" width="100%" align=center cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="<%=border%>">
<form name=form1 method="POST" action="NewsClassEdit3.asp?NewsID=<%=NewsID%>">
<tr align="center">
<td colspan="1" height="25" width="100%"  class="TDtop">
	<div align="center">┊ <B>修改文章类别（修改NewsID为<%=NewsID%>的文章的类别）</B> ┊</div>
</td>
</tr>
<input type=hidden name="E_typeid" value=<%=E_typeid%>>
<tr bgcolor="#FFFFFF">
<td align="center">
<p><br><br>
所属大类：
<select name="E_BigClassID" onChange="changelocation(document.form1.E_BigClassID.options[document.form1.E_BigClassID.selectedIndex].value)" size="1">
<option selected value="">请选择大类(不选为无大类文章)</option>
<%dim E_BigClassID
do while not rs2.eof
	E_BigClassID=rs2("E_BigClassID")
%>
<option value="<%=trim(rs2("E_BigClassID"))%>"><%=trim(rs2("E_BigClassName"))%></option>
<%
	rs2.movenext
loop
rs2.close
%>
</select>
所属小类：
<select name="E_SmallClassID" size="1">
<option selected value="">请选择小类(不选为无小类文章)</option>
</select>
</p>
<p>　</p>
</td>
</tr>
<tr bgcolor="#FFFFFF">
<td align="center" height="25" class="TDtop">
<input type="button" value="  返回  " name="B1" onclick=javascript:history.go(-1)  >
<input type="submit" value="  修改  " name="B1"  >

</td>
</tr>
</form>
</table>
</body>
</html>
<%end if%>
<!--#include file="Admin_Bottom.asp"-->