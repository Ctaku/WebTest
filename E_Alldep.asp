<%@ Language=VBScript%>
<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; chars8et=gb2312">
<LINK href=news.css rel=stylesheet>
<title>单位部门稿件排行__<%=eChSys%></title>
</head>
<body topmargin="0">
<!--#include file="other_Top.asp"-->
<table width="780" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#F1F7FC">
        <tr> 
          <td height="25" background="Images/dh_bg.gif" class="daohang">&nbsp;&nbsp;&nbsp;&nbsp;<img src="Images/arrow_dh_1.gif" width="7" height="7">&nbsp;&nbsp;&nbsp;栏目导航&nbsp;&nbsp;<span class="daohang">当前位置</span>：<a class="daohang" href="./" >网站首页</a>&gt;&gt;单位部门稿件排行</td>
        </tr>
        <tr> 
          <td>
		  <table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">

		<%set rs8=server.CreateObject("ADODB.RecordSet")
		sql8="select * from "& db_EC_Dep_Table &" where depnumber>=0 order by depnumber desc"
		rs8.open sql8,Conn,1,1
		if not rs8.EOF then
		
		%>

		<tr align="center" class="TDtop1"> 
			<td width="10%" height="25" >排名</td>
			<td width="40%" >单位名称</td>
			<td width="25%" >所属部门</td>
			<td width="25%" >累计发稿</td>
		</tr>
		<% 
		i=1
		while not rs8.EOF 
		%>
		<tr> 
			<td width="10%" align="center" bgcolor="#FFFFFF"><%=i%></td>
			<td width="40%" align="center" bgcolor="#FFFFFF"><%=rs8("E_DepName")%></td>
			<td width="25%" height="25" align="center" bgcolor="#FFFFFF"><%=rs8("E_DepType")%></td>
			<td width="25%" bgcolor="#FFFFFF" align="center"><%=rs8("depnumber")%></td>
		</tr>

		<%rs8.MoveNext
		i=i+1
		wend
		end if
		rs8.close
		set rs8=nothing
		%>

</table>
</td>
</tr>
</table>
</body>
</html>
<!--#include file="other_bottom.asp"-->