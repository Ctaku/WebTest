<HTML>
<HEAD>
<TITLE></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<META content="MSHTML 6.00.3790.118" name=GENERATOR>
<LINK href=site.css rel=stylesheet>
<SCRIPT src="inc/exit.js" type=text/javascript></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
function runClock() {
	theTime = window.setTimeout("runClock()", 1000);
	var today = new Date();
	var display= today.toLocaleString();
	status="现在是："+display+"，欢迎您的光临！";
}
</SCRIPT>
<body oncontextmenu="self.event.returnValue=false" onLoad="runClock()">
<table  align="center" border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFDFBF" width="100%" id="AutoNumber1">
<tr>
<td height="25" align="center" background="IMAGES/admin_topbg.gif" class="title">
	版权： <a href="<%=logourl%>" target="_blank"><%=copyright%></a> <%=eChSys%>
	</td>
</tr>
</table>
</BODY>
</HTML>