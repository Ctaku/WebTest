<table width="1002" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#EFF5FA">
<tr><td height="5"></td></tr>
            <tr>
              <%if showlinkmap=1 then%>
              <td align="center" height="49"><%if linkshownum >8 then
		Response.Write "<script language=JavaScript>marquee_logo_news();</script><P align=left>"& vbcrlf
	end if
	set rs10=server.CreateObject("ADODB.RecordSet") 
	rs10.Source="select top "& linkshownum &" * from "& db_EC_Link_Table &" where linktype=2 and pass=1 order by ID DESC "
	rs10.Open rs10.Source,conn,1,1
	for i=1 to linkshownum
		if not rs10.EOF then%>
                  <a href="<%=rs10("weburl")%>" target="_blank" title="<%=rs10("webname")%>&#13;&#10;简介：<%=rs10("content")%>&#13;&#10;站长：<%=rs10("webmaster")%>&#13;&#10;申请时间：<%=rs10("dateandtime")%>"><img src="<%=rs10("logo")%>" width="88" height="31" border="0" align=left></a>
                  <%rs10.MoveNext
		else%>
                  <a href="#" onClick="javascript:linkreg()"><img src="images/nologo.gif" width=88 height=31 border=0 align=left></a>
<%end if
Next%>              </td>
 <%
	rs10.Close
	set rs10=nothing
end if%>
            </tr>
</table>
<table width="1002" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#EFF5FA" >
              <tr>
                <%if showlink=1 then
set rs10=server.CreateObject("ADODB.RecordSet") 
rs10.Source="select top "& linkshownum &" * from "& db_EC_Link_Table &" where linktype=0 and pass=1 order by ID DESC "
rs10.Open rs10.Source,conn,1,1
for i=1 to 6
	if not rs10.EOF then
	%>
                <td height=21 align="center"><a class="middle" href="<%=rs10("weburl")%>" target="_blank" title="<%=rs10("content")%>
站长：<%=rs10("webmaster")%>
申请时间：<%=rs10("dateandtime")%>"><%=rs10("webname")%></a></td>
                <%
rs10.MoveNext
else
%>
                <td height=21 align="center"><a class="middle" href="#" onClick="javascript:linkreg()">您的位置</a></td>
 <%end if
	Next
rs10.Close
set rs10=nothing
%>
                <td align="center"><input type=button style="cursor:hand" class="login_username" onMouseOver ="this.style.backgroundColor='#F1F7FC'" onMouseOut ="this.style.backgroundColor='#FFFFFF'" name=link1 value="申请" onClick="javascript:linkreg()">
                    <input type=button style="cursor:hand" class="login_username" onMouseOver ="this.style.backgroundColor='#F1F7FC'" onMouseOut ="this.style.backgroundColor='#FFFFFF'" name=link1 value="更多" onClick="javascript:morelink()">                </td>
              </tr>
 <%end if%>
</table>
  <table width="1002" border="0" align="center" cellpadding="0" cellspacing="4" bgcolor="#EFF5FA">
    <tr>
      <td width="141"><div align="center">
          <select style="WIDTH: 90%" 
                  onchange="window.open(this[this.selectedIndex].value,'','')" 
                  name="select4">
            <option selected="selected">-银监会网站-</option>
            <option value="http://10.1.3.60/">银监会</option>
          
          </select>
      </div></td>
      <td width="142"><div align="center">
          <select style="WIDTH: 90%" 
                  onchange="window.open(this[this.selectedIndex].value,'','')" 
                  name="select3">
            <option selected="selected">-银监局网站-</option>
            <option value="http://www.beijing.gov.cn">北京银监局</option>
            
          </select>
      </div></td>
      <td width="140"><div align="center">
          <select style="WIDTH: 90%" 
                  onchange="window.open(this[this.selectedIndex].value,'','')" 
                  name="select">
            <option>-银监分局-</option>
            <option value="http://www.22.com">遵义</option>
          </select>
      </div></td>
      <td width="142"><div align="center">
          <select style="WIDTH: 90%" 
                  onchange="window.open(this[this.selectedIndex].value,'','')" 
                  name="select5">
            <option selected="selected">-其他-</option>
            <option value="http://www.xxxxxxxxx.com">其他</option>
          </select>
      </div></td>
      <td width="141"><select style="WIDTH: 90%" 
                  onchange="window.open(this[this.selectedIndex].value,'','')" 
                  name="select2">
          <option selected="selected">-其他-</option>
         
          <option value="http://www.bjradio.com.cn">广播电台</option>
      </select></td>
    </tr>
</table>
