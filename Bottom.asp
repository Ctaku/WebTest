<table width="1002" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td  height="31"background="images/index_b_c.jpg" class="bottom"><div align="center">
      <%
dim mymenu1
mymenu1=BOTTOMmenu
Response.Write mymenu1
%>
    </div></td>
  </tr>
  <tr>
    <td height="125" colspan="3" background="Images/index_bottom_bg.gif" class="bottom_right"><div align="center"><br>
版权所有： <a href="" target="_blank"  class="bottom_right"><%=copyright%></a> 黔东南银监分局 <br>
地址：凯里市风情大道13号 邮编：556000 <br>

<!--电话：8659047 -->

&nbsp;admin邮箱: chenhongyu_b@cbrc.gov.cn</a>&nbsp;</a><br />
<%dim endtime
endtime=timer()
response.write "页面执行时间："&FormatNumber((endtime-starttime)*1000,3)&"毫秒"%>
&nbsp;<a href="login.asp" target="_blank" class="bottom_right">[后台管理]</a><br>
<br>
</div></td>
  </tr>
</table>
