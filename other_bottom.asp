<table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
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
��Ȩ���У�ǭ��������־� <br>
��ַ��<%=Address%> �ʱࣺ<%=Zip%> <br>
�绰��<%=Tel%>&nbsp;����: <a  class="bottom_right"><%=email%></a>&nbsp; <a href="tencent://message/?uin=<%=QQ%>&amp;Site=��վ����ϵͳ&amp;Menu=yes" class="bottom_right"><%=QQ%></a><br />
<%dim endtime
endtime=timer()
response.write "ҳ��ִ��ʱ�䣺"&FormatNumber((endtime-starttime)*1000,3)&"����"%>
&nbsp;<a href="login.asp" target="_blank" class="bottom_right">[��̨����]</a><br>
<br>
</div></td>
  </tr>
</table>
