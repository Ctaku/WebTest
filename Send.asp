<%@ Language=VBScript%>
<!--#include file="conn.asp" -->
<!--#include file="config.asp" -->
<%
NewsID=Request.QueryString("NewsID")

if NewsID="" then
Response.Write "<script>alert('未指定参数');history.back()</script>"
response.end
else
if not IsNumeric(NewsID) then
 response.write "<script>alert('非法参数');history.back()</script>"
 response.end
else

reffer=Request.ServerVariables("HTTP_REFERER")

end if 
end if
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>文章评论_<%=eChSys%></title>
<LINK href=news.css rel=stylesheet>
</head>

<body topmargin="0"><form method="POST" action="Send_submit.asp"><input type=hidden name=NewsID value=<%=NewsID%>><input type=hidden name=reffer value=<%=reffer%>>
  <!--#include file="other_top.asp" -->
  <table width="780" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
    <tr valign="top"> 
      <td height="356"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="100%" height="25" background="IMAGES/dh_bg.gif" class="daohang" >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;栏目导航<a class="daohang" href="./" >&nbsp;</a>&nbsp;&nbsp;当前位置：<a class="daohang" href="./" >网站首页</a>&gt;&gt;发送邮件</td>
          </tr>
          <tr bgcolor="CDCCCC"> 
            <td height="48" align="center" background="IMAGES/BookBG.gif" bgcolor="#FFFFFF"><a href="guestadd.asp" class=class>我要留言</a> 
              | <a class=class href="javascript:window.location.reload()">刷新</a>            </td>
          </tr>
          <tr bgcolor="CDCCCC"> 
            <td height="25" align="center" bgcolor="#FFFFFF"><table width="95%"  border="0" align="center" cellpadding="0" cellspacing="0" id="AutoNumber3" style="border-collapse: collapse">
                <tr> 
                  <td valign=top>&nbsp;</td>
                </tr>
                <tr> 
                  <td valign=top><table border="0" cellpadding="3" cellspacing="0" width="100%">
                      <tr> 
                        <td width="100%" height="50"><font color="000099">请输入你朋友的名字： 
                          <input type="text" name="name" size="50">
                          </font></td>
                      </tr>
                      <tr> 
                        <td width="100%" height="50"><font color="000099">请输入你朋友的MAIL：</font><font color="000099"> 
                          <input type="text" name="email" size="50">
                          </font></td>
                      </tr>
                      <tr> 
                        <td width="100%" height="50"><font color="000099">请输入你自己的名字： 
                          <input type="text" name="myname" size="50">
                          </font></td>
                      </tr>
                      <tr> 
                        <td width="100%" align="center" height="50"> <input type="submit" value="  发送  " name="B1"></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td valign=top>&nbsp;</td>
                </tr>
            </table></td>
          </tr>
          <tr bgcolor="CDCCCC"> 
            <td height="19" align="center" bgcolor="#FFFFFF">&nbsp;</td>
          </tr>
        </table>        
      </td>
    </tr>
  </table>
<!--#include file="other_bottom.asp" -->
</form>
</body>

</html>