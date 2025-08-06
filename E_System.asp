<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="inc/config.asp"-->
<!--'#include file="char.inc"-->
<!--#include file="Function.asp"-->
<!--#include file="ChkUser.asp" -->
<!--#include file="ChkManage.asp" -->
<SCRIPT event=onclick for=Ok language=JavaScript>
  window.returnValue = moveurl.value
  window.close();
</SCRIPT>
<script>
function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
</script>
<%
IF request.cookies(eChsys)("ManageKEY")<>"super" then
	Show_Err("对不起，您的后台管理权限不够！")
	response.end
else
	response.buffer=true
	Response.Expires=0

	Set rs9 = Server.CreateObject("ADODB.Recordset")
	sql9 ="SELECT * From "& db_EC_System_Table &" Order By id DESC"
	RS9.open sql9,Conn,1,1
	if rs9("system")="1" or request.cookies(eChsys)("ManagePurview")="99999" then
		%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 系统设置</title>
<style type="text/css">
<!--
.style1 {color: #FF0000}
.STYLE2 {color: #FF6600}
-->
</style>
<SCRIPT language="JavaScript" type="text/javascript">
// Begin color
	function color(color)
	{
	url = 'color.htm';
	window.open(url,color,"width=430,height=440,status=no,toolbar=no,menubar=no,resizable=no");
	}
// End color-->

function getServeUrl() 
	{
	document.system.xpurl.value = "<%=ServerHttpUrl("E_System.asp")%>";
	}
</script>
</head>
<body topmargin="0">

<div align="center">
<table width="100%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolor="#FFDFBF" bgcolor="#ffffff" id="AutoNumber1" style="border-collapse: collapse">
<form method="post" action="E_SystemSave.asp" name="system">
<tr> 
	<td colspan="3" height=25 class="title"  background="images/admin_topbg.gif"> 
		<div align="center">| 网站属性设置 |</div>	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" ><font color=red>*</font>网站名称：</td>
	<td  colspan="2" align="left" >
		<input type="text" name="SiteName" size="50" value="<%=rs9("SiteName")%>" >	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" ><font color=red>*</font>网站地址：</td>
	<td  colspan="2" align="left">
		<input type="text" name="xpurl" size="40" value="<%=rs9("xpurl")%>" > 
		<input type="button" name="geturl" value="当前获取"  onclick="javascript:getServeUrl()" > <font color=red>非常重要</font>(系统安装所在http地址)	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" ><font color=red>*</font>网管信箱：</td>
	<td  colspan="2" align="left" >
		<input type="text" name="email" size="50" value="<%=rs9("email")%>" >	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" >网管QQ：</td>
	<td  colspan="2" align="left" >
		<input type="text" name="QQ" size="50" value="<%=rs9("QQ")%>" >	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" >版权声明：</td>
	<td  colspan="2" align="left" > 
	<input type="text" name="Copyright" size="50" value="<%=rs9("Copyright")%>" ></td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" >单位地址：</td>
	<td  colspan="2" align="left" > 
	<input type="text" name="Address" size="50" value="<%=rs9("Address")%>" ></td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" >邮政编码：</td>
	<td  colspan="2" align="left" > 
	<input type="text" name="Zip" size="50" value="<%=rs9("Zip")%>" ></td>
</tr>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" >联系电话：</td>
	<td  colspan="2" align="left" > 
	<input type="text" name="Tel" size="50" value="<%=rs9("Tel")%>" ></td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" >版本信息：</td>
	<td  colspan="2" align="left" > <font color="#000000"><b>[<%=rs9("version")%>&nbsp;<%=rs9("ver")%>] </b></font></td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" height="18" align="right" bgcolor="#FFFFFF">导航菜单分级：</td>
	<td colspan="2" align="left">
		<select name="B_BG" size="1" id="gd1" >
		<option <%if rs9("B_BG")="1" then%>selected<%end if%> value="1">一级</option>
		<option <%if rs9("B_BG")="0" then%>selected<%end if%> value="0">二级</option>
		<option <%if rs9("B_BG")="3" then%>selected<%end if%> value="3">三级</option>
		</select>	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" height="18" align="right" bgcolor="#FFFFFF">LOGO图标：</td>
	<td colspan="2" align="left">
		<input name="logo" type="text" id="logo"  value="<%=rs9("logo")%>" size="50">	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" height="18" align="right" bgcolor="#FFFFFF">LOGO图标URL：</td>
	<td colspan="2" align="left">
		<input name="logourl" type="text" id="logourl"  value="<%=rs9("logourl")%>" size="50">	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" bgcolor="#FFFFFF">LOGO类型：</td>
	<td colspan="2" align="left">
		<select name="gd1" size="1" id="gd1" >
		<option <%if rs9("gd1")<>"0" then%>selected<%end if%> value="1">photo</option>
		<option <%if rs9("gd1")="0" then%>selected<%end if%> value="0">flash</option>
		</select>	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" height="18" align="right" bgcolor="#FFFFFF">首页顶部Banner条：</td>
	<td colspan="2" align="left">
		<input name="TopBanner" type="text" id="TopBanner"  value="<%=rs9("TopBanner")%>" size="50">	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" height="18" align="right" bgcolor="#FFFFFF">首页顶部Banner条URL：</td>
	<td colspan="2" align="left">
		<input name="TopBannerurl" type="text" id="TopBannerurl"  value="<%=rs9("TopBannerurl")%>" size="50">	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" height="18" align="right" bgcolor="#FFFFFF">子页面顶部Banner条：</td>
	<td colspan="2" align="left">
		<input name="OtherTopBanner" type="text" id="OtherTopBanner"  value="<%=rs9("OtherTopBanner")%>" size="50">	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" height="18" align="right" bgcolor="#FFFFFF">子页面顶部Banner条URL：</td>
	<td colspan="2" align="left">
		<input name="OtherTopBannerurl" type="text" id="OtherTopBannerurl"  value="<%=rs9("OtherTopBannerurl")%>" size="50">	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" bgcolor="#FFFFFF">Banner条类型：</td>
	<td colspan="2" align="left">
		<select name="gd2" size="1" id="gd2" >
		<option <%if rs9("gd2")<>"0" then%>selected<%end if%> value="1">photo</option>
		<option <%if rs9("gd2")="0" then%>selected<%end if%> value="0">flash</option>
		</select>	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" height="18" align="right" bgcolor="#FFFFFF">首页视频点播URL：</td>
	<td colspan="2" align="left">

<IFRAME name=ad src="UploadMove.htm" frameBorder=0 width="100%" scrolling=no 
height=20></IFRAME>
		<input id=moveurl name=moveurl value='<%=rs9("moveurl")%>'  title='可直接输入网上MediaPlay文件的地址，或由上传程序自动产生MediaPlay文件地址' size="50" tyle="font-size: 9pt;  color: #000000; background-color: #FEEBE4; solid #FEEBE4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#FEEBE4'"><font color="#FF0000">从上传中自动获取地址，或者手动添加链接地址！</font>  	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" >上传文件类型：</td>
	<td  colspan="2" align="left" >
		<input type="text" name="UpFileType" size="50" value="<%=UpFileType%>" ><font color="#FF0000">用小写“|”分开</font>	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" >上传文件大小：</td>
	<td  colspan="2" align="left" >
		<input type="text" name="MaxFileSize" size="50" value="<%=MaxFileSize%>" ><font color="#FF0000">K</font>	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" >留言本评论过滤词语：</td>
	<td height="17" colspan="2" align="left" > 
		<input name="byteType" type="text"  value="<%=byteType%>" size="50"><font color="#FF0000">用小写“|”分开</font>	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" ><font color="#FF0000">网站恶意ＩＰ地址言论屏蔽：<font ></td>
	<td height="17" colspan="2" align="left" > 
		<input name="byteipType" type="text"  value="<%=byteipType%>" size="50"><font color="#FF0000">用小写“|”分开</font>	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" ><font color="#FF0000">网站恶意广告言论字符屏蔽：<font ></td>
	<td height="17" colspan="2" align="left" > 
		<input name="bytezfType" type="text"  value="<%=bytezfType%>" size="50"><font color="#FF0000">用小写“|”分开</font>	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" >浏览文章IP限制设置，<br>
	  (填写为允许、发文需要相应设置)：</td>
	<td height="17" colspan="2" align="left" > 
		<input name="newsipType" type="text"  value="<%=newsipType%>" size="50"><font color="#FF0000">用小写“|”分开</font>	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="193"  align="right" bgcolor="#FFFFFF">
		自定义顶部（TOP）菜单：<br>
		<br>
		<font color="#FF0000">HTML格式书写，如不支持FSO，<br>
		编辑config.asp</font><br>	</td>
	<td colspan="2" align="left">
		<textarea name="echuangmenu" cols="80" rows="6" ><%=echuangmenu%></textarea><font color="#FF0000"></font>	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="193"  align="right" bgcolor="#FFFFFF">
		自定义底部（BOTTOM）菜单：<br>
		<br>
		<font color="#FF0000">HTML格式书写，如不支持FSO，<br>
		编辑config.asp</font>	</td>
	<td colspan="2" align="left">
		<textarea name="BOTTOMmenu" cols="80" rows="6" ><%=BOTTOMmenu%></textarea>	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" bgcolor="#FFFFFF">首页浮动广告：</td>
	<td colspan="2" align="left">
		<select name="R_BG" size="1" >
		<option <%if rs9("R_BG")<>"0" then%>selected<%end if%> value="1">启用</option>
		<option <%if rs9("R_BG")="0" then%>selected<%end if%> value="0">关闭</option>
		</select>	</td>
    </tr>
<tr bgcolor="#FFFFFF">
  <td align="right" bgcolor="#FFFFFF">其他分页浮动广告：</td>
  <td colspan="2" align="left"><select name="AD_Show" size="1" >
    <option <%if rs9("AD_Show")="1" then%>selected<%end if%> value="1">启用</option>
    <option <%if rs9("AD_Show")="0" then%>selected<%end if%> value="0">关闭</option>
  </select></td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" bgcolor="#FFFFFF">浮动广告类型：</td>
	<td colspan="2" align="left">
		<select name="ad_class" size="1" >
		<option <%if rs9("ad_class")="1" then%>selected<%end if%> value="1">图片</option>
		<option <%if rs9("ad_class")="0" then%>selected<%end if%> value="0">FLASH</option>
		</select>	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" height="18" align="right" bgcolor="#FFFFFF">浮动广告地址：</td>
	<td colspan="2" align="left">
		<input type="text" name="R_TOP" size="50" value="<%=rs9("R_TOP")%>" >
		*注意广告文件类型要与上行选择相配	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" bgcolor="#FFFFFF">浮动广告链接地址：</td>
	<td colspan="2" align="left">
		<input type="text" name="L_MAIN" size="50" value="<%=rs9("L_MAIN")%>" >	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" bgcolor="#FFFFFF">浮动广告说明：</td>
	<td colspan="2" align="left">
		<input type="text" name="R_MAIN" size="50" value="<%=rs9("R_MAIN")%>" >	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" bgcolor="#FFFFFF">首页弹出公告：</td>
	<td colspan="2" align="left">
		<select size="1" name="E_announceview" style="font-family: 宋体; font-size: 9pt" >
	<option <%if rs9("E_announceview")=1 then%>selected<%end if%> value="1">开启</option>
	<option <%if rs9("E_announceview")=0 then%>selected<%end if%> value="0">关闭</option>
	</select>	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" bgcolor="#FFFFFF"> <p align="right">状态栏信息：</p></td>
	<td colspan="2" align="left" >
	<textarea  name="gg1" cols="80" rows="4" ><%=rs9("gg1")%></textarea>	</td>
</tr>
<tr class="TDtop"> 
	<td colspan="3"><div align="center">显示条数设置</div></td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" >首页文章配图显示张数：</td>
	<td colspan="2" align="left">
		<select size="1" name="L_BG" >
		<%for i=1 to 3%>
			<option <%if rs9("L_BG")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" >文章显示条数：</td>
	<td colspan="2" align="left">
		<select size="1" name="top_news" >
		<%for i=1 to 30%>
			<option <%if rs9("top_news")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" >大类页面文章显示条数：</td>
	<td  colspan="2" align="left" >
		<select size="1" name="bigclassshownum" >
		<%for i=1 to 30%>
			<option <%if rs9("bigclassshownum")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" >专题显示条数：</td>
	<td  colspan="2" align="left" >
		<select size="1" name="top_sp" >
		<%for i=1 to 30%>
			<option <%if rs9("top_sp")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" >热门文章显示条数：</td>
	<td  colspan="2" align="left" > <select size="1" name="top_txt" >
		<%for i=1 to 30%>
			<option <%if rs9("top_txt")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" >最新图文显示条数：</td>
	<td  colspan="2" align="left" > <select size="1" name="top_img" >
		<%for i=1 to 30%>
			<option <%if rs9("top_img")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" >阅读文章页面评论条数：</td>
	<td  colspan="2" align="left" >
		<select size="1" name="reviewnum" >
		<%for i=1 to 30%>
			<option <%if rs9("reviewnum")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" >图片JS文章调用数：</td>
	<td  colspan="2" align="left" >
		<select size="1" name="picnum" >
		<%for i=1 to 30%>
			<option <%if rs9("picnum")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" >一般JS文章调用数：</td>
	<td  colspan="2" align="left" >
		<select size="1" name="newsnum" >
		<%for i=1 to 30%>
			<option <%if rs9("newsnum")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" >用户、单位部门排行榜显示数：</td>
	<td  colspan="2" align="left" >
		<select size="1" name="topuser" >
		<%for i=1 to 30%>
			<option <%if rs9("topuser")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="193" align="right" >首页友情链接（LOGO）链接显示数：</td>
	<td  colspan="2" align="left" >
		<select size="1" name="linkshownum" >
		<%for i=1 to 30%>
			<option <%if rs9("linkshownum")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td colspan="3">
		<div align="center"> 
		<input type="submit" name="Submit" value="提交" >
		<input type="reset" name="Submit2" value="重设" >
		</div>	</td>
</tr>
</form>
</table>
</div>
</body>
</html>
<!--#include file="Admin_Bottom.asp"-->
		<%rs9.close
		set rs9=nothing
		conn.close
		set conn=nothing
	else
		Show_Err("对不起，该功能已经被超级系统管理员关闭，您没有权限操作！")
		response.end
	end if
end if%>