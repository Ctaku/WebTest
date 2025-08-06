<%
set rs3=server.CreateObject("ADODB.RecordSet")
if uselevel=1 then
	if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then
		rs3.Source ="select top " & top_img & " * from "& db_EC_News_Table &" where picnews=1 and checkked=1 and picname<>'null' order by NewsID DESC"
	end if
	if Request.cookies(eChsys)("key")="" then
		rs3.Source ="select top " & top_img & " * from "& db_EC_News_Table &" where picnews=1 and checkked=1 and newslevel=0  and picname<>'null' order by NewsID DESC"
	end if
	if Request.cookies(eChsys)("key")="selfreg" then
		if Request.cookies(eChsys)("reglevel")=3 then
			rs3.Source ="select top " & top_img & " * from "& db_EC_News_Table &" where picnews=1 and checkked=1 and picname<>'null' order by NewsID DESC"
		end if
		if Request.cookies(eChsys)("reglevel")=2 then
			rs3.Source ="select top " & top_img & " * from "& db_EC_News_Table &" where picnews=1 and checkked=1 and picname<>'null' order by NewsID DESC"
		end if
		if Request.cookies(eChsys)("reglevel")=1 then
			rs3.Source ="select top " & top_img & " * from "& db_EC_News_Table &" where picnews=1 and checkked=1 and picname<>'null' order by NewsID DESC"
		end if
	end if
else
	rs3.Source ="select top " & top_img & " * from "& db_EC_News_Table &" where picnews=1 and checkked=1 and picname<>'null' order by NewsID DESC"
end if
rs3.Open rs3.Source,conn,3,3
if not rs3.EOF then%>
<div align='center' id='demo' style='overflow:hidden;height:110px;width:608px;'><!--滚动区的高度和宽度-->
<table align='center' cellpadding='0' cellspace='0' border='0'>
<tr>
	<td id='demo1' valign='top'>
		<table width='100%' cellpadding='0' cellspacing='0' border='0' align='center'>
		<tr valign='top'>
	<%
	while not rs3.EOF
		fileExt=lcase(getFileExtName(rs3("picname")))
		%>
			<td align="center">
	<div class="pic_cont">					
		<%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
		<a class=middle href='E_ReadNews.asp?NewsID=<%=rs3("NewsID")%>' target=_blank title='<%=rs3("title")%>'><img  src='<%=FileUploadPath & rs3("picname")%>' width='100' height="75" border=0></a>					
		<%else if fileext="swf" then%>
							<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width='100' height='75'>
								<param name=movie value='<%=FileUploadPath & rs3("picname")%>'>
								<param name=quality value=high>
								<param name='Play' value='-1'>
								<param name='Loop' value='0'>
								<param name='Menu' value='-1'>
								<param name='wmode' value='transparent'>
								<embed src='<%=FileUploadPath & rs3("picname")%>' width='100' height='75' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
</object>						
<%end if%>
<%end if%>				
<a class=middle href='E_ReadNews.asp?NewsID=<%=rs3("NewsID")%>' target=_blank title='<%=rs3("title")%>'><%=CutStr(htmlencode4(rs3("title")),12)%></a>
</div>
</td>
	  <%
	  rs3.MoveNext
	wend
	%>
		</tr>
		</table>
	</td>
	<td id=demo2 valign=top></td>
</tr>
</table>
</div>
	<%if top_img >4 then%>
<script>
var speed=30
demo2.innerHTML=demo1.innerHTML
function Marquee1(){
if(demo2.offsetWidth-demo.scrollLeft<=0)
demo.scrollLeft-=demo1.offsetWidth
else{
demo.scrollLeft++
}
}
var MyMar1=setInterval(Marquee1,speed)
demo.onmouseover=function() {clearInterval(MyMar1)}
demo.onmouseout=function() {MyMar1=setInterval(Marquee1,speed)}
</script>
	<%end if
else
	Response.Write "暂 无 最 新 图 文"
end if
rs3.close
set rs3=nothing
%>
<!--最新图文代码结束-->