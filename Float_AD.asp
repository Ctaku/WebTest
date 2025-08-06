<%@ Language=VBScript%>
<!--#include file="conn.asp" -->
<!--#include file="function.asp" -->
<!--#include file="config.asp" -->
<%if ad_class=1 then %>
suspendcode="<DIV id=lovexin1 style='Z-INDEX: 10; LEFT: 10px; POSITION: absolute; TOP: 105px; width: 95; height: 203px;'><img src='images/ad_bar.gif' onClick='javascript:window.hide()' width='84' height='20' border='0' vspace='0' alt='关闭对联广告'><a href='<%=L_MAIN%>' target=_blank><img src='<%=R_TOP%>' border=0 width='84' height='300' alt='<%=R_MAIN%>' align=top></a></DIV>"
document.write(suspendcode);
suspendcode="<DIV id=lovexin2 style='Z-INDEX: 10; LEFT: 912px; POSITION: absolute; TOP: 105px; width: 95; height: 203px;'><img src='images/ad_bar.gif' onClick='javascript:window.hide()' width='84' height='20' border='0' vspace='0' alt='关闭对联广告'><a href='<%=L_MAIN%>' target=_blank><img src='<%=R_TOP%>' border='0' width='84' height='300' alt='<%=R_MAIN%>' align=top></a></DIV>"
document.write(suspendcode);
lastScrollY=0;
function heartBeat(){
diffY=document.body.scrollTop;
percent=.3*(diffY-lastScrollY);
if(percent>0)percent=Math.ceil(percent);
else percent=Math.floor(percent);
document.all.lovexin1.style.pixelTop+=percent;
document.all.lovexin2.style.pixelTop+=percent;
lastScrollY=lastScrollY+percent;
}
function hide()  
{   
lovexin1.style.visibility="hidden"; 
lovexin2.style.visibility="hidden";
}
window.setInterval("heartBeat()",1);
<%else%>

suspendcode="<DIV id=lovexin1 style='Z-INDEX: 10; LEFT: 10px; POSITION: absolute; TOP: 105px; width: 90; height: 203px;'><img src='images/ad_bar.gif' onClick='javascript:window.hide()' width='84' height='20' border='0' vspace='0' alt='关闭对联广告'><object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0' width='84' height='300'><param name='movie' value='<%=R_TOP%>'><param name='quality' value='high'><embed src='<%=R_TOP%>' quality='high' pluginspage='http://www.macromedia.com/go/getflashplayer' type='application/x-shockwave-flash' width='100' height='300'></embed></object></DIV>"
document.write(suspendcode);
suspendcode="<DIV id=lovexin2 style='Z-INDEX: 10; RIGHT: 10px; POSITION: absolute; TOP: 105px; width: 90; height: 203px;'><img src='images/ad_bar.gif' onClick='javascript:window.hide()' width='84' height='20' border='0' vspace='0' alt='关闭对联广告'><object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0' width='84' height='300'><param name='movie' value='<%=R_TOP%>'><param name='quality' value='high'><embed src='<%=R_TOP%>' quality='high' pluginspage='http://www.macromedia.com/go/getflashplayer' type='application/x-shockwave-flash' width='100' height='300'></embed></object></DIV>"
document.write(suspendcode);
lastScrollY=0;
function heartBeat(){
diffY=document.body.scrollTop;
percent=.3*(diffY-lastScrollY);
if(percent>0)percent=Math.ceil(percent);
else percent=Math.floor(percent);
document.all.lovexin1.style.pixelTop+=percent;
document.all.lovexin2.style.pixelTop+=percent;
lastScrollY=lastScrollY+percent;
}
function hide()  
{   
lovexin1.style.visibility="hidden"; 
lovexin2.style.visibility="hidden";
}
window.setInterval("heartBeat()",1);
<%end if%>