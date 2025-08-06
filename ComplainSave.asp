<%@ LANGUAGE = VBScript.Encode %>
<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="md5.asp"-->
<!--#include file="ChkURL.asp"-->

<%#@~^iAgAAA==@#@&ZGsw^lk	qGxZ43]+$EndD`D5E/YvE/Wsw^lrx(9r#SF*7B防注入@#@&eG!D1Ch'^t^0/OM`M+;!+kOR6W.:vJeG!D1mh+r##@#@&3:mk^'^tn^0/ODv.+$EndDR0G.s`E2sCbVE*#@#@&KVKtKxn'1tn^0/YMcD;EdOR6WM:cJPn^+KtKU+r#b@#@&bN[./d'14mVkYM`D;!n/DRWWM:cEzNNMn/kJ#*@#@&}bw{m4+^VkY.`Mn;!+dOc0W.hvJtkaE*#@#@&Z^l/kqGx`M+5E/O 6WDscJ;Vlkd(fr#*@#@&KG2bmxm4nm0/O.vD+5;/OR6GM:crKKwk1J*b@#@&mGxD+UO{YDbh`4Y:^nUmKNFc`.n$En/D 0KDhcrmWUOxOJ*b*#@#@&mKxYxDxDwsl1+c^KxYUYBJ@!a@*~JBJr#@#@&^G	YnxDxDwsC1+`^G	YnxDSr@!K@*Pr~Jr#@#@&Sl[Hmks(h'm4nm0/YMc]+$E/ORdnM\nDjCDbl8s/`E]AH6KAmzf9"J*#@#@&IdwKxd+cmGG0k+kc+;t/Hdb`rmKxO+UOr#xmKUYxO@#@&@#@&[rsP8XDn8@#@&(XD+F{/askD`8XD+Pza+~rkJ*@#@&@#@&WWMPb'TPOG,E8W!UNv4zOF#@#@&imGxDn	YxM+aVl1+v^W	YnxD~O.b:`(zYF`bbb~reCeE#@#@&	+aY@#@&@#@&Nrh,4XOny@#@&4HO xkw^kYv4HO+bwPXa+SE-J#@#@&@#@&0WM~r'ZPDW~E8G!x[`(zY b@#@&k0~];;+kOc/nM\D.mDbC4^+d`rI3\}K2|)fGIJ*x8XD+y`r#~O4+U@#@&7?4WAmADDcE对不起，你的&n地址已被锁定，请联系管理员！！！。@!4M@*@!4.@*@!mPtM+6xBNl-lkm.raY)4r/DWDH 8l13v#v@*返回@!&m@*E#@#@&d"+d2Kx/n 1WG3bnk`n;tkX/*`r^W	YnxDJbxrJ@#@&7I/wKUd+c2	N~@#@&n	N~k6@#@&	+aO@#@&@#@&[b:~4HO&@#@&4HY+2'k2VbYc4HYn"6KXan~ruJ*@#@&@#@&0KD~kxT,YGP!8W!x[c(XYnf*@#@&k6~&xdDDvD+$EdYvJ^W	YnUDJ#B8XD+&vrb#@*!,POtnU@#@&7?4Gh|2..vJ对不起，请不要发布非法小广告！！！。@!8.@*@!8D@*@!mP4M+6'BNl7C/1DrwD)4rkYWMzR(lm0cbB@*返回@!Jl@*Jb@#@&d]+k2W	/n 1WWVr/c+;4kXd*`rmW	YUYr#xJr@#@&7"+/aGxk+RAU[P@#@&x[PrW@#@&U+XO@#@&@#@&r6PqUdDDcD5!+dD`rmW	YUYr#SJEJb@*ZPWM~q	/YMc.+$E/O`E^KxO+	OJ*~Ed1Dk2Or#@*!,GMP(	/DD`M+$;+kYcJ1WUOxYrb~rWx;srm0J*@*TP~GMP(xkODvDn5!+/OcrmGxDn	YE*~rWx^Wm[J*@*TPDtnU@#@&dU4WS{2M.cJ对不起，您输入的投诉内容包含有非法字符。@!4M@*@!4.@*@!l~tMn0{B%C7l/^.bwO)4rkYGMXc4l13vbB@*返回@!&l@*Jb@#@&dIdwKx/ 3x9P@#@&nx[~b0@#@&@#@&k6P/GswVCr	q9'rE,Y4x@#@&dk+D~Dk'd+M\n.cmDCYW4Nn^YvJmNGN8 M+^WM[/YEb@#@&dd5^'E/smO,e,0DK:,E[,N8{AZm/K:w^Ck	{Km8s+,[rJ~@#@&7M/ Wanx,/5sBmWUUBFS&@#@&iDdcl9Nxh@#@&dM/cJIW;.gl:E#{5W!.Hls+@#@&7Ddcr+hlbsJ*'nhmkV77i@#@&dMdvJPVntKxE#{KnVn4G	+@#@&7Dk`JdnCNtlbV(nEb{Snl9\lbV(K@#@&d.dvJ)N9./dr#{bN9Dd/@#@&7Dk`EtbwJ*x}bw@#@&7./vJ!w[lOnDkh+rb'	WAc*@#@&7.k`EZ^Ck/(GJ*'Z^lkdqG@#@&dM/cE:Wwb^J*'KK2rm@#@&iDd`E^KxO+	OJ*'^G	Y+UOi@#@&dMdcE29lD+@#@&dMdR1VG/@#@&7k+Y,./{xWD4rxT@#@&d]+d2Kxd+c^WK3rnk`+/4kXd#vE1WUD+	YJ*'rE@#@&+UN,kW@#@&dG8CAA==^#~@%>
<html>
<head>
<title>投诉举报成功__<%=#@~^BgAAAA==n;tjH/TwIAAA==^#~@%></title>
<meta http-equiv=refresh content="1;url=Complain.asp">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="news.css" rel=stylesheet type=text/css>
</head>
<body>
<!--#include file="other_Top.asp"-->
      <table width="780" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#F1F7FC">
        <tr valign="top">
          <td>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr bgcolor="CDCCCC">
                <td height="25" background="Images/dh_bg.gif" class="daohang">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;栏目导航&nbsp;&nbsp;<a class=daohang href="./" >网站首页</a>&gt;&gt;<a href="Complain.asp" class="daohang">网上投诉举报</a>&gt;&gt;投诉举报成功</td>
              </tr>
          </table></td>
        </tr>
        <tr valign="top">
          <td>
            <table width="99%" height="235" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
              <tr>
                <td height="235" align="center">
                  <table border="0" cellspacing="0" width="30%" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF" cellpadding="0">
                    <tr>
                      <td width="100%" height="20"><p align="center"><font color=red><b>投诉举报成功&nbsp;&nbsp;&nbsp;&nbsp;</b></font><a class=daohang href="Complain.asp">返回</a> </p></td>
                    </tr>
                </table></td>
              </tr>
          </table></td>
        </tr>
        <tr valign="top">
          <td height="10"></td>
        </tr>
</table>
<!--#include file="other_bottom.asp"-->
<%#@~^SwAAAA==]/2Kxk+R1WKVk/c+;tdzk#`r^W	Y+	OE#{Jr@#@&B.nkwGxknRM+[rM+mO~rZG:asmkUclkwJcRkAAA==^#~@%>