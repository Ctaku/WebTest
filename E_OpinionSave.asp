<%@ LANGUAGE = VBScript.Encode %>
<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="md5.asp"-->
<!--#include file="ChkURL.asp"-->

<%#@~^4AgAAA==@#@&r2bxbWx&f{/t0In;!+dOvD+$;+kY`r62k	kKx(fEbBFbdE防注入@#@&}wrUbWx/sm/d'14mVkYM`D;!n/DRWWM:cE}wk	rW	ZVmddJ*#@#@&eW;.glh+{^tmVdDD`.n$En/D 6W.s`r5W!DgC:Jb#@#@&3hmkV{^tm3kO.`M+$En/O 6W.:vE2slrsr##@#@&:+s+h4Kxn{m4+m0/D.`M+5E/O 6WDscJ:+VK4W	+r#b@#@&)9N.+kd'1tn^0/Y.cM+5EdDRWKDs`JzN9.+k/E#*@#@&tbw'14+13/D.cD;!+dY WKDh`rtkaJbb@#@&ZsCk/(f{cM+5!+kYR6WMh`rZslk/(9r##@#@&KKwk1x^tm0/ODc.;;+kOR6W.hvJKG2bmE#*@#@&mG	YxY{YMr:vtO:^+U^KN+8c`M+;!ndYc0KDh`E^KxO+	OJ*#bb@#@&mGUD+UY{.wsmm`mKxDnxD~E@!a@*~EBJJ*@#@&1WxDnUY{Dwsl^nvmGxDnxD~E@!h@*PESrJb@#@&6akUbW	qn{m4nm0/ODvIn5!+/D /D\.#lMkm4s+dcrI3H}P2|b99"J#b@#@&In/aG	/ncmKW3b+kc+;tdXk#cE1WxDnxDJ#{^GxD+	Y@#@&@#@&9khP(zYF@#@&(XYnq{/2VbOv4zD+:Xw~rkJ*@#@&@#@&0G.,k'Z~YKPE(G;x9`(XO+qb@#@&7mKUYxOxM+wsC1+cmKUD+UD~DDks`(zYFck*#SECeerb@#@&x+XO@#@&@#@&9khP8zD++@#@&8XD++xkwVrOv4zYraKza+BJur#@#@&@#@&0GD,kxT,YW,;4KEx9c8XD+y#@#@&rW,In;!n/DRdnM\+.#mDrl(s/crIAHr:2|)fGIE#{4zO `bbPDt+	@#@&dUtKhm2..vJ对不起，你的(n地址已被锁定，请联系管理员！！！。@!8D@*@!8.@*@!l~4M+W'E%m\CkmMkwD)4r/DW.Xc4C^0`#E@*返回@!Jl@*rb@#@&iI/2WUdR^WKVk/cn;t/zd*`EmKUD+UDJ*'Jr@#@&7I/2W	/n AxN,@#@&xN,rW@#@&x6O@#@&@#@&Nr:,8XD+f@#@&4XOn2'dw^rD`8HYy0:Xan~ruE#@#@&@#@&6WD,r'ZPYK~;4KE	Nc4zO&b@#@&r0,qUdDD`.n$En/DcrmG	YxYr#B8XD+f`b#b@*ZPPD4+	@#@&ij4WS{AD.`E对不起，请不要发布非法小广告！！！。@!(D@*@!(.@*@!l~4M+0xvNl-lk^Mk2D)4k/DWMzR(l^3v#v@*返回@!zl@*E#@#@&d"ndwKxk+ mGG0kn/vnZ4/zd*`J^G	YnxDE*'Er@#@&dI/aGxk+ 2	N~@#@&+x9~k6@#@&	naY@#@&@#@&r0~(	/ODv.+$EndD`J^G	YnxDE*~EEJ*@*!,WM~q	/ODvDn5!+/DcJ1WxDnUYr#BJdm.raYE#@*TPKD~(	/Y.cM+5EdD`E1W	Y+	Yrb~rWUZ^k^Vr#@*Z~PKDP&UdYM`M+5EndD`EmKUYxOE*~JGU^WCNrb@*!~Dtx@#@&dU4WS{3DM`E对不起，您输入的信息内容包含有非法字符。@!(D@*@!8D@*@!l,4.+6'ELC\Cd1DrwDltb/OGMXR8C13c#E@*返回@!zC@*J*@#@&iIdwKxd+c2U[,@#@&UN,k0@#@&@#@&b0,r2kUrKx(f{EJ,Y4n	@#@&7dY~Dkxk+.7+MRmM+mO+K4%+1YcEmNW98RM+mK.[/Yr#@#@&7d$VxJknVmO~CP0.GsPE[,[({3;{}wk	kKU{:l8VP'ErP@#@&7DkRWanUPk;^~^WUUBFS&@#@&dM/ C9NxnA@#@&7Dkcrr2bxbWx;Vmd/r#xrakUrKxZ^C/k@#@&i.d`r5KE.1ChJb'IGEM1Ch@#@&7.k`E+sCbVE*':lbVi7d@#@&7Dk`EPV+h4W	+J*xP+^+htGxn@#@&d./vErakUrKxqKE*'6wbUbWU&n@#@&dM/vEb9N.+k/Eb{bN9.+k/@#@&7./vJ\k2Jbx\k2@#@&7Dk`E;aNlOnDkh+rb{xGS`*@#@&iDkcJ;VC/kq9E*'Z^C/kqf@#@&7Dk`rKGwr^r#xKK2k1@#@&7M/`E^KxO+	Or#x1W	Y+	Y@#@&d@#@&7DkR;29lY@#@&iD/c^sWk+@#@&7/nO,Dd'	GY4kUL@#@&d]nkwGxkncmGK3b+/v+;4/H/b`rmGUD+xDE#{JJ@#@&nx9Pb0@#@&c4wCAA==^#~@%>
<html>
<head>
<title>发言成功__<%=#@~^BgAAAA==n;tjH/TwIAAA==^#~@%></title>
<meta http-equiv=refresh content="1;url=E_Opinion.asp">
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
                <td height="25" background="Images/dh_bg.gif" class="daohang">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;栏目导航&nbsp;&nbsp;<a class=daohang href="./" >网站首页</a>&gt;&gt;<a href="E_Opinion.asp" class="daohang">发信成功__&lt;%=eChSys%&gt;</a>&gt;&gt;发言成功</td>
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
                      <td width="100%" height="20"><p align="center"><font color=red><b>发言成功&nbsp;&nbsp;&nbsp;&nbsp;</b></font><a class=daohang href="E_Opinion.asp">返回</a> </p></td>
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
<%#@~^TAAAAA==]/2Kxk+R1WKVk/c+;tdzk#`r^W	Y+	OE#{Jr@#@&B.nkwGxknRM+[rM+mO~r2mrar	kG	Rm/wrvhkAAA==^#~@%>