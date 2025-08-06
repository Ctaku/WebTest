<%@ LANGUAGE = VBScript.Encode %>
<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<%#@~^UAEAAA==@#@&@#@&0+HhWMN{ODb:cm4+^VkYDv]+$E+kOcJ0+HhGD[E*#b@#@&@#@&hlLnUtWAjbynP{~8!7E每页显示多少个页@#@&HXhlTn?bynP,Px~8!dE每页显示多少条@#@&i@#@&&W~1KY,qd1;hDrmv]+$EndD`J2CT+E#*~}D~&/A:wDXv]+$En/D`E2mo+rb#,rD,]n;!+kYcJ2CT+E#,@!'ZPP4x@#@&7tXKlTn{F@#@&2^/+@#@&i\XhlL+{qUOvb4kcI;EdO`rwmonJbb*@#@&2	[Pb0@#@&@#@&kW~0+zhK.9'ErPKDP,3zhKD['r请输入标题关键字J~O4+x@#@&dg1QAAA==^#~@%>
	<script language=javascript>
		alert("请输入查询关键字！")
		history.back()
	</script>
	<body onload=javascript:window.close()></body> 
	<%#@~^3wAAAA==@#@&d]/aWxk+c3x9@#@&+	N~r6@#@&@#@&/YPMdx/D7+.R/.lO+}8LmOcrbf69~R]+1GMNjYr#@#@&/$sF{Jd+^+^O,eP6.WsPJL~[4|2;{J+C[tlrV|Pl(Vn~LJPA4DnPDGak^,Vb3+,BuE[0+zhKD['r]B,~WMN+M~8X,Sl[HCr^q9PG3?;J@#@&M/R62x~/$s8~^Kx	~FBF@#@&@#@&nz8AAA==^#~@%>
<html>
<head>
<title>信件搜索结果__<%=#@~^BgAAAA==n;tjH/TwIAAA==^#~@%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="news.css" rel=stylesheet type=text/css>
</head>
<body bgcolor="#FFFFFF" text="#000000" topmargin="0">
<!--#include file="other_Top.asp"-->

<table width="780" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
        <tr valign="top">
          <td height="163"><table width="780" height="194" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td width="190" valign="top" background="images\left_bg.gif"><table width="169" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="169" height="43" background="images\title_bg.gif"><div align="center" class="yellow_title">领导信箱</div></td>
                </tr>
              </table>
                <table width="169" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td><img src="Images/btn_box_top.gif" width="169" height="20"></td>
                  </tr>
                  <tr>
                    <td width="169" height="121" valign="top" background="images\bg_content_169px.gif"><div align="center" class="yellow_title">
                      <table width="98%" height="127" border="0" cellpadding="0" cellspacing="0">
                        <tr>
                          <td width="31%" height="33" align="center" valign="middle"><img src="Images/Sicon_QYTS.gif" width="22" height="22"></td>
                          <td width="69%" align="left" valign="middle" ><a href="E_LeadMail.asp" class="top1">领导信箱</a></td>
                        </tr>
                        <tr>
                          <td height="32" align="center" valign="middle"><img src="Images/Sicon_XZBSZX.gif" width="22" height="22"></td>
                          <td align="left" valign="middle"><a href="E_GuestBook.Asp" class="top1">公众留言</a></td>
                        </tr>
                        <tr>
                          <td height="31" align="center" valign="middle"><img src="Images/Sicon_XZTS.gif" width="22" height="22"></td>
                          <td align="left" valign="middle"><a href="Complain.asp" class="top1">投诉举报</a></td>
                        </tr>
                        <tr>
                          <td align="center" valign="middle"><img src="Images/Sicon_MYZJ.gif" width="22" height="22"></td>
                          <td align="left" valign="middle"><a href="E_Opinion.asp" class="top1">建言献策</a></td>
                        </tr>
                      </table>
                    </div></td>
                  </tr>
                  <tr>
                    <td><img src="Images/btn_box_bottom.gif" width="169" height="22"></td>
                  </tr>
                </table></td>
              <td width="588" valign="top"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="9%" height="53" valign="middle" background="Images/location_bg.gif" align="center" >&nbsp;<img src="Images/point_location.gif" width="24" height="24"></td>
                  <td width="91%" valign="middle" background="Images/location_bg.gif" class="daohang">&nbsp;您的位置：&nbsp;<a class=daohang href="./" >网站首页</a>&gt;&gt;<a href="E_LeadMail.asp" class="daohang">领导信箱</a>&gt;&gt;信件搜索</td>
                </tr>
              </table>
                <table width="560" height="300" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td height="42" background="images\topic_bg.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="yellow_title">信件搜索结果</span></td>
                  </tr>
                  <tr>
                    <td height="300" valign="top" background="images\bg_content_560px.gif">
                      <br>
                      <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#D3E3ED" id="AutoNumber1" style="border-collapse: collapse">	
                        <tr class="TDtop1" align="center">
                          <td width="7%" height="25" bgcolor="#D3E3ED"><span class="top1">状态</span></td>
                          <td width="43%" bgcolor="#D3E3ED" class="top1">信件标题</td>
                          <td width="12%" bgcolor="#D3E3ED" class="top1">来信日期</td>
                        </tr>
                        <%#@~^3QAAAA==@#@&d7i@#@&k0,xKOPM/ 2}s~O4+x@#@&DkRnmLn?byP~P~~{P\XhCo?r"@#@&\CXnCod,P~,P,PP,',./cnCoZG;	Y@#@&./cl4kGsED+alL+~x,HznmL+@#@&OGDlV~~,P~P,~,P~,',D/cI^WMN/W!xO@#@&Nks~k@#@&k,x~F@#@&9W~EUObV~Dk 2K0~GMPk~x,DdRhCT+jby@#@&EjoAAA==^#~@%> 
				<%#@~^RQAAAA==~iYGak1'YMkscDk`EYKwr^r##@#@&diddGrdw^lHKGwr^{:rNvPWak^S8~ 0b@#@&7di7MBIAAA==^#~@%>
                          <tr bgcolor=ffffff>
                            <td height="23" align=center bgcolor="#ffffff" >
					        <%#@~^JwAAAA==r6POMks`Dk`r]+aVzZKxOn	YJ*b@!@*JJ,O4+	PCgwAAA==^#~@%>
					        <font color=red >已阅批</font>
					        <%#@~^BAAAAA==n^/nqQEAAA==^#~@%>
					        <font color=red >未阅批</font>
					        <%#@~^BgAAAA==n	N~b0JgIAAA==^#~@%>                            </td>
<td align="left" bgcolor="#ffffff">
·<a  class=middle href='E_ReadLeadMail.asp?LeadMailID=<%=#@~^EAAAAA==.k`Ed+mNHmk^(fr#AAUAAA==^#~@%>' target="_black"><%=#@~^GAAAAA==4D:sx1WN`Gr/aVCX:W2r1#SQkAAA==^#~@%>...</a></td>
                            <td align=center bgcolor="#ffffff"><%=#@~^FwAAAA==zl.vDk`J`w9CYKr:Jbb,jgcAAA==^#~@%>-<%=#@~^GAAAAA==\KxO4`M/`rja[lD+Pks+Eb*P4wcAAA==^#~@%>-<%=#@~^FgAAAA==9mXcM/vJjaNmO+:kh+r#b~+wYAAA==^#~@%></td>
                          </tr>
					    <%#@~^MAAAAA==@#@&d7M/cHW7+gn6D@#@&dik~x,kP3~F@#@&disGWa@#@&d7kggAAA==^#~@%>
		<tr> 
	<td colspan="5" width="100%" align=center height=25 bgcolor="#ffffff">共搜索到 <%=#@~^BQAAAA==OKYC^JAIAAA==^#~@%> 条，当前第 <%=#@~^BgAAAA==\HwCT+YwIAAA==^#~@%>/<%=#@~^CAAAAA==\m62mo/NgMAAA==^#~@%> 页，每页 <%=#@~^CgAAAA==\HnCT+Uky3gMAAA==^#~@%> 条
	<%#@~^dwcAAA==@#@&d;MV{J2|SCNtlrV|Ind!VYcC/ag3zAWMN{J~[~VXAWM[@#@&nCL1+aOUk"+{r	YcvHHnlT+ q#JnCo?4GS?k.n#3F@#@&KCoYalL+xr	Yc`DGYmVRq*zDd hlL+Ur.+b3F@#@&@#@&k6~nmon16Ojby+,@*F,YtU@#@&hlT+KDn-{nCojtKhjr.+ecKmon1aD?r.+ F#@#@&"n/aWU/RA.bY+,E@!mPm^Cd/{4^l^3~4M+W'EEPLPi.^P[~ELwCoxrP',nmo+hD-PLPEB,YrO^+'E上EPLPnmLn?4WS?ryn~LPE页B@*上一翻页@!zm@*~E@#@&IndaWU/ SDrD+,J@!mP1slk/x4^l^V,tDW'EJPL~iD^PLPE[2CT+xFE~YbYsn{B第F页v@*页首@!zC@*,E@#@&n	N,k0@#@&bWPtXKlT+Rq,@*PZ~Y4+x@#@&KD\|nCon~{P\XhCoPR~8@#@&]nkwGxknch.bYPJ@!l,^Vm/d'(VC^0PtMn0{BJ,'~jMV,[~J'2mon'r~[,n.n7{nCLP'Prv,YrDV'B第rPL~nM+-{hlLn,[Pr页v@*上一页@!zl@*~E@#@&+	N~kW@#@&@#@&k6~Hm62CT+/@*xhlL+gnXYjbyenmojtKhjk.+~O4+x@#@&nmo+Ur"+UtKh~'~Kmon?4GhUk"n@#@&2sd@#@&nmL?r.+UtWSP{~Hm62lT+dRhlojtKh?b"nevnmon1naD?ryRF*@#@&3	NPrW@#@&(0,KmonUk.+?4WS~@!,F~K4+U~hlojk.+?4GAP{P8@#@&0G.,nCo/W!xOnM?k"n{F~YK~hlL?by+UtKA@#@&nCoSrU0P',cnmo+;G;xD+M?rynQhlL+gn6D?r"enCL?4WSjbyn*Ohlo?4GhUk"+@#@&rW,nlTnSbx3,@!@*PtXhlL+~P4+U@#@&]+kwGUk+RA.bYnPr@!mP^^lk/'(Vm^3,t.+6'vE,[P`.V,[Pr'2lT+{J~[~KmonSbU3,[~EE@*$E~LPKlTndkU0PLPJY@!JC@*,J@#@&Vdn@#@&IdwKx/ 	DbYPE@!$@*]J'PhCoSrU0P[ED@!z$@*,E@#@&n	N,k0@#@&&WPhlL+dkUV,'PtC6hlod~K4+	P36rO,0GD@#@&16O@#@&@#@&rW,HzwmL_q,@!{nlT+D2lT+~PDtnU@#@&1aY|nlTn~',HHnCon~3Pq@#@&]+kwGUk+RA.bYnPr@!mP^^lk/'(Vm^3,t.+6'vE,[P`.V,[Pr'2lT+{J~[~H6O{hCoP'~rBPOrDVn'E第E,[~g+XY{hlTnPLPE页B@*下一页@!&)@*J@#@&nx9Pk6@#@&@#@&k6P\laKmon/,@*PhlLnUtWAjbynehCT+H6D?k.+,Otx@#@&hlLng+6D~',nlTnjtKhUk"+~M,nCoH+XYjr.+PQ~8@#@&IdaWUk+chDbY~J,@!)P1VCdk'4^Cm0PtMnW'EJ,[~j.s,[~JL2lT+xE,[PKCT+OwmLP',JEPYbY^n'E第J'PhlLnDwlTnPLJ页B@*页尾@!&b@*J@#@&]+d2Kxd+cADbYn~rP@!C~1VC/kx(VC13,tD0{vJ,[~jMV~',J[aCo'J,'~nmo1n6O~LPEB,OkDVnxE下JP'~hlL+U4KhjbyP[,J页E@*下一翻页@!Jl@*J@#@&3U9Pk6@#@&V/@#@&I/aWU/n SDrY~J@!Y.@*@!YN~C^kLx{^xOD,4o1W^GD{aoZwf3$,mW^dwmx'2~4+bo4YxFTT@*@!WW	OP1WsGM'Dn[@*对不起，没有找到你要搜索的信件！@!&0KUD@*@!JY9@*@!JYM@*J@#@&7di@#@&3	NP&W@#@&D/c^sWk+@#@&d+O~M/xxKOtbxL@#@&QRQCAA==^#~@%>
</td>
</tr>			
</table>

                      <br>
				    </td>
                  </tr>
                  <tr>
                    <td height="13"><img src="Images/shadow_690px.gif" width="560" height="17"></td>
                  </tr>
                </table></td>
            </tr>
          </table>
          </td>
        </tr>
</table>
<!--#include file="other_bottom.asp"-->
