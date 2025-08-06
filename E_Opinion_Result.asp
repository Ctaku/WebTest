<%@ LANGUAGE = VBScript.Encode %>
<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<%#@~^UAEAAA==@#@&@#@&0+HhWMN{ODb:cm4+^VkYDv]+$E+kOcJ0+HhGD[E*#b@#@&@#@&hlLnUtWAjbynP{~8!7E每页显示多少个页@#@&HXhlTn?bynP,Px~8!dE每页显示多少条@#@&i@#@&&W~1KY,qd1;hDrmv]+$EndD`J2CT+E#*~}D~&/A:wDXv]+$En/D`E2mo+rb#,rD,]n;!+kYcJ2CT+E#,@!'ZPP4x@#@&7tXKlTn{F@#@&2^/+@#@&i\XhlL+{qUOvb4kcI;EdO`rwmonJbb*@#@&2	[Pb0@#@&@#@&kW~0+zhK.9'ErPKDP,3zhKD['r请输入标题关键字J~O4+x@#@&dg1QAAA==^#~@%>
	<script language=javascript>
		alert("请输入查询关键字！")
		history.back()
	</script>
	<body onload=javascript:window.close()></body> 
	<%#@~^4QAAAA==@#@&d]/aWxk+c3x9@#@&+	N~r6@#@&@#@&@#@&/+D~./{/D-+. ;DnlDnr(Ln^D`J)9}f$R"n1W.9?YJ*@#@&d;^FxJk+sn1YPC~0MW:,E'P94|2/{62bxrW	mKm4sn,[J~A4+.+,OKwr1P^k3PEYJL3nXSW.[LJ]E~PKDN.~4HP}wrxrG	q9PG3?;J@#@&M/R62x~/$s8~^Kx	~FBF@#@&@#@&fD8AAA==^#~@%>
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
                  <td width="169" height="43" background="images\title_bg.gif"><div align="center" class="yellow_title">建言献策</div></td>
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
                  <td width="91%" valign="middle" background="Images/location_bg.gif" class="daohang">&nbsp;您的位置：&nbsp;<a class=daohang href="./" >网站首页</a>&gt;&gt;<a href="E_Opinion.asp" class="daohang">建言献策</a>&gt;&gt;发言搜索</td>
                </tr>
              </table>
                <table width="560" height="300" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td height="42" background="images\topic_bg.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="yellow_title">发言搜索结果</span></td>
                  </tr>
                  <tr>
                    <td height="300" valign="top" background="images\bg_content_560px.gif">
                      <br>
                      <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#D3E3ED" id="AutoNumber1" style="border-collapse: collapse">	
                        <tr class="TDtop1" align="center">
                          <td width="7%" height="25" bgcolor="#D3E3ED"><span class="top1">状态</span></td>
                          <td width="43%" bgcolor="#D3E3ED" class="top1">发言标题</td>
                          <td width="12%" bgcolor="#D3E3ED" class="top1">发言日期</td>
                        </tr>
                        <%#@~^3QAAAA==@#@&d7i@#@&k0,xKOPM/ 2}s~O4+x@#@&DkRnmLn?byP~P~~{P\XhCo?r"@#@&\CXnCod,P~,P,PP,',./cnCoZG;	Y@#@&./cl4kGsED+alL+~x,HznmL+@#@&OGDlV~~,P~P,~,P~,',D/cI^WMN/W!xO@#@&Nks~k@#@&k,x~F@#@&9W~EUObV~Dk 2K0~GMPk~x,DdRhCT+jby@#@&EjoAAA==^#~@%> 
				<%#@~^RQAAAA==~iYGak1'YMkscDk`EYKwr^r##@#@&diddGrdw^lHKGwr^{:rNvPWak^S8~ 0b@#@&7di7MBIAAA==^#~@%>
                          <tr bgcolor=ffffff>
                            <td height="23" align=center bgcolor="#Ffffff" >
					        <%#@~^JwAAAA==r6POMks`Dk`r]+aVzZKxOn	YJ*b@!@*JJ,O4+	PCgwAAA==^#~@%>
					        <font color=red >已处理</font>
					        <%#@~^BAAAAA==n^/nqQEAAA==^#~@%>
					        <font color=red >未处理</font>
					        <%#@~^BgAAAA==n	N~b0JgIAAA==^#~@%>
							</td>
<td align="left" bgcolor="#ffffff">
·<a  class=middle href='E_ReadOpinion.asp?OpinionID=<%=#@~^DwAAAA==.k`E}wbxkKx&9J*4wQAAA==^#~@%>' target="_black"><%=#@~^GAAAAA==4D:sx1WN`Gr/aVCX:W2r1#SQkAAA==^#~@%>...</a></td>
 <td align=center bgcolor="#ffffff"><%=#@~^FwAAAA==zl.vDk`J`w9CYKr:Jbb,jgcAAA==^#~@%>-<%=#@~^GAAAAA==\KxO4`M/`rja[lD+Pks+Eb*P4wcAAA==^#~@%>-<%=#@~^FgAAAA==9mXcM/vJjaNmO+:kh+r#b~+wYAAA==^#~@%></td>
 </tr>
       <%#@~^MAAAAA==@#@&d7M/cHW7+gn6D@#@&dik~x,kP3~F@#@&disGWa@#@&d7kggAAA==^#~@%>
		<tr> 
	<td colspan="5" width="100%" align=center height=25 bgcolor="#ffffff">共搜索到 <%=#@~^BQAAAA==OKYC^JAIAAA==^#~@%> 条，当前第 <%=#@~^BgAAAA==\HwCT+YwIAAA==^#~@%>/<%=#@~^CAAAAA==\m62mo/NgMAAA==^#~@%> 页，每页 <%=#@~^CgAAAA==\HnCT+Uky3gMAAA==^#~@%> 条
	<%#@~^dgcAAA==@#@&d;MV{J2|rarxbWU{"+d;^YRmdwQ3+HAGD9'rP'PVnHhGD9@#@&hlLng+6Ojbyn'bUD`ctXhloO8bzhlL+UtGAUkyb_8@#@&hCL+Dwmon'rUD`cYKOl^OqbJD/ Kmon?b"#Q8@#@&@#@&b0,KlT+H+XYjr.+P@*qPDt+	@#@&nmon.+-xhlL+U4WS?r"e`KCT+H+XOUk"O8#@#@&IdwKxd+ch.rD+Pr@!l,mVmdd'(VmmVP4.0xBr~[,j.s,[PE'alL+{E,[~hlT+nM+7~[,JvPDkOs'B上r~[,nlTnjtKhUk"+~',J页v@*上一翻页@!&l@*PE@#@&I+d2Kxd+cAMkOPr@!l,m^C/k'8VmmV~4D+6xBrP[,i.V,[,J'wCL'qB,OkDVnxE第F页B@*页首@!Jl@*Pr@#@&+U9Pb0@#@&k6~HHnCoOq~@*P!,Otx@#@&K.+7{hlL+~x,HznmL+,O~q@#@&IndaWU/ SDrD+,J@!mP1slk/x4^l^V,tDW'EJPL~iD^PLPE[2CT+xJ,'PhDn-|nlLn,[~JE~DkO^+{B第J,[,KD\mnmon~LPJ页E@*上一页@!Jl@*,E@#@&x9Pr0@#@&@#@&r0,\lXwCL/@*xKmon1aD?r.+CnlT+U4WS?ryPO4x@#@&KlT+?b"n?4WSPxPKCT+jtKA?byn@#@&2Vdn@#@&KlTnUk"?4Wh,',\lXwCo/RKmo+U4WS?k.nM`hlT+H+aOUk"+ q#@#@&3U9PkW@#@&qWPhCT+jby?tKh,@!P8PPtx~Kmo+Ury?tKA~',F@#@&WW.~hlL+;GE	Yn.Ukynx8POW,KmonUk.+?4WS@#@&hlL+dkUV,'PvKlT+ZK;UYDUk"+QKmon1aYUk"nCnlLnUtGhUr.+b nmo+UtKA?byn@#@&kW~hloJk	3P@!@*~HHnmonPP4x@#@&"n/aWUdRh.rD+~J@!C,msm/k'4^l1VP4Dn0{BE~LPjMsPLPJL2Co'rP'PKCT+Jk	VPLPEv@*$J~',nCoJbxV,[,JT@!zm@*Pr@#@&+^/n@#@&I+k2W	/+c	.kD+,J@!A@*,r[~nmL+dkUV,[JD@!JA@*Pr@#@&+U9Pb0@#@&q6~nmonSbxV~{PHmanmo+k~Ptx,2akO~6W.@#@&H+XY@#@&@#@&kW~tX2lTn3F~@!'hloYaCoP~Y4+U@#@&1+XO{hlo~xPtXhlL+~Q,F@#@&"n/aWUdRh.rD+~J@!C,msm/k'4^l1VP4Dn0{BE~LPjMsPLPJL2Co'rP'PHnXYmnmL+,[~EEPYrO^+xB第r~LPH6D{nmo~[,J页v@*下一页@!z)@*r@#@&UN,k0@#@&@#@&b0,HC6KCT+dP@*~nmonj4Whjr.+MnmL1nXYUkyPD4+	@#@&nmonH6Y,xPhloj4WS?bynPM~hlL+gn6D?r"P_~q@#@&]+k2KxdRSDkD+,EP@!b~m^ldd{4Vm^3,tDWxBrPLPiDs~LPE[aCo'E~LPnCLY2lTn,[~rB,YkDVxB第r[~nmonOalo~[r页B@*页尾@!&)@*r@#@&In/2G	/nRS.kD+~E,@!l~^^ld/{8^l^0P4D+6'EEPLPiD^P'~r[wmL+{JPL~KlT+g+aY~',JvPDrY^+xv下rP[~Kmon?4GS?r.+,[Pr页B@*下一翻页@!zm@*E@#@&2U[,k0@#@&+^/+@#@&]+kwKxd+ AMkO+,E@!DD@*@!DNPCsboU'1n	YnMP(omKVK.'[s/sG2$~1WVk2l	'&,4nkTtD'q!T@*@!0GxD~mKVG.{D+[@*对不起，没有找到你要搜索的信件！@!zWW	O@*@!&DN@*@!zDD@*E@#@&d7d@#@&3U9Pq6@#@&M/R1sG/@#@&/nY~.k'UWD4k	o@#@&JBQCAA==^#~@%>
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
