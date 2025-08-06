<%@ Language=VBScript%>
<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="function.asp"-->
<!--#include file="char.inc"-->
<%
E_typeid=ChkRequest(request("E_typeid"),1)	'防注入
request_E_BigClassID=ChkRequest(Request.QueryString("E_BigClassID"),1)	'防注入
if E_typeid="" or request_E_BigClassID="" then
	Show_Err("未指定参数<br><br><a href='javascript:history.back()'>返回</a>")
	response.end
else
	if not IsNumeric(request_E_BigClassID) or not IsNumeric(E_typeid) then
		Show_Err("非法参数<br><br><a href='javascript:history.back()'>返回</a>")
		response.end
	else
		set rs=server.CreateObject("ADODB.RecordSet")
		rs.Source="select * from "& db_EC_BigClass_Table &" where E_BigClassID=" & request_E_BigClassID &" order by E_bigclassorder"
		rs.Open rs.Source,conn,1,1
		E_BigClassName=rs("E_BigClassName")
		E_bigclassorder=rs("E_bigclassorder")
		bigclasszs=rs("bigclasszs")
		rs.close
		set rs=nothing

		set rs=server.CreateObject("ADODB.RecordSet")
        rs.Source="select * from "& db_EC_BigClass_Table &" where E_typeid=" & E_typeid &" order by E_bigclassorder"
        rs.Open rs.Source,conn,1,1

		dim E_BigClassID
		E_BigClassID=ChkRequest(request("E_BigClassID"),1)	'防注入
		set rs=server.CreateObject("ADODB.RecordSet")
		rs.Source="select * from "& db_EC_SmallClass_Table &" where E_BigClassID="& E_BigClassID &" order by smallClassorder"
		rs.Open rs.Source,conn,1,1
		i=1

		Dim ArrayE_SmallClassID(), ArrayE_smallclassname(), ArrayE_smallclassview()
		ReDim ArrayE_SmallClassID(int(rs.RecordCount)), ArrayE_smallclassname(int(rs.RecordCount)), ArrayE_smallclassview(int(rs.RecordCount))

        if not rs.EOF then
		rseof=1
		end if
	 end if
while not rs.EOF

		abcount=rs.RecordCount
		E_smallclassview=rs("E_smallclassview")
		E_SmallClassID=rs("E_SmallClassID")
		
		E_smallclassname=trim(rs("E_smallclassname"))
		
		ArrayE_smallclassview(i)=E_smallclassview
		ArrayE_SmallClassID(i)=E_SmallClassID
		ArrayE_smallclassname(i)=E_smallclassname

		i=i+1
		
		rs.MoveNext
	    wend
		rs.close

		
		set rs2=server.CreateObject("ADODB.RecordSet")  '专题
		rs2.Source="select Top " & top_sp & " * from "& db_EC_Special_Table &"  order by SpecialID DESC "
		rs2.Open rs2.Source,conn,1,1
		rs2.close
		set rs2=nothing
		
		set rs4=server.CreateObject("Adodb.RecordSet")
		rs4.source="select * from "& db_EC_SmallClass_Table &" Where E_BigClassID=" & request_E_BigClassID &" order by SmallClassorder"
		rs4.open rs4.source,conn,1,1
		
		SmallClassCount=rs4.RecordCount
		
		for i=1 to SmallClassCount
			ArrayE_SmallClassID(i)=rs4("E_SmallClassID")
			ArrayE_smallclassname(i)=rs4("E_smallclassname")
			rs4.MoveNext
		next
		
		rs4.Close
		set rs4=nothing
		
		dim E_typename
		set rs5=server.CreateObject("ADODB.RecordSet")
		rs5.Source="select * from "& db_EC_Type_Table &" where E_typeid=" & E_typeid &" order by E_typeorder"
		rs5.Open rs5.Source,conn,1,1
		E_typename=rs5("E_typename")
		E_typeorder=rs5("E_typeorder")
		rs5.Close
		set rs5=nothing
end if
%>

<html><head>
<meta http-equiv="Content-Type" content="">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=E_BigClassName%>__<%=E_typename%>__<%=eChSys%></title>
<LINK href="news.css" rel="stylesheet">
<style type="text/css">
<!--
.STYLE1 {color: #FFFFFF}
-->
</style>
</head>
<body>
<%
dim E_BigClassName
E_BigClassID=ChkRequest(request("E_BigClassID"),1)	'防注入
set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from "& db_EC_BigClass_Table &" where E_BigClassID="& E_BigClassID &" order by E_bigclassorder"
rs.Open rs.Source,conn,1,1
E_BigClassName=rs("E_BigClassName")
bigclasszs=rs("bigclasszs")
rs.close
set rs=nothing
%>
<!--#include file="other_Top.asp"-->
<%if bigclasszs=5 then%>

<table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#F1F7FC">
  <tr>
    <td height="25">&nbsp;&nbsp;&nbsp;<img src="Images/arrow_dh_1.gif" width="7" height="7">&nbsp;&nbsp;&nbsp;<font class=m_tittle>&nbsp;</font>栏目导航&nbsp;&nbsp;<a href="./" class="daohang">网站首页</a>&gt;&gt;<a href=E_Type.asp?E_typeid=<%=E_typeid%> class="daohang"><%=E_typename%></a>
    <% Response.Write "&gt;&gt;" & E_BigClassName & "" %></td>
  </tr>
  <tr>
    <td height="66" valign="top"><table width="99%" border="0" align="center" cellpadding="0" cellspacing="8" bgcolor="#FFFFFF">
      <tr>
        <td height="44" valign="top">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<%set rsnosmall=server.CreateObject("ADODB.RecordSet")
			if uselevel=1 then
			    
				if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then
				rsnosmall.Source="select top 1 * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("key")="" then
					rsnosmall.Source="select top 1 * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("key")="selfreg" then
					if Request.cookies(eChsys)("reglevel")=3 then
						rsnosmall.Source="select top 1 * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
					end if
				if Request.cookies(eChsys)("reglevel")=2 then
					rsnosmall.Source="select top 1 * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("reglevel")=1 then
					rsnosmall.Source="select top 1 * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				
			end if
			else
			rsnosmall.Source="select top 1 * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1) order by newsid DESC"
			end if
				rsnosmall.Open rsnosmall.Source,conn,1,1
			if not rsnosmall.EOF then
%>
        <tr>
          <td><div align="center">
            <table width="100%" border="0" cellpadding="3" cellspacing="0">
              <tr>
                <td width="67%"><%=rsnosmall("content")%> </td>
                </tr>
            </table>
          </div></td>
        </tr>
      <%
		else 
			Response.Write "<tr><td> <div align='center'><table width='98%' border='0' cellpadding='3' cellspacing='0'><tr><td width='62%'>&nbsp;&nbsp;<img src='IMAGES/006.gif' width='10' height='10'> 暂无本栏信息</td></tr></table></div></td></tr>"
		%>
      <%end if
		rsnosmall.close
		set rsnosmall=nothing
		%>
    </table>		
		
		</td>
      </tr>
    </table></td>
  </tr>
</table>
<%else%>
<table width="80%" height="400" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr valign="top">
    <td width="210" rowspan="0" bgcolor="F1F7FC">
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="50" align="center" background="IMAGES/type_dh.gif"><img src="Images/arrow_dh_1.gif" width="7" height="7"><font color="#ffffff" class="yellow_title">&nbsp;</font><span class="yellow_title"><%=E_BigClassName%>·大类 <%=E_typeorder%>-<%=E_bigclassorder%></span></td>
      </tr>
      <%
set rs6=server.CreateObject("ADODB.RecordSet")
rs6.Source="select * from "& db_EC_SmallClass_Table &" where E_BigClassID="&request_E_BigClassID&" order by SmallClassorder"
rs6.Open rs6.Source,conn,1,1
do while not rs6.eof%>
      <tr>
        <td align=center valign="middle"><a class=class href="E_SmallClass.asp?E_typeid=<%=E_typeid%>&E_BigClassID=<%=rs6("E_BigClassID")%>&E_SmallClassID=<%=rs6("E_SmallClassID")%>"><%=rs6("E_smallclassname")%></a> </td>
      </tr>
      <%rs6.movenext
loop
rs6.Close
set rs6=nothing
%>
      <tr>
        <td height="50" align=center background="Images/type_dh.gif"><span class="yellow_title"><img src="Images/arrow_dh_1.gif" width="7" height="7"><font color="#ffffff" class="yellow_title">&nbsp;</font>&nbsp;最新图文</span> </td>
      </tr>
      <%
set rs3=server.CreateObject("ADODB.RecordSet")
	if uselevel=1 then
		if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then
		rs3.Source ="select top " & top_img & " * from "& db_EC_News_Table &" where picnews=1 and checkked=1 and E_typeid="&E_typeid&" and picname is not null order by NewsID DESC"
		end if
		if Request.cookies(eChsys)("key")="" then
			rs3.Source ="select top " & top_img & " * from "& db_EC_News_Table &" where picnews=1 and checkked=1 and newslevel=0 and E_typeid="&E_typeid&" and picname is not null order by NewsID DESC"
		end if
		if Request.cookies(eChsys)("key")="selfreg" then
			if Request.cookies(eChsys)("reglevel")=3 then
				rs3.Source ="select top " & top_img & " * from "& db_EC_News_Table &" where picnews=1 and checkked=1  and E_typeid="&E_typeid&" and picname is not null order by NewsID DESC"
			end if
			if Request.cookies(eChsys)("reglevel")=2 then
				rs3.Source ="select top " & top_img & " * from "& db_EC_News_Table &" where picnews=1 and checkked=1  and E_typeid="&E_typeid&" and picname is not null order by NewsID DESC"
			end if
			if Request.cookies(eChsys)("reglevel")=1 then
				rs3.Source ="select top " & top_img & " * from "& db_EC_News_Table &" where picnews=1 and checkked=1  and E_typeid="&E_typeid&" and picname is not null order by NewsID DESC"
			end if
		end if
	else
		rs3.Source ="select top " & top_img & " * from "& db_EC_News_Table &" where picnews=1 and checkked=1 and E_typeid="&E_typeid&" and picname is not null order by NewsID DESC"
	end if
rs3.Open rs3.Source,conn,1,1
	if not rs3.EOF then
		while not rs3.EOF
			fileExt=lcase(getFileExtName(rs3("picname")))
			Content=htmlencode4(rs3("Content"))
			%>
      <tr>
       <td style="WORD-WRAP: break-word" align="center">
		<a class=middle href="E_ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=htmlencode4(rs3("title"))%>">
          <%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
          <img  src="<%=FileUploadPath & rs3("picname")%>" width="80" height="80" border=0 align="center">
          <%end if%>
          <%if fileext="swf" then%>
          <object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="80" height="80" border=0 align="center">
            <param name=movie value="<%=FileUploadPath & rs3("picname")%>">
            <param name=quality value=high>
            <param name='Play' value='-1'>
            <param name='Loop' value='0'>
            <param name='Menu' value='-1'>
            <embed src="<%=FileUploadPath & rs3("picname")%>" width="80" height="80" border=0 align="center" pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
          </object>
          <%end if%>
		  </a>
		</td>
      </tr>
	  <tr>
		<td style="WORD-WRAP: break-word" align="center"><a class=middle href="E_ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=htmlencode4(rs3("title"))%>"><%=CutStr(htmlencode4(rs3("title")),24)%></a>
		</td>
	 </tr>
      <%
rs3.MoveNext
	wend
	    else
		Response.Write "<tr><td width=100% align=center height=18 >暂无</td></tr>"
	    end if
	rs3.close
	set rs3=nothing
%>
    </table>	
</td>
    <td width="570" rowspan="0">
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="25" background="Images/dh_bg.gif">&nbsp;&nbsp;&nbsp;<img src="Images/arrow_dh_1.gif" width="7" height="7">&nbsp;&nbsp;&nbsp;<font class=m_tittle>&nbsp;</font>栏目导航&nbsp;&nbsp;<a href="./" class="daohang">网站首页</a>&gt;&gt;<a href=E_Type.asp?E_typeid=<%=E_typeid%> class="daohang"><%=E_typename%></a><% Response.Write "&gt;&gt;" & E_BigClassName & "" %></td>
      </tr>
	</table>
	<!--
	<table width="100%" height="32" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="24"><div align="center">
              <script language=javascript src=./zongg/ad.asp?i=14></script>
            </div></td>
          </tr>
    </table>
	-->
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="25" background="Images/dh_bg.gif">&nbsp;&nbsp;&nbsp;&nbsp;<img src="Images/arrow_dh_1.gif" width="7" height="7">&nbsp;&nbsp;<font class=body> <%=E_BigClassName%>[本栏]</font></td>
      </tr>
	  </table>
        <!--无小类文章区开始-->
        <% select case bigclasszs%>
        <% case 1 %>
        <!--图片模版-->
        <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber5">
          <%set rsnosmall=server.CreateObject("ADODB.RecordSet")
			if uselevel=1 then		    
				if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then
				rsnosmall.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("key")="" then
					rsnosmall.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("key")="selfreg" then
					if Request.cookies(eChsys)("reglevel")=3 then
						rsnosmall.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
					end if
				if Request.cookies(eChsys)("reglevel")=2 then
					rsnosmall.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("reglevel")=1 then
					rsnosmall.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				
			end if
			else
			rsnosmall.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1) order by newsid DESC"
			end if
				rsnosmall.Open rsnosmall.Source,conn,1,1
				%>
          <tr>
            <td><table width="98%" border="0" cellspacing="0" cellpadding="3" align="center">
              <!--图片换行显示1-->
              <%
				if not rsnosmall.EOF then
				do while not rsnosmall.EOF
				%>
              <tr>
                <%
					fileExt=lcase(getFileExtName(rsnosmall("picname")))
					Content=htmlencode4(rsnosmall("Content"))
					content=replace(content,"[---分页---]","")
					%>
                <td width=25% align=center valign="top" style="table-layout:fixed; word-break:break-all"><%if not rsnosmall.EOF then%>
                        <div align="center">
                          <% content=htmlencode4(rsnosmall("Content"))
					content=replace(content,"[---分页---]","")%>
                          <a class=middle href="E_ReadNews.asp?NewsID=<%=rsnosmall("NewsID")%>" target=_blank title="<%=CutStr(nohtml(Content),150)%>...">
                          <%if   rsnosmall("picname")=("") then%>
                          <img  src="IMAGES/flashorno.gif" width="110" height="80" border=1 style=border-color:#000000 align=top>
                          <%else%>
                          <%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
                          <img  src="<%=FileUploadPath & rsnosmall("picname")%>" width="110" height="80" border=1 style=border-color:#000000 align=top>
                          <%end if%>
                          <%if fileext="swf" then%>
                          <object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="110" height="80" border=0 >
                            <param name=movie value="<%=FileUploadPath & rsnosmall("picname")%>">
                            <param name=quality value=high>
                            <param name='Play' value='-1'>
                            <param name='Loop' value='0'>
                            <param name='Menu' value='-1'>
                            <embed src="<%=FileUploadPath & rsnosmall("picname")%>" width="110" height="80" border=0  pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
                          </object>
                          <%end if%>
                          <%end if%>
                          </a> <br>
                          <br>
                      <a class=middle href="E_ReadNews.asp?NewsID=<%=rsnosmall("NewsID")%>" target=_blank title="<%=rsnosmall("title")%>"><%=CutStr(rsnosmall("title"),18)%></a> </div></td>
                <%rsnosmall.movenext   
					end if %>
                <!--图片换行显示2-->
                <td width=25% align=center valign="top" style="table-layout:fixed; word-break:break-all"><%if not rsnosmall.EOF then%>
                        <div align="center">
                          <%
					content=htmlencode4(rsnosmall("Content"))
					content=replace(content,"[---分页---]","")%>
                          <a class=middle href="E_ReadNews.asp?NewsID=<%=rsnosmall("NewsID")%>" target=_blank title="<%=CutStr(nohtml(Content),150)%>">
                          <%if rsnosmall("picname")=("") then%>
                          <img  src="IMAGES/flashorno.gif" width="110" height="80" border=1 style=border-color:#000000 align=top>
                          <%else%>
                          <%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
                          <img  src="<%=FileUploadPath & rsnosmall("picname")%>" width="110" height="80" border=1 style=border-color:#000000 align=top>
                          <%end if%>
                          <%if fileext="swf" then%>
                          <object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="110" height="80" border=0 >
                            <param name=movie value="<%=FileUploadPath & rsnosmall("picname")%>">
                            <param name=quality value=high>
                            <param name='Play' value='-1'>
                            <param name='Loop' value='0'>
                            <param name='Menu' value='-1'>
                            <embed src="<%=FileUploadPath & rsnosmall("picname")%>" width="110" height="80" border=0  pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
                          </object>
                          <%end if%>
                          <%end if%>
                          </a><br>
                          <br>
                      <a class=middle href="E_ReadNews.asp?NewsID=<%=rsnosmall("NewsID")%>" target=_blank title="<%=rsnosmall("title")%>"><%=CutStr(rsnosmall("title"),18)%></a> </div></td>
                <%rsnosmall.movenext   
					end if %>
                <!--图片换行显示3-->
                <td width=25% align=center valign="top" style="table-layout:fixed; word-break:break-all"><%if not rsnosmall.EOF then%>
                        <div align="center">
                          <%
					content=htmlencode4(rsnosmall("Content"))
					content=replace(content,"[---分页---]","")%>
                          <a class=middle href="E_ReadNews.asp?NewsID=<%=rsnosmall("NewsID")%>" target=_blank title="<%=CutStr(nohtml(Content),150)%>">
                          <%if rsnosmall("picname")=("") then%>
                          <img  src="IMAGES/flashorno.gif" width="110" height="80" border=1 style=border-color:#000000 align=top>
                          <%else%>
                          <%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
                          <img  src="<%=FileUploadPath & rsnosmall("picname")%>" width="110" height="80" border=1 style=border-color:#000000 align=top>
                          <%end if%>
                          <%if fileext="swf" then%>
                          <object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="110" height="80" border=0 >
                            <param name=movie value="<%=FileUploadPath & rsnosmall("picname")%>">
                            <param name=quality value=high>
                            <param name='Play' value='-1'>
                            <param name='Loop' value='0'>
                            <param name='Menu' value='-1'>
                            <embed src="<%=FileUploadPath & rsnosmall("picname")%>" width="110" height="80" border=0  pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
                          </object>
                          <%end if%>
                          <%end if%>
                          </a> <br>
                          <br>
                      <a class=middle href="E_ReadNews.asp?NewsID=<%=rsnosmall("NewsID")%>" target=_blank title="<%=rsnosmall("title")%>"><%=CutStr(rsnosmall("title"),18)%></a> </div></td>
                <%rsnosmall.movenext   
				end if %>
                <!--图片换行显示4-->
                <td width=25% align=center valign="top" style="table-layout:fixed; word-break:break-all"><%if not rsnosmall.EOF then%>
                        <div align="center">
                          <%content=rsnosmall("Content")
					content=replace(content,"[---分页---]","")%>
                          <a class=middle href="E_ReadNews.asp?NewsID=<%=rsnosmall("NewsID")%>" target=_blank title="<%=CutStr(nohtml(Content),150)%>...">
                          <%if rsnosmall("picname")=("") then%>
                          <img  src="IMAGES/flashorno.gif" width="110" height="80" border=1 style=border-color:#000000 align=top>
                          <%else%>
                          <%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
                          <img  src="<%=FileUploadPath & rsnosmall("picname")%>" width="110" height="80" border=1 style=border-color:#000000 align=top>
                          <%end if%>
                          <%if fileext="swf" then%>
                          <object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="110" height="80" border=0 >
                            <param name=movie value="<%=FileUploadPath & rsnosmall("picname")%>">
                            <param name=quality value=high>
                            <param name='Play' value='-1'>
                            <param name='Loop' value='0'>
                            <param name='Menu' value='-1'>
                            <embed src="<%=FileUploadPath & rsnosmall("picname")%>" width="110" height="80" border=0  pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
                          </object>
                          <%end if%>
                          <%end if%>
                          </a> <br>
                          <br>
                      <a class=middle href="E_ReadNews.asp?NewsID=<%=rsnosmall("NewsID")%>" target=_blank title="<%=rsnosmall("title")%>"><%=CutStr(rsnosmall("title"),18)%></a> </div></td>
                <%rsnosmall.movenext   
					end if %>
              </tr>
              <!--图片换行结束-->
              <%loop%>
              <%
			else
			Response.Write "<tr><td> <div align='center'><table width='98%' border='0' cellpadding='3' cellspacing='0'><tr><td width='62%'>&nbsp;&nbsp;<img src='IMAGES/006.gif' width='10' height='10'> 暂无本栏信息</td></tr></table></div></td></tr>"
		    end if
			rsnosmall.close
			set rsnosmall=nothing
			%>
            </table>
		</table>
        <!--图片模版-->
        <% case 2 %>
        <!--新闻模版-->
<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<%set rsnosmall=server.CreateObject("ADODB.RecordSet")
			if uselevel=1 then
			    
				if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then
				rsnosmall.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("key")="" then
					rsnosmall.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("key")="selfreg" then
					if Request.cookies(eChsys)("reglevel")=3 then
						rsnosmall.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
					end if
				if Request.cookies(eChsys)("reglevel")=2 then
					rsnosmall.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("reglevel")=1 then
					rsnosmall.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				
			end if
			else
			rsnosmall.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1) order by newsid DESC"
			end if
				rsnosmall.Open rsnosmall.Source,conn,1,1
			if not rsnosmall.EOF then
			while not rsnosmall.EOF
				       if 1 then 
                                     //if showyear=1 then
					newsurl="E_ReadNews.asp?NewsID=" & rsnosmall("NewsID")
					newswwwurl=rsnosmall("titleface")
					datetime="<font class=middle>(" & year(rsnosmall("UpdateTime"))  &"年"& Month(rsnosmall("UpdateTime"))  &"月"& Day(rsnosmall("UpdateTime")) &"日)</font>"
				else
					newsurl="E_ReadNews.asp?NewsID=" & rsnosmall("NewsID")
					newswwwurl=rsnosmall("titleface")
				datetime="<font class=middle>(" & year(rsnosmall("UpdateTime"))  &"年"& Month(rsnosmall("UpdateTime"))  &"月"& Day(rsnosmall("UpdateTime")) &"日)</font>"	
                          datetime="<font class=middle>("& Month(rsnosmall("UpdateTime"))  &"月"& Day(rsnosmall("UpdateTime")) &"日)</font>"
				end if
				if rsnosmall("picnews")=1 then
					img="<font class=pic_word>[图]</font>"
				else
					img=""
				end if
					title=trim(rsnosmall("title"))
					title=replace(title,"<br>","")
		%>
        <tr>
          <td><div align="center">
            <table width="100%" border="0" cellpadding="3" cellspacing="0">
              <tr>
                <td width="67%">&nbsp;&nbsp;<img src="IMAGES/006.gif" width="10" height="10"> <%=img%> <a class=middle href="<%if rsnosmall("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" title="<%=rsnosmall("title")%>" target="_blank"> <<%=rsnosmall("titletype")%>><font color="<%=rsnosmall("titlecolor")%>"> <%=CutStr(title,46)%></font></<%=rsnosmall("titletype")%>></a>
                      <!--标题后评论提示-->
                      <% if rsnosmall("titlesize")>=1 then %>
                      <A class=middle HREF="<%=path%>review.asp?NewsID=<%=rsnosmall("NewsID")%>" target="_blank" ><font color=red><b>评</b></font></A>
                      <%end if %>
                      <!--标题后评论提示-->
                </td>
                <td width="18%" align="left"><div align="right">
                  <%if showtime="1" then%>
                  <%=datetime%>
                  <%end if%>
                  &nbsp; </div></td>
                <td width="15%" align="left"><%if year(rsnosmall("updatetime"))=year(date()) and month(rsnosmall("updatetime"))=month(date()) and day(rsnosmall("updatetime"))=day(date()) then%>
                      <img src="images/new.gif">
                      <%end if%>
                      <%if rsnosmall("goodnews")="1" then%>
                      <img src="images/g.gif" >
                      <%end if%>
                </td>
              </tr>
            </table>
          </div></td>
        </tr>
      <%
		rsnosmall.MoveNext
		wend
		else 
			Response.Write "<tr><td> <div align='center'><table width='98%' border='0' cellpadding='3' cellspacing='0'><tr><td width='62%'>&nbsp;&nbsp;<img src='IMAGES/006.gif' width='10' height='10'> 暂无本栏信息</td></tr></table></div></td></tr>"
		%>
      <%end if
		rsnosmall.close
		set rsnosmall=nothing
		%>
    </table>
        <% case 3 %>
        <!--网址模版-->
        <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber5">
          <%set rsnosmall=server.CreateObject("ADODB.RecordSet")
		if uselevel=1 then
			if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then
				rsnosmall.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("key")="" then
					rsnosmall.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("key")="selfreg" then
					if Request.cookies(eChsys)("reglevel")=3 then
						rsnosmall.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
					end if
				if Request.cookies(eChsys)("reglevel")=2 then
					rsnosmall.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("reglevel")=1 then
					rsnosmall.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
			end if
			else
			rsnosmall.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1) order by newsid DESC"
			end if
			rsnosmall.Open rsnosmall.Source,conn,1,1
		%>
          <tr>
            <td><table width="98%" border="1" cellspacing="0" cellpadding="3" bordercolorlight=#cccccc  bordercolordark=#ffffff align="center">
                <!--网址换行显示1-->
                <%
					if not rsnosmall.EOF then
					do while not rsnosmall.EOF%>
                <tr>
                  <%
						Content=htmlencode4(rsnosmall("Content"))
					%>
                  <td width=25% align=center valign="middle"><%if not rsnosmall.EOF then%>
                      <div align="center"> <a class=middle href="<%=rsnosmall("Original")%>" target=_blank title="<%=CutStr(nohtml(rsnosmall("Content")),80)%>"><%=CutStr(rsnosmall("title"),30)%></a> </div></td>
                  <%rsnosmall.movenext   
					end if %>
                  <!--网址换行显示2-->
                  <td width=25% align=center valign="middle"><%if not rsnosmall.EOF then%>
                      <div align="center"> <a class=middle href="<%=rsnosmall("Original")%>" target=_blank title="<%=CutStr(nohtml(rsnosmall("Content")),80)%>"><%=CutStr(rsnosmall("title"),30)%></a> </div></td>
                  <%rsnosmall.movenext   
					end if %>
                </tr>
                <!--网址换行结束-->
                <%loop%>
                <%
			else
			Response.Write "<tr><td> <div align='center'><table width='98%' border='0' cellpadding='3' cellspacing='0'><tr><td width='62%'>&nbsp;&nbsp;<img src='IMAGES/006.gif' width='10' height='10'> 暂无本栏信息</td></tr></table></div></td></tr>"
		    end if
			rsnosmall.close
			set rsnosmall=nothing
			%>
              </table>
          </tr>
          <tr>
            <td>  
          </tr>
        </table>
      <!--网址模版-->
        <% case 4 %>
        <!--软件模版-->
        <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber5">
          <%set rsnosmall=server.CreateObject("ADODB.RecordSet")
			if uselevel=1 then
				if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then
				rsnosmall.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("key")="" then
					rsnosmall.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("key")="selfreg" then
					if Request.cookies(eChsys)("reglevel")=3 then
						rsnosmall.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
					end if
				if Request.cookies(eChsys)("reglevel")=2 then
					rsnosmall.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("reglevel")=1 then
					rsnosmall.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				end if
				else
					rsnosmall.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1) order by newsid DESC"
				end if
			rsnosmall.Open rsnosmall.Source,conn,1,1
			    if not rsnosmall.EOF then
				while not rsnosmall.EOF
					    if 1 then
                                            //if showyear=1 then
						newsurl="E_ReadNews.asp?NewsID=" & rsnosmall("NewsID")
						newswwwurl=rsnosmall("titleface")
						datetime="<font class=middle>(" & year(rsnosmall("UpdateTime"))  &"年"& Month(rsnosmall("UpdateTime"))  &"月"& Day(rsnosmall("UpdateTime")) &"日)</font>"
					else
						newsurl="E_ReadNews.asp?NewsID=" & rsnosmall("NewsID")
						newswwwurl=rsnosmall("titleface")
						datetime="<font class=middle>("& Month(rsnosmall("UpdateTime"))  &"月"& Day(rsnosmall("UpdateTime")) &"日)</font>"
					end if
					if rsnosmall("picnews")=1 then
						img="<font class=pic_word>[图]</font>"
					else
						img=""
					end if
						title=trim(rsnosmall("title"))
						title=replace(title,"<br>","")
			%>
          <tr>
            <%
		fileExt=lcase(getFileExtName(rsnosmall("picname")))
		Content=htmlencode4(rsnosmall("Content"))
		content=replace(content,"[---分页---]","")%>
            <td><div align="center">
                <table width="100%" border="0" cellpadding="0" cellspacing="0">
                  <tr bgcolor="#EFEFEF">
                    <td colspan="2">&nbsp;<img src="images/news_img.gif" width="9" height="9"> <a class=middle href="<%if rsnosmall("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" title="<%=htmlencode4(title)%>" target="_blank"> <font color="<%=rsnosmall("titlecolor")%>"><strong> <%=CutStr(title,40)%> </strong></font> </a> </td>
                    <td width="17%" align="right" bgcolor="#EFEFEF"><%if showtime="1" then%>
                        <%=datetime%>
                        <%end if%>
                    &nbsp;</td>
                    <td width="16%" align="center"><%if year(rsnosmall("updatetime"))=year(date()) and month(rsnosmall("updatetime"))=month(date()) and day(rsnosmall("updatetime"))=day(date()) then%>
                        <img src="images/new.gif">
                        <%end if%>
                        <%if rsnosmall("goodnews")="1" then%>
                        <img src="images/g.gif" >
                        <%end if%>
                    </td>
                  </tr>
                  <tr>
                    <td width="39%"><a class=middle href="<%if rsnosmall("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" target=_blank title="<%=htmlencode4(rsnosmall("title"))%>">
                      <%if   rsnosmall("picname")=("")   then%>
                      <img  src="IMAGES/softno.gif" width="65" height="65" border=0 align="left">
                      <%else%>
                      <%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
                      <img  src="<%=FileUploadPath & rsnosmall("picname")%>" width="65" height="65" border=0 align="left">
                      <%end if%>
                      <%if fileext="swf" then%>
                      <object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="65" height="65" border=0 >
                        <param name=movie value="<%=FileUploadPath & rsnosmall("picname")%>">
                        <param name=quality value=high>
                        <param name='Play' value='-1'>
                        <param name='Loop' value='0'>
                        <param name='Menu' value='-1'>
                        <embed src="<%=FileUploadPath & rsnosmall("picname")%>" width="6" 5height="65" border=0  pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
                      </object>
                      <%end if%>
                      <%end if%>
                    </a> </td>
                    <td colspan="3" valign="top"><a class=middle href="<%if rsnosmall("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" target="_blank" title="<%=htmlencode4(title)%>"><%=CutStr(nohtml(rsnosmall("Content")),250)%></a> </td>
                  </tr>
              </table>
            </div></td>
          </tr>
          <%
		rsnosmall.MoveNext
		wend
		%>
          <%
			else
			Response.Write "<tr><td> <div align='center'><table width='98%' border='0' cellpadding='3' cellspacing='0'><tr><td width='62%'>&nbsp;&nbsp;<img src='IMAGES/006.gif' width='10' height='10'> 暂无本栏信息</td></tr></table></div></td></tr>"
		    end if
			rsnosmall.close
			set rsnosmall=nothing
			%>
        </table>
      <!--软件模版-->
        <%end select%>
        <!--无小类文章区结束-->
        <TABLE cellSpacing=0 cellPadding=3 width="100%" align=center border=0>
          <tr>
            <td width=100% height="19" align=right><a class=class href='E_nosmallClass.asp?E_typeid=<%=E_typeid%>&E_BigClassID=<%=E_BigClassID%>'><img src="images/more.gif" border="0" alt="更多<%=BigClassName%>"></a>&nbsp;&nbsp;</td>
          </tr>
        </table>
      <!--模版开始-->
        <% select case bigclasszs%>
        <% case 1 %>
        <!--图片模版-->
        <%
	    if rseof=1 then
		for i=1 to abcount
		E_SmallClassID=ArrayE_SmallClassID(i)
		E_SmallClassName=ArrayE_SmallClassName(i)
		E_smallClassview=ArrayE_smallClassview(i)
		if ArrayE_smallClassview(i)=1 then 
		%>
        <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber5">
          <tr>
            <td height="25" background="Images/dh_bg.gif">&nbsp;&nbsp;&nbsp;&nbsp;<img src="Images/arrow_dh_1.gif" width="7" height="7">&nbsp;&nbsp;<font class=body><%=E_SmallClassName%></font></td>
          </tr>
          <%set rs3=server.CreateObject("ADODB.RecordSet")
	if uselevel=1 then
		if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then
		rs3.Source="select top "& bigclassshownum &" * from "& db_EC_News_Table &" where (E_SmallClassID="& E_SmallClassID &" and checkked=1) order by newsid DESC"
		end if
		if Request.cookies(eChsys)("key")="" then
		rs3.Source="select top "& bigclassshownum &" * from "& db_EC_News_Table &" where (E_SmallClassID="& E_SmallClassID &" and checkked=1 ) order by newsid 	DESC"
		end if
		if Request.cookies(eChsys)("key")="selfreg" then
		if Request.cookies(eChsys)("reglevel")=3 then
		rs3.Source="select top "& bigclassshownum &" * from "& db_EC_News_Table &" where (E_SmallClassID="& E_SmallClassID &" and checkked=1 ) order by newsid DESC"
		end if
		if Request.cookies(eChsys)("reglevel")=2 then
		rs3.Source="select top "& bigclassshownum &" * from "& db_EC_News_Table &" where (E_SmallClassID="& E_SmallClassID &" and checkked=1 ) order by newsid DESC"
		end if
		if Request.cookies(eChsys)("reglevel")=1 then
		rs3.Source="select top "& bigclassshownum &" * from "& db_EC_News_Table &" where (E_SmallClassID="& E_SmallClassID &" and checkked=1 ) order by newsid DESC"
		end if
		end if
	else
	rs3.Source="select top "& bigclassshownum &" * from "& db_EC_News_Table &" where (E_SmallClassID="& E_SmallClassID &" and checkked=1) order by newsid DESC"
	end if
	rs3.Open rs3.Source,conn,1,1
	%>
          <tr>
            <td><table width="98%" border="0" cellspacing="0" cellpadding="3" align="center">
                <!--图片换行显示1-->
                <%
				if not rs3.EOF then
				do while not rs3.EOF
				%>
                <tr>
                  <%
					fileExt=lcase(getFileExtName(rs3("picname")))
					Content=htmlencode4(rs3("Content"))
					content=replace(content,"[---分页---]","")
					%>
                  <td width=25% align=center valign="top" style="table-layout:fixed; word-break:break-all"><%if not rs3.EOF then%>
                      <div align="center">
                        <% content=htmlencode4(rs3("Content"))
					content=replace(content,"[---分页---]","")%>
                        <a class=middle href="E_ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=CutStr(nohtml(Content),150)%>...">
                        <%if   rs3("picname")=("") then%>
                        <img  src="IMAGES/flashorno.gif" width="110" height="80" border=1 style=border-color:#000000 align=top>
                        <%else%>
                        <%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
                        <img  src="<%=FileUploadPath & rs3("picname")%>" width="110" height="80" border=1 style=border-color:#000000 align=top>
                        <%end if%>
                        <%if fileext="swf" then%>
                        <object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="110" height="80" border=0 >
                          <param name=movie value="<%=FileUploadPath & rs3("picname")%>">
                          <param name=quality value=high>
                          <param name='Play' value='-1'>
                          <param name='Loop' value='0'>
                          <param name='Menu' value='-1'>
                          <embed src="<%=FileUploadPath & rs3("picname")%>" width="110" height="80" border=0  pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
                        </object>
                        <%end if%>
                        <%end if%>
                        </a> <br>
                        <br>
                    <a class=middle href="E_ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=rs3("title")%>"><%=CutStr(rs3("title"),18)%></a> </div></td>
                  <%rs3.movenext   
					end if %>
                  <!--图片换行显示2-->
                  <td width=25% align=center valign="top" style="table-layout:fixed; word-break:break-all"><%if not rs3.EOF then%>
                      <div align="center">
                        <%
					content=htmlencode4(rs3("Content"))
					content=replace(content,"[---分页---]","")%>
                        <a class=middle href="E_ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=CutStr(nohtml(Content),150)%>">
                        <%if rs3("picname")=("") then%>
                        <img  src="IMAGES/flashorno.gif" width="110" height="80" border=1 style=border-color:#000000 align=top>
                        <%else%>
                        <%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
                        <img  src="<%=FileUploadPath & rs3("picname")%>" width="110" height="80" border=1 style=border-color:#000000 align=top>
                        <%end if%>
                        <%if fileext="swf" then%>
                        <object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="110" height="80" border=0 >
                          <param name=movie value="<%=FileUploadPath & rs3("picname")%>">
                          <param name=quality value=high>
                          <param name='Play' value='-1'>
                          <param name='Loop' value='0'>
                          <param name='Menu' value='-1'>
                          <embed src="<%=FileUploadPath & rs3("picname")%>" width="110" height="80" border=0  pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
                        </object>
                        <%end if%>
                        <%end if%>
                        </a><br>
                        <br>
                    <a class=middle href="E_ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=rs3("title")%>"><%=CutStr(rs3("title"),18)%></a> </div></td>
                  <%rs3.movenext   
					end if %>
                  <!--图片换行显示3-->
                  <td width=25% align=center valign="top" style="table-layout:fixed; word-break:break-all"><%if not rs3.EOF then%>
                      <div align="center">
                        <%
					content=htmlencode4(rs3("Content"))
					content=replace(content,"[---分页---]","")%>
                        <a class=middle href="E_ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=CutStr(nohtml(Content),150)%>">
                        <%if rs3("picname")=("") then%>
                        <img  src="IMAGES/flashorno.gif" width="110" height="80" border=1 style=border-color:#000000 align=top>
                        <%else%>
                        <%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
                        <img  src="<%=FileUploadPath & rs3("picname")%>" width="110" height="80" border=1 style=border-color:#000000 align=top>
                        <%end if%>
                        <%if fileext="swf" then%>
                        <object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="110" height="80" border=0 >
                          <param name=movie value="<%=FileUploadPath & rs3("picname")%>">
                          <param name=quality value=high>
                          <param name='Play' value='-1'>
                          <param name='Loop' value='0'>
                          <param name='Menu' value='-1'>
                          <embed src="<%=FileUploadPath & rs3("picname")%>" width="110" height="80" border=0  pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
                        </object>
                        <%end if%>
                        <%end if%>
                        </a> <br>
                        <br>
                    <a class=middle href="E_ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=rs3("title")%>"><%=CutStr(rs3("title"),18)%></a> </div></td>
                  <%rs3.movenext   
				end if %>
                  <!--图片换行显示4-->
                  <td width=25% align=center valign="top" style="table-layout:fixed; word-break:break-all"><%if not rs3.EOF then%>
                      <div align="center">
                        <%content=rs3("Content")
					content=replace(content,"[---分页---]","")%>
                        <a class=middle href="E_ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=CutStr(nohtml(Content),150)%>...">
                        <%if rs3("picname")=("") then%>
                        <img  src="IMAGES/flashorno.gif" width="110" height="80" border=1 style=border-color:#000000 align=top>
                        <%else%>
                        <%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
                        <img  src="<%=FileUploadPath & rs3("picname")%>" width="110" height="80" border=1 style=border-color:#000000 align=top>
                        <%end if%>
                        <%if fileext="swf" then%>
                        <object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="110" height="80" border=0 >
                          <param name=movie value="<%=FileUploadPath & rs3("picname")%>">
                          <param name=quality value=high>
                          <param name='Play' value='-1'>
                          <param name='Loop' value='0'>
                          <param name='Menu' value='-1'>
                          <embed src="<%=FileUploadPath & rs3("picname")%>" width="110" height="80" border=0  pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
                        </object>
                        <%end if%>
                        <%end if%>
                        </a> <br>
                        <br>
                    <a class=middle href="E_ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=rs3("title")%>"><%=CutStr(rs3("title"),18)%></a> </div></td>
                  <%rs3.movenext   
					end if %>
                </tr>
                <!--图片换行结束-->
                <%loop
				rs3.Close
				set rs3=nothing
				%>
                <%
			else
			Response.Write "<tr><td> <div align='center'><table width='98%' border='0' cellpadding='3' cellspacing='0'><tr><td width='62%'>&nbsp;&nbsp;<img src='IMAGES/006.gif' width='10' height='10'> 暂无信息</td></tr></table></div></td></tr>"
		    end if
			%>
              </table>
                <TABLE cellSpacing=0 cellPadding=3 width="100%" align=center border=0>
                  <tr>
                    <td width=100% height="19" align=right><a class=class href='E_SmallClass.asp?E_typeid=<%=E_typeid%>&E_BigClassID=<%=E_BigClassID%>&E_SmallClassID=<%=E_SmallClassID%>'><img src="images/more.gif" border="0" alt="更多<%=E_SmallClassName%>"></a>&nbsp;&nbsp;</td>
                  </tr>
                </table>
          </tr>
        </table>
      <%
		end if
		next
		end if
	   %>
        <% case 2 %>
        <!--新闻模版-->
        <%
	    if rseof=1 then
			for i=1 to abcount
			E_SmallClassID=ArrayE_SmallClassID(i)
			E_SmallClassName=ArrayE_SmallClassName(i)
			E_smallClassview=ArrayE_smallClassview(i)
		if ArrayE_smallClassview(i)=1 then 
		%>
        <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber5">
          <tr>
            <td height="25" background="Images/dh_bg.gif">&nbsp;&nbsp;&nbsp;<img src="Images/arrow_dh_1.gif" width="7" height="7">&nbsp;&nbsp;&nbsp;<font class=body><%=E_SmallClassName%></font></td>
          </tr>
          <%set rs3=server.CreateObject("ADODB.RecordSet")
			if uselevel=1 then
				if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then
				rs3.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_SmallClassID=" & E_SmallClassID &" and checkked=1) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("key")="" then
					rs3.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_SmallClassID=" & E_SmallClassID &" and checkked=1 ) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("key")="selfreg" then
					if Request.cookies(eChsys)("reglevel")=3 then
						rs3.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_SmallClassID=" & E_SmallClassID &" and checkked=1 ) order by newsid DESC"
					end if
				if Request.cookies(eChsys)("reglevel")=2 then
					rs3.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_SmallClassID=" & E_SmallClassID &" and checkked=1 ) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("reglevel")=1 then
					rs3.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_SmallClassID=" & E_SmallClassID &" and checkked=1 ) order by newsid DESC"
				end if
			end if
			else
			rs3.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_SmallClassID=" & E_SmallClassID &" and checkked=1) order by newsid DESC"
			end if
				rs3.Open rs3.Source,conn,1,1
			if not rs3.EOF then
			while not rs3.EOF
				
                             if  1 then 
                             //if showyear=1 then
					newsurl="E_ReadNews.asp?NewsID=" & rs3("NewsID")
					newswwwurl=rs3("titleface")
					datetime="<font class=middle>(" & year(rs3("UpdateTime"))  &"年"& Month(rs3("UpdateTime"))  &"月"& Day(rs3("UpdateTime")) &"日)</font>"
				else
					newsurl="E_ReadNews.asp?NewsID=" & rs3("NewsID")
					newswwwurl=rs3("titleface")
					datetime="<font class=middle>("& Month(rs3("UpdateTime"))  &"月"& Day(rs3("UpdateTime")) &"日)</font>"
				end if
				if rs3("picnews")=1 then
					img="<font class=pic_word>[图]</font>"
				else
					img=""
				end if
					title=trim(rs3("title"))
					title=replace(title,"<br>","")
		%>
          <tr>
            <td><div align="center">
                <table width="100%" border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td width="65%">&nbsp;&nbsp;<img src="IMAGES/006.gif" width="10" height="10"> <%=img%> <a class=middle href="<%if rs3("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" title="<%=rs3("title")%>" target="_blank"> <<%=rs3("titletype")%>><font color="<%=rs3("titlecolor")%>"> <%=CutStr(title,46)%></font></<%=rs3("titletype")%>></a>
                        <!--标题后评论提示-->
                        <% if rs3("titlesize")>=1 then %>
                        <A class=middle HREF="<%=path%>review.asp?NewsID=<%=rs3("NewsID")%>" target="_blank" ><font color=red><b>评</b></font></A>
                        <%end if %>
                        <!--标题后评论提示-->
                    </td>
                    <td width="19%" align="left">
                      <div align="right">
                      <%if showtime="1" then%>
                        <%=datetime%>
                        <%end if%>
                    &nbsp; </div></td>
                    <td width="16%" align="left"><%if year(rs3("updatetime"))=year(date()) and month(rs3("updatetime"))=month(date()) and day(rs3("updatetime"))=day(date()) then%>
                        <img src="images/new.gif">
                        <%end if%>
                        <%if rs3("goodnews")="1" then%>
                        <img src="images/g.gif" >
                        <%end if%>
                    </td>
                  </tr>
              </table>
            </div></td>
          </tr>
          <%
		rs3.MoveNext
		wend
		Rs3.close
		set rs3=nothing
		%>
          <%
			else
			Response.Write "<tr><td> <div align='center'><table width='98%' border='0' cellpadding='3' cellspacing='0'><tr><td width='62%'>&nbsp;&nbsp;<img src='IMAGES/006.gif' width='10' height='10'> 暂无信息</td></tr></table></div></td></tr>"
		    end if
			%>
          <tr>
            <td width=100% height="19" align=right><a class=class href='E_SmallClass.asp?E_typeid=<%=E_typeid%>&E_BigClassID=<%=E_BigClassID%>&E_SmallClassID=<%=E_SmallClassID%>'><img src="images/more.gif" border="0" alt="更多<%=E_SmallClassName%>"></a>&nbsp;&nbsp;</td>
          </tr>
        </table>
      <%
		end if
			next
		end if
		%>
        <% case 3 %>
        <!--网址模版-->
        <%
	    if rseof=1 then
		for i=1 to abcount
		E_SmallClassID=ArrayE_SmallClassID(i)
		E_SmallClassName=ArrayE_SmallClassName(i)
		E_smallClassview=ArrayE_smallClassview(i)
		if ArrayE_smallClassview(i)=1 then 
		%>
        <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber5">
          <tr>
            <td height="25" background="Images/dh_bg.gif">&nbsp;&nbsp;&nbsp;<img src="Images/arrow_dh_1.gif" width="7" height="7">&nbsp;&nbsp;&nbsp;<font class=body><%=E_SmallClassName%></font></td>
          </tr>
          <%set rs3=server.CreateObject("ADODB.RecordSet")
		if uselevel=1 then
			if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then
				rs3.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_SmallClassID=" & E_SmallClassID &" and checkked=1) order by newsid DESC"
			end if
			if Request.cookies(eChsys)("key")="" then
				rs3.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_SmallClassID=" & E_SmallClassID &" and checkked=1 ) order by newsid DESC"
			end if
			if Request.cookies(eChsys)("key")="selfreg" then
				if Request.cookies(eChsys)("reglevel")=3 then
					rs3.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_SmallClassID=" & E_SmallClassID &" and checkked=1 ) order by newsid DESC"
				end if
			if Request.cookies(eChsys)("reglevel")=2 then
				rs3.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_SmallClassID=" & E_SmallClassID &" and checkked=1 ) order by newsid DESC"
				end if
			if Request.cookies(eChsys)("reglevel")=1 then
				rs3.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_SmallClassID=" & E_SmallClassID &" and checkked=1 ) order by newsid DESC"
				end if
			end if
			else
			rs3.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_SmallClassID=" & E_SmallClassID &" and checkked=1) order by newsid DESC"
			end if
			rs3.Open rs3.Source,conn,1,1
		%>
          <tr>
            <td><table width="98%" border="1" cellspacing="0" cellpadding="3" bordercolorlight=#cccccc  bordercolordark=#ffffff align="center">
                <!--网址换行显示1-->
                <%
					if not rs3.EOF then
					do while not rs3.EOF%>
                <tr>
                  <%
						Content=htmlencode4(rs3("Content"))
					%>
                  <td width=25% align=center valign="middle"><%if not rs3.EOF then%>
                      <div align="center"> <a class=middle href="<%=rs3("Original")%>" target=_blank title="<%=CutStr(nohtml(rs3("Content")),80)%>"><%=CutStr(rs3("title"),30)%></a> </div></td>
                  <%rs3.movenext   
					end if %>
                  <!--网址换行显示2-->
                  <td width=25% align=center valign="middle"><%if not rs3.EOF then%>
                      <div align="center"> <a class=middle href="<%=rs3("Original")%>" target=_blank title="<%=CutStr(nohtml(rs3("Content")),80)%>"><%=CutStr(rs3("title"),30)%></a> </div></td>
                  <%rs3.movenext   
					end if %>
                </tr>
                <!--网址换行结束-->
                <%loop
				rs3.Close
				set rs3=nothing
				%>
                <%
			else
			Response.Write "<tr><td> <div align='center'><table width='98%' border='0' cellpadding='3' cellspacing='0'><tr><td width='62%'>&nbsp;&nbsp;<img src='IMAGES/006.gif' width='10' height='10'> 暂无信息</td></tr></table></div></td></tr>"
		    end if
			%>
              </table>
          </tr>
          <tr>
            <td><table cellSpacing=0 cellPadding=3 width="100%" align=center border=0>
                <tr>
                  <td width=100% height="19" align=right><a class=class href='E_SmallClass.asp?E_typeid=<%=E_typeid%>&E_BigClassID=<%=E_BigClassID%>&E_SmallClassID=<%=E_SmallClassID%>'><img src="images/more.gif" border="0" alt="更多<%=E_SmallClassName%>"></a>&nbsp;&nbsp;</td>
                </tr>
              </table>
          </tr>
          <%
		end if
		next
		end if
		%>
        </table>
      <% case 4 %>
        <!--软件模版-->
        <%
	    if rseof=1 then
			for i=1 to abcount
			E_SmallClassID=ArrayE_SmallClassID(i)
			E_SmallClassName=ArrayE_SmallClassName(i)
			E_smallClassview=ArrayE_smallClassview(i)
		if ArrayE_smallClassview(i)=1 then 
		%>
        <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber5">
          <tr>
            <td height="25" background="Images/dh_bg.gif">&nbsp;&nbsp;&nbsp;<img src="Images/arrow_dh_1.gif" width="7" height="7">&nbsp;&nbsp;&nbsp;<font class=body><%=E_SmallClassName%></font></td>
          </tr>
          <%set rs3=server.CreateObject("ADODB.RecordSet")
			if uselevel=1 then
				if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then
					rs3.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_SmallClassID=" & E_SmallClassID &" and checkked=1) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("key")="" then
				rs3.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_SmallClassID=" & E_SmallClassID &" and checkked=1 ) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("key")="selfreg" then
					if Request.cookies(eChsys)("reglevel")=3 then
						rs3.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_SmallClassID=" & E_SmallClassID &" and checkked=1 ) order by newsid DESC"
					end if
				if Request.cookies(eChsys)("reglevel")=2 then
					rs3.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_SmallClassID=" & E_SmallClassID &" and checkked=1 ) order by newsid DESC"
					end if
				if Request.cookies(eChsys)("reglevel")=1 then
					rs3.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_SmallClassID=" & E_SmallClassID &" and checkked=1 ) order by newsid DESC"
					end if
				end if
				else
					rs3.Source="select top " & bigclassshownum & " * from "& db_EC_News_Table &" where (E_SmallClassID=" & E_SmallClassID &" and checkked=1) order by newsid DESC"
				end if
			rs3.Open rs3.Source,conn,1,1
			    if not rs3.EOF then
				while not rs3.EOF
					      if 1 then 
                                             //if showyear=1 then
						newsurl="E_ReadNews.asp?NewsID=" & rs3("NewsID")
						newswwwurl=rs3("titleface")
						datetime="<font class=middle>(" & year(rs3("UpdateTime"))  &"年"& Month(rs3("UpdateTime"))  &"月"& Day(rs3("UpdateTime")) &"日)</font>"
					else
						newsurl="E_ReadNews.asp?NewsID=" & rs3("NewsID")
						newswwwurl=rs3("titleface")
						datetime="<font class=middle>(" & Month(rs3("UpdateTime"))  &"月"& Day(rs3("UpdateTime")) &"日)</font>"
					end if
					if rs3("picnews")=1 then
						img="<img src='images/news_img.gif' border='0'>"
					else
						img=""
					end if
						title=trim(rs3("title"))
						title=replace(title,"<br>","")
			%>
          <tr>
            <%
		fileExt=lcase(getFileExtName(rs3("picname")))
		Content=htmlencode4(rs3("Content"))
		content=replace(content,"[---分页---]","")%>
            <td><div align="center">
                <table width="100%" border="0" cellpadding="0" cellspacing="0">
                  <tr bgcolor="#EFEFEF">
                    <td colspan="2">&nbsp;<img src="images/news_img.gif" width="9" height="9"> <a class=middle href="<%if rs3("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" title="<%=htmlencode4(title)%>" target="_blank"> <font color="<%=rs3("titlecolor")%>"><strong> <%=CutStr(title,40)%> </strong></font> </a> </td>
                    <td width="22%" align="right" bgcolor="#EFEFEF"><div align="right">
                      <%if showtime="1" then%>
                        <%=datetime%>
                        <%end if%>
                    &nbsp;</div></td>
                    <td width="14%" align="center"><%if year(rs3("updatetime"))=year(date()) and month(rs3("updatetime"))=month(date()) and day(rs3("updatetime"))=day(date()) then%>
                        <img src="images/new.gif">
                        <%end if%>
                        <%if rs3("goodnews")="1" then%>
                        <img src="images/g.gif" >
                        <%end if%>
                    </td>
                  </tr>
                  <tr>
                    <td width="39%"><a class=middle href="<%if rs3("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" target=_blank title="<%=htmlencode4(rs3("title"))%>">
                      <%if   rs3("picname")=("")   then%>
                      <img  src="IMAGES/softno.gif" width="65" height="65" border=0 align="left">
                      <%else%>
                      <%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
                      <img  src="<%=FileUploadPath & rs3("picname")%>" width="65" height="65" border=0 align="left">
                      <%end if%>
                      <%if fileext="swf" then%>
                      <object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="65" height="65" border=0 >
                        <param name=movie value="<%=FileUploadPath & rs3("picname")%>">
                        <param name=quality value=high>
                        <param name='Play' value='-1'>
                        <param name='Loop' value='0'>
                        <param name='Menu' value='-1'>
                        <embed src="<%=FileUploadPath & rs3("picname")%>" width="6" 5height="65" border=0  pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
                      </object>
                      <%end if%>
                      <%end if%>
                    </a> </td>
                    <td colspan="3" valign="top"><a class=middle href="<%if rs3("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" target="_blank" title="<%=htmlencode4(title)%>"><%=CutStr(nohtml(rs3("Content")),250)%></a> </td>
                  </tr>
              </table>
            </div></td>
          </tr>
          <%
		rs3.MoveNext
		wend
		Rs3.close
		set rs3=nothing
		%>
          <%
			else
			Response.Write "<tr><td> <div align='center'><table width='98%' border='0' cellpadding='3' cellspacing='0'><tr><td width='62%'>&nbsp;&nbsp;<img src='IMAGES/006.gif' width='10' height='10'> 暂无信息</td></tr></table></div></td></tr>"
		    end if
			%>
          <tr>
            <td width=100% height="19" align=right><a class=class href='E_SmallClass.asp?E_typeid=<%=E_typeid%>&E_BigClassID=<%=E_BigClassID%>&E_SmallClassID=<%=E_SmallClassID%>'><img src="images/more.gif" border="0" alt="更多<%=E_SmallClassName%>"></a>&nbsp;&nbsp; </td>
          </tr>
        </table>
      <%
	end if
		next
	end if
	%>
        <%end select
	conn.close
	set conn=nothing
	%></td>
  </tr>
</table>
<%end if%>
<!--#include file="other_bottom.asp"-->
</body>
</html>
