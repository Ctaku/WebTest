<%@ Language=VBScript%>
<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="function.asp"-->
<!--#include file="char.inc"-->
<%
E_typeid=checkstr(request("E_typeid"))
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

		PageShowSize = 10            '每页显示多少个页
		MyPageSize   = 40           '每页显示多少条新闻
			
		if Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
			MyPage=1
		else
			MyPage=Int(Abs(Request("page")))
		end if

		set rs=server.CreateObject("ADODB.RecordSet")
        rs.Source="select * from "& db_EC_BigClass_Table &" where E_typeid=" & E_typeid &" order by E_bigclassorder"
        rs.Open rs.Source,conn,1,1

		dim E_BigClassID
		E_BigClassID=checkstr(request("E_BigClassID"))
		set rs=server.CreateObject("ADODB.RecordSet")
		rs.Source="select * from "& db_EC_SmallClass_Table &" where E_BigClassID="& E_BigClassID &" order by smallClassorder"
		rs.Open rs.Source,conn,1,1
		i=1
		Dim ArrayE_typeid(10000),ArrayE_BigClassName(10000),ArrayE_smallclassview(10000)

        if not rs.EOF then
		rseof=1
		end if
	 end if

		rs.close

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
<html>
<head>
<meta http-equiv="Content-Type" content="">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=E_BigClassName%>__<%=E_typename%>__<%=eChSys%></title>
<LINK href=news.css rel=stylesheet>
</head>
<body topmargin="0">
<%
dim E_BigClassName
E_BigClassID=checkstr(request("E_BigClassID"))
set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from "& db_EC_BigClass_Table &" where E_typeid="& E_BigClassID &" order by E_bigclassorder"
rs.Open rs.Source,conn,1,1
'E_BigClassName=rs("E_BigClassName")
rs.close
set rs=nothing
%>
<!--#include file="other_Top.asp"-->
<table width="780" height="400" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="210" height="400" valign="top" bgcolor="F1F7FC"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#999999" width="100%" id="AutoNumber4" height="25">
      <tr>
        <td height="50" align=center background="Images/type_dh.gif"><font color="#000000" class="yellow_title"><img src="Images/arrow_dh_1.gif" width="7" height="7">&nbsp;</font><span class="yellow_title"><font color="#000000">&nbsp;</font>最新图文</span> </td>
      </tr>
      <tr>
        <td width="100%" height="20" align=center><br>
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
              <table width="90%" border="0" cellspacing="0" cellpadding="3" align="center" style="TABLE-LAYOUT: fixed">
                <tr>
                  <td style="WORD-WRAP: break-word" align="center"><a class=middle href="E_ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=htmlencode4(rs3("title"))%>">
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
 </a> </td>
                </tr>
				<tr>
				<td style="WORD-WRAP: break-word" align="center"><a class=middle href="E_ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=htmlencode4(rs3("title"))%>"><%=CutStr(htmlencode4(rs3("title")),24)%></a></td>
				</tr>
              </table>
          <%
        rs3.MoveNext
	    wend
	    else
		Response.Write "<tr><td width=100% align=center height=18>暂无</td></tr>"
	    end if
	    rs3.close
	    set rs3=nothing
%>
        </td>
      </tr>
    </table></td>
    <td  width="570"valign="top" bgcolor="#FFFFFF">
	  <table width="570" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="570" height="25" background="Images/dh_bg.gif">&nbsp;&nbsp;&nbsp;<img src="Images/arrow_dh_1.gif" width="7" height="7">&nbsp;&nbsp;&nbsp;<font class=m_tittle>&nbsp;</font>栏目导航&nbsp;&nbsp;<a href="./" class="daohang">网站首页</a>&gt;&gt;<a href=E_Type.asp?E_typeid=<%=E_typeid%> class="daohang"><%=E_typename%></a>
          <% Response.Write "&gt;&gt;" & E_BigClassName & "" %></td>
      </tr>
	   <tr>
        <td height="5"><div align="center">
          <script language=javascript src=./zongg/ad.asp?i=14></script>
        </div></td>
      </tr>
      <tr>
        <td height="25" background="Images/dh_bg.gif">&nbsp;&nbsp;&nbsp;&nbsp;<img src="Images/arrow_dh_1.gif" width="7" height="7">&nbsp;&nbsp;<font class=body>
          <% Response.Write "" & E_BigClassName & "" %>
          >无小类文章区</font></td>
      </tr>
      <!--无小类文章区-->
      <!--模版开始-->
      <% select case bigclasszs %>
      <% case "1" %>
      <!--图片模版-->
      <%set rsnosmall=server.CreateObject("ADODB.RecordSet")
			if uselevel=1 then
			    
				if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then
				rsnosmall.Source="select * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("key")="" then
					rsnosmall.Source="select * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("key")="selfreg" then
					if Request.cookies(eChsys)("reglevel")=3 then
						rsnosmall.Source="select * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
					end if
				if Request.cookies(eChsys)("reglevel")=2 then
					rsnosmall.Source="select * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("reglevel")=1 then
					rsnosmall.Source="select * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				
			end if
			else
			rsnosmall.Source="select * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1) order by newsid DESC"
			end if
			rsnosmall.Open rsnosmall.Source,conn,1,1

if not rsnosmall.EOF then
rsnosmall.PageSize     = MyPageSize
MaxPages         = rsnosmall.PageCount
rsnosmall.absolutepage = MyPage
total            = rsnosmall.RecordCount
i = 0

%>
      <tr>
          <td>
		<table width="90%" border="0" cellspacing="0" cellpadding="3" align="center">
            <!--图片换行显示1-->
            <% do until rsnosmall.Eof or i = rsnosmall.PageSize %>
            <tr>
              <%
fileExt=lcase(getFileExtName(rsnosmall("picname")))
Content=htmlencode4(rsnosmall("Content"))
content=replace(content,"[---分页---]","")
%>
              <% if not rsnosmall.EOF then %>
              <td width=25% align=center valign="top" style="table-layout:fixed; word-break:break-all"><div align="center">
                  <% content=htmlencode4(rsnosmall("Content"))
                       content=replace(content,"[---分页---]","")%>
                  <a class=middle href="E_ReadNews.asp?NewsID=<%=rsnosmall("NewsID")%>" target=_blank title="<%=CutStr(nohtml(Content),150)%>">
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
                  </a><br>
                  <br>
              <a class=middle href="E_ReadNews.asp?NewsID=<%=rsnosmall("NewsID")%>" target=_blank title="<%=htmlencode4(rsnosmall("title"))%>"> <%=CutStr(rsnosmall("title"),18)%> </a> </div></td>
              <%rsnosmall.movenext
i = i + 1
end if
%>
              <!--图片换行显示2-->
              <td width=25% align=center valign="top" style="table-layout:fixed; word-break:break-all"><% if not rsnosmall.EOF then %>
                  <div align="center">
                    <%
fileExt=lcase(getFileExtName(rsnosmall("picname")))
Content=htmlencode4(rsnosmall("Content"))
content=replace(content,"[---分页---]","")
%>
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
                <a class=middle href="E_ReadNews.asp?NewsID=<%=rsnosmall("NewsID")%>" target=_blank title="<%=htmlencode4(rsnosmall("title"))%>"> <%=CutStr(rsnosmall("title"),18)%> </a> </div></td>
              <%rsnosmall.movenext
i = i + 1
end if
%>
              <!--图片换行显示3-->
              <td width=25% align=center valign="top" style="table-layout:fixed; word-break:break-all"><% if not rsnosmall.EOF then %>
                  <div align="center">
                    <%
fileExt=lcase(getFileExtName(rsnosmall("picname")))
Content=htmlencode4(rsnosmall("Content"))
content=replace(content,"[---分页---]","")
%>
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
                <a class=middle href="E_ReadNews.asp?NewsID=<%=rsnosmall("NewsID")%>" target=_blank title="<%=htmlencode4(rsnosmall("title"))%>"> <%=CutStr(rsnosmall("title"),18)%> </a> </div></td>
              <%rsnosmall.movenext
i = i + 1
end if
%>
              <!--图片换行显示4-->
              <td width=25% align=center valign="top" style="table-layout:fixed; word-break:break-all"><% if not rsnosmall.EOF then %>
                  <div align="center">
                    <%
fileExt=lcase(getFileExtName(rsnosmall("picname")))
Content=htmlencode4(rsnosmall("Content"))
content=replace(content,"[---分页---]","")
%>
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
                <a class=middle href="E_ReadNews.asp?NewsID=<%=rsnosmall("NewsID")%>" target=_blank title="<%=htmlencode4(rsnosmall("title"))%>"> <%=CutStr(rsnosmall("title"),18)%> </a> </div></td>
              <%rsnosmall.movenext
i = i + 1
end if
%>
            </tr>
            <!--图片换行结束-->
            <%loop%>
          </table>		  </td>
        </tr>
		
      <tr>
        <td>
		<TABLE cellSpacing=0 cellPadding=3 width="100%" align=center border=0 >
            <tr>
              <td width="100%" align=center valign="middle">共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 页，每页 <%=MyPageSize%> 条
                <%
url="E_nosmallclass.asp?E_typeid=" & E_typeid & "&E_BigClassID=" & request_E_BigClassID
request_E_SmallClassID
PageNextSize=int((MyPage-1)/PageShowSize)+1
Pagetpage=int((total-1)/rsnosmall.PageSize)+1

if PageNextSize >1 then
PagePrev=PageShowSize*(PageNextSize-1)
Response.write "<a class=black href='" & Url & "&page=" & PagePrev & "' title='上" & PageShowSize & "页'>上一翻页</a> "
Response.write "<a class=black href='" & Url & "&page=1' title='第1页'>页首</a> "
end if
if MyPage-1 > 0 then
Prev_Page = MyPage - 1
Response.write "<a class=black href='" & Url & "&page=" & Prev_Page & "' title='第" & Prev_Page & "页'>上一页</a> "
end if

if Maxpages>=PageNextSize*PageShowSize then
PageSizeShow = PageShowSize
Else
PageSizeShow = Maxpages-PageShowSize*(PageNextSize-1)
End if
If PageSizeShow < 1 Then PageSizeShow = 1
for PageCounterSize=1 to PageSizeShow
PageLink = (PageCounterSize+PageNextSize*PageShowSize)-PageShowSize
if PageLink <> MyPage Then
Response.write "<a class=black href='" & Url & "&page=" & PageLink & "'>[" & PageLink & "]</a> "
else
Response.Write "<B>["& PageLink &"]</B> "
end if
If PageLink = MaxPages Then Exit for
Next

if Mypage+1 <=Pagetpage  then
Next_Page = MyPage + 1
Response.write "<a class=black href='" & Url & "&page=" & Next_Page & "' title='第" & Next_Page & "页'>下一页</A>"
end if

if MaxPages > PageShowSize*PageNextSize then
PageNext = PageShowSize * PageNextSize + 1
Response.write " <A class=black href='" & Url & "&page=" & Pagetpage & "' title='第"& Pagetpage &"页'>页尾</A>"
Response.write " <a class=black href='" & Url & "&page=" & PageNext & "' title='下" & PageShowSize & "页'>下一翻页</a>"
End if
else
Response.write "<tr><td align=center>&nbsp;本大类无小类文章区暂无信息</td></tr>"
				
End If
rsnosmall.close
set rsnosmall=nothing
%>              </td>
            </tr>
          </table>
            <!--图片模版-->			
            <!--无小类文章区-->			
			     <% case "2" %>
				<!--新闻模版-->
                <%set rsnosmall=server.CreateObject("ADODB.RecordSet")
			if uselevel=1 then
			    
				if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then
				rsnosmall.Source="select * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("key")="" then
					rsnosmall.Source="select * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("key")="selfreg" then
					if Request.cookies(eChsys)("reglevel")=3 then
						rsnosmall.Source="select * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
					end if
				if Request.cookies(eChsys)("reglevel")=2 then
					rsnosmall.Source="select * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("reglevel")=1 then
					rsnosmall.Source="select * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				
			end if
			else
			rsnosmall.Source="select * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1) order by newsid DESC"
			end if
			rsnosmall.Open rsnosmall.Source,conn,1,1

			if not rsnosmall.EOF then

			rsnosmall.PageSize     = MyPageSize
			MaxPages         = rsnosmall.PageCount
			rsnosmall.absolutepage = MyPage
			total            = rsnosmall.RecordCount

			i = 0
			do until rsnosmall.Eof or i = rsnosmall.PageSize

				if showyear=1 then
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

                <table width="100%" border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td width="68%">&nbsp;&nbsp;<img src="IMAGES/006.gif" width="10" height="10"> <%=img%> <a class=middle href="<%if rsnosmall("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" title="<%=rsnosmall("title")%>" target="_blank"> <<%=rsnosmall("titletype")%>><font color="<%=rsnosmall("titlecolor")%>"> <%=CutStr(title,46)%></font></<%=rsnosmall("titletype")%>></a>
                        <!--标题后评论提示-->
                        <% if rsnosmall("titlesize")>=1 then %>
                        <A class=middle HREF="<%=path%>review.asp?NewsID=<%=rsnosmall("NewsID")%>" target="_blank" ><font color=red><b>评</b></font></A>
                        <%end if %>
                    <!--标题后评论提示-->                    </td>
                    <td width="16%" align="left">
					    <div align="right">
					      <%if showtime="1" then%>
                          <%=datetime%>
                          <%end if%>
                    &nbsp; </div></td>
                    <td width="16%" align="left"><%if year(rsnosmall("updatetime"))=year(date()) and month(rsnosmall("updatetime"))=month(date()) and day(rsnosmall("updatetime"))=day(date()) then%>
                        <img src="images/new.gif">
                        <%end if%>
                        <%if rsnosmall("goodnews")="1" then%>
                        <img src="images/g.gif" >
                    <%end if%>   </td>
                  </tr>
            </table>
          <%
		rsnosmall.MoveNext
		i = i + 1
		loop
		%>
          <!--无小类文章区-->
    </table>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" align=center border=0>
        <tr>
          <td width="100%" align=center valign="middle">共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 页，每页 <%=MyPageSize%> 条
            <%
url="E_nosmallclass.asp?E_typeid=" & E_typeid & "&E_BigClassID=" & request_E_BigClassID 
PageNextSize=int((MyPage-1)/PageShowSize)+1
Pagetpage=int((total-1)/rsnosmall.PageSize)+1

if PageNextSize >1 then
PagePrev=PageShowSize*(PageNextSize-1)
Response.write "<a class=black href='" & Url & "&page=" & PagePrev & "' title='上" & PageShowSize & "页'>上一翻页</a> "
Response.write "<a class=black href='" & Url & "&page=1' title='第1页'>页首</a> "
end if
if MyPage-1 > 0 then
Prev_Page = MyPage - 1
Response.write "<a class=black href='" & Url & "&page=" & Prev_Page & "' title='第" & Prev_Page & "页'>上一页</a> "
end if

if Maxpages>=PageNextSize*PageShowSize then
PageSizeShow = PageShowSize
Else
PageSizeShow = Maxpages-PageShowSize*(PageNextSize-1)
End if
If PageSizeShow < 1 Then PageSizeShow = 1
for PageCounterSize=1 to PageSizeShow
PageLink = (PageCounterSize+PageNextSize*PageShowSize)-PageShowSize
if PageLink <> MyPage Then
Response.write "<a class=black href='" & Url & "&page=" & PageLink & "'>[" & PageLink & "]</a> "
else
Response.Write "<B>["& PageLink &"]</B> "
end if
If PageLink = MaxPages Then Exit for
Next

if Mypage+1 <=Pagetpage  then
Next_Page = MyPage + 1
Response.write "<a class=black href='" & Url & "&page=" & Next_Page & "' title='第" & Next_Page & "页'>下一页</A>"
end if

if MaxPages > PageShowSize*PageNextSize then
PageNext = PageShowSize * PageNextSize + 1
Response.write " <A class=black href='" & Url & "&page=" & Pagetpage & "' title='第"& Pagetpage &"页'>页尾</A>"
Response.write " <a class=black href='" & Url & "&page=" & PageNext & "' title='下" & PageShowSize & "页'>下一翻页</a>"
End if
else
Response.write "<tr><td align=center>&nbsp;本大类无小类文章区暂无信息</td></tr>"
				
End If
rsnosmall.close
set rsnosmall=nothing
%>
          </td>
        </tr>
        <!--新闻模版-->
        <% case "3" %>
        <!--网址模版-->
        <%set rsnosmall=server.CreateObject("ADODB.RecordSet")
			if uselevel=1 then
			    
				if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then
				rsnosmall.Source="select * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("key")="" then
					rsnosmall.Source="select * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("key")="selfreg" then
					if Request.cookies(eChsys)("reglevel")=3 then
						rsnosmall.Source="select * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
					end if
				if Request.cookies(eChsys)("reglevel")=2 then
					rsnosmall.Source="select * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("reglevel")=1 then
					rsnosmall.Source="select * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				
			end if
			else
			rsnosmall.Source="select * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1) order by newsid DESC"
			end if
			rsnosmall.Open rsnosmall.Source,conn,1,1

if not rsnosmall.EOF then
rsnosmall.PageSize     = MyPageSize
MaxPages         = rsnosmall.PageCount
rsnosmall.absolutepage = MyPage
total            = rsnosmall.RecordCount
i = 0
%>
  <td><table width="98%" border="1" cellspacing="0" cellpadding="3" bordercolorlight=#cccccc  bordercolordark=#ffffff align="center">
    <!--网址换行显示1-->
    <% do until rsnosmall.Eof or i = rsnosmall.PageSize %>
    <tr>
      <%
Content=htmlencode4(rsnosmall("Content"))
%>
      <% if not rsnosmall.EOF then %>
      <td width=50% align=center valign="middle"><div align="center"> <a class=middle href="<%=rsnosmall("Original")%>" target=_blank title="<%=CutStr(nohtml(rsnosmall("Content")),80)%>"> <%=CutStr(rsnosmall("title"),30)%> </a></div></td>
      <%rsnosmall.movenext
i = i + 1
end if
%>
      <!--网址换行显示2-->
      <% if not rsnosmall.EOF then %>
      <td width=50% align=center valign="middle"><div align="center"> <a class=middle href="<%=rsnosmall("Original")%>" target=_blank title="<%=CutStr(nohtml(rsnosmall("Content")),80)%>"> <%=CutStr(rsnosmall("title"),30)%> </a> </div></td>
      <%rsnosmall.movenext
i = i + 1
end if
%>
    </tr>
    <!--网址换行结束-->
    <%loop%>
  </table></td>
  <tr>
    <td height="271"><TABLE cellSpacing=0 cellPadding=3 width="100%" align=center border=0 >
      <tr>
        <td width="100%" align=center valign="middle">共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 页，每页 <%=MyPageSize%> 条
          <%
url="E_nosmallclass.asp?E_typeid=" & E_typeid & "&E_BigClassID=" & request_E_BigClassID
request_E_SmallClassID
PageNextSize=int((MyPage-1)/PageShowSize)+1
Pagetpage=int((total-1)/rsnosmall.PageSize)+1

if PageNextSize >1 then
PagePrev=PageShowSize*(PageNextSize-1)
Response.write "<a class=black href='" & Url & "&page=" & PagePrev & "' title='上" & PageShowSize & "页'>上一翻页</a> "
Response.write "<a class=black href='" & Url & "&page=1' title='第1页'>页首</a> "
end if
if MyPage-1 > 0 then
Prev_Page = MyPage - 1
Response.write "<a class=black href='" & Url & "&page=" & Prev_Page & "' title='第" & Prev_Page & "页'>上一页</a> "
end if

if Maxpages>=PageNextSize*PageShowSize then
PageSizeShow = PageShowSize
Else
PageSizeShow = Maxpages-PageShowSize*(PageNextSize-1)
End if
If PageSizeShow < 1 Then PageSizeShow = 1
for PageCounterSize=1 to PageSizeShow
PageLink = (PageCounterSize+PageNextSize*PageShowSize)-PageShowSize
if PageLink <> MyPage Then
Response.write "<a class=black href='" & Url & "&page=" & PageLink & "'>[" & PageLink & "]</a> "
else
Response.Write "<B>["& PageLink &"]</B> "
end if
If PageLink = MaxPages Then Exit for
Next

if Mypage+1 <=Pagetpage  then
Next_Page = MyPage + 1
Response.write "<a class=black href='" & Url & "&page=" & Next_Page & "' title='第" & Next_Page & "页'>下一页</A>"
end if

if MaxPages > PageShowSize*PageNextSize then
PageNext = PageShowSize * PageNextSize + 1
Response.write " <A class=black href='" & Url & "&page=" & Pagetpage & "' title='第"& Pagetpage &"页'>页尾</A>"
Response.write " <a class=black href='" & Url & "&page=" & PageNext & "' title='下" & PageShowSize & "页'>下一翻页</a>"
End if
else
Response.write "<tr><td align=center>&nbsp;本大类无小类文章区暂无信息</td></tr>"
				
End If

rsnosmall.close
set rsnosmall=nothing
%>
        </td>
      </tr>
    </table>
        <!--网址模版-->
        <% case "4" %>
        <!--软件模版-->
        <%set rsnosmall=server.CreateObject("ADODB.RecordSet")
			if uselevel=1 then
			    
				if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then
				rsnosmall.Source="select * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("key")="" then
					rsnosmall.Source="select * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("key")="selfreg" then
					if Request.cookies(eChsys)("reglevel")=3 then
						rsnosmall.Source="select * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
					end if
				if Request.cookies(eChsys)("reglevel")=2 then
					rsnosmall.Source="select * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("reglevel")=1 then
					rsnosmall.Source="select * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				
			end if
			else
			rsnosmall.Source="select * from "& db_EC_News_Table &" where (E_BigClassID=" & E_BigClassID &"  and E_SmallClassID is null and checkked=1) order by newsid DESC"
			end if
			rsnosmall.Open rsnosmall.Source,conn,1,1

if not rsnosmall.EOF then
rsnosmall.PageSize     = MyPageSize
MaxPages         = rsnosmall.PageCount
rsnosmall.absolutepage = MyPage
total            = rsnosmall.RecordCount

i = 0
do until rsnosmall.Eof or i = rsnosmall.PageSize

newsurl="E_ReadNews.asp?NewsID=" & rsnosmall("NewsID")
newswwwurl=rsnosmall("titleface")
datetime="<font class=middle>(" & Month(rsnosmall("UpdateTime"))  &"月"& Day(rsnosmall("UpdateTime")) &"日)</font>"


title=htmlencode4(trim(rsnosmall("title")))
%>
        <div align="center">
          <table width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr bgcolor="#EFEFEF">
              <td colspan="2">&nbsp;<img src="images/news_img.gif" width="9" height="9" border="0">&nbsp; <a class=middle href="<%if rsnosmall("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" target="_blank" title="<%=htmlencode4(title)%>"><strong><font color="<%=rsnosmall("titlecolor")%>"> <%=CutStr(title,40)%> </font></strong></a></td>
              <td width="17%" align="right" bgcolor="#EFEFEF"><%if showtime="1" then%>
                  <%=datetime%>
                  <%end if%>
              </td>
              <td width="15%" align="center" bgcolor="#EFEFEF"><div align="right">
                  <%if year(rsnosmall("updatetime"))=year(date()) and month(rsnosmall("updatetime"))=month(date()) and day(rsnosmall("updatetime"))=day(date()) then%>
                  <img src="images/new.gif" border="0">
                  <%end if%>
                  <%if rsnosmall("goodnews")="1" then%>
                  <img src="images/g.gif" border="0">
                  <%end if%>
              </div></td>
            </tr>
            <tr>
              <%
fileExt=lcase(getFileExtName(rsnosmall("picname")))
Content=htmlencode4(rsnosmall("Content"))
content=replace(content,"[---分页---]","")%>
              <td width="35%"><a class=middle href="<%if rsnosmall("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" target=_blank title="<%=htmlencode4(rsnosmall("title"))%>">
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
              </a></td>
              <td colspan="3" valign="top"><a class=middle href="<%if rsnosmall("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" target="_blank" title="<%=htmlencode4(title)%>"><%=CutStr(nohtml(rsnosmall("Content")),250)%></a> </td>
            </tr>
          </table>
          <%
rsnosmall.MoveNext
i = i + 1
loop
%>
          <TABLE cellSpacing=0 cellPadding=3 width="100%" align=center border=0 >
            <tr>
              <td width="100%" align=center valign="middle">共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 页，每页 <%=MyPageSize%> 条
                <%
url="E_nosmallclass.asp?E_typeid=" & E_typeid & "&E_BigClassID=" & request_E_BigClassID
request_E_SmallClassID
PageNextSize=int((MyPage-1)/PageShowSize)+1
Pagetpage=int((total-1)/rsnosmall.PageSize)+1

if PageNextSize >1 then
PagePrev=PageShowSize*(PageNextSize-1)
Response.write "<a class=black href='" & Url & "&page=" & PagePrev & "' title='上" & PageShowSize & "页'>上一翻页</a> "
Response.write "<a class=black href='" & Url & "&page=1' title='第1页'>页首</a> "
end if
if MyPage-1 > 0 then
Prev_Page = MyPage - 1
Response.write "<a class=black href='" & Url & "&page=" & Prev_Page & "' title='第" & Prev_Page & "页'>上一页</a> "
end if

if Maxpages>=PageNextSize*PageShowSize then
PageSizeShow = PageShowSize
Else
PageSizeShow = Maxpages-PageShowSize*(PageNextSize-1)
End if
If PageSizeShow < 1 Then PageSizeShow = 1
for PageCounterSize=1 to PageSizeShow
PageLink = (PageCounterSize+PageNextSize*PageShowSize)-PageShowSize
if PageLink <> MyPage Then
Response.write "<a class=black href='" & Url & "&page=" & PageLink & "'>[" & PageLink & "]</a> "
else
Response.Write "<B>["& PageLink &"]</B> "
end if
If PageLink = MaxPages Then Exit for
Next

if Mypage+1 <=Pagetpage  then
Next_Page = MyPage + 1
Response.write "<a class=black href='" & Url & "&page=" & Next_Page & "' title='第" & Next_Page & "页'>下一页</A>"
end if

if MaxPages > PageShowSize*PageNextSize then
PageNext = PageShowSize * PageNextSize + 1
Response.write " <A class=black href='" & Url & "&page=" & Pagetpage & "' title='第"& Pagetpage &"页'>页尾</A>"
Response.write " <a class=black href='" & Url & "&page=" & PageNext & "' title='下" & PageShowSize & "页'>下一翻页</a>"
End if
else
Response.write "<tr><td align=center>&nbsp;本大类无小类文章区暂无信息</td></tr>"
			
End If
rsnosmall.close
set rsnosmall=nothing
%>
              </td>
            </tr>
            <!--软件模版-->
            <!--模版结束-->
            <!--无小类文章区-->
          </table>
 <% end select%>
</div></td>
  </table>
 </td>
</tr>
</table>
<!--#include file="other_bottom.asp" -->
</body>
</html>
