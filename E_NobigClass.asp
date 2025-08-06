<%@ Language=VBScript%>
<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="function.asp"-->
<!--#include file="char.inc"-->
<%
E_typeid=checkstr(request("E_typeid"))
if E_typeid="" then
	Response.Write "<script>alert('未指定参数');history.back()</script>"
	response.end
else
	if not IsNumeric(E_typeid) then
		response.write "<script>alert('非法参数');history.back()</script>"
		response.end
	else
		dim tE_typename
		set rs5=server.CreateObject("ADODB.RecordSet")
		rs5.Source="select * from "& db_EC_Type_Table &" where E_typeid=" & E_typeid &" order by E_typeorder"
		rs5.Open rs5.Source,conn,1,1
		tE_typename=rs5("E_typename")
		E_typeorder=rs5("E_typeorder")
		rs5.Close
		set rs5=nothing

		PageShowSize = 10            '每页显示多少个页
		MyPageSize   = 40           '每页显示多少条新闻
			
		if Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
			MyPage=1
		else
			MyPage=Int(Abs(Request("page")))
		end if

		%>
		<%
		dim E_typeid
		E_typeid=checkstr(request("E_typeid"))
		set rs=server.CreateObject("ADODB.RecordSet")
		rs.Source="select * from "& db_EC_BigClass_Table &" where E_typeid="& E_typeid &" order by E_bigclassorder"
		rs.Open rs.Source,conn,1,1
		i=1

		if not rs.EOF then
			rseof=1
		end if
	end if
	
	rs.close
	%>

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=tE_typename%>__<%=eChSys%></title>
<LINK href=news.css rel=stylesheet>
<script language="JavaScript1.2">
<!--
function makevisible(cur,which){
if (which==0)
cur.filters.alpha.opacity=100
else
cur.filters.alpha.opacity=20
}

function MM_preloadImages() { //v3.0
var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}
//-->
</script>
</head>
<body>
<%
dim E_typename
E_typeid=checkstr(request("E_typeid"))
set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from "& db_EC_Type_Table &" where E_typeid="& E_typeid &" order by E_typeorder"
rs.Open rs.Source,conn,1,1
mode=rs("mode")
E_typename=rs("E_typename")
rs.close
set rs=nothing%>
<!--#include file="other_Top.asp"-->
<table width="780" height="400" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
       <tr>
         <td width="210" valign="top" bgcolor="F1F7FC">
	   <table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="50" align=center background="Images/type_dh.gif" class="yellow_title"><img src="Images/arrow_dh_1.gif" width="7" height="7">&nbsp;&nbsp;<span class="yellow_title">最新图文</span></td>
      </tr>
      <tr>
        <td width="100%" height="20" align=center><%
									set rs3=server.CreateObject("ADODB.RecordSet")
									if uselevel=1 then
										if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then
											rs3.Source ="select top " & top_img & " * from "& db_EC_News_Table &" where picnews=1 and checkked=1 and E_typeid="&E_typeid&" and picname is not null order by NewsID DESC"
										end if
										if Request.cookies(eChsys)("key")="" then
											rs3.Source ="select top " & top_img & " * from "& db_EC_News_Table &" where picnews=1 and checkked=1  and E_typeid="&E_typeid&" and picname is not null order by NewsID DESC"
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
                  <td style="WORD-WRAP: break-word"><a class=middle href="E_ReadNews.asp?NewsID=<%=rs3("NewsID")%>" target=_blank title="<%=htmlencode4(rs3("title"))%>">
                    <%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
                    <img  src="<%=FileUploadPath & rs3("picname")%>" width="65" height="65" border=0 align="left">
                    <%end if%>
                    <%if fileext="swf" then%>
                    <object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="65" height="65" border=0 align="left">
                      <param name=movie value="<%=FileUploadPath & rs3("picname")%>">
                      <param name=quality value=high>
                      <param name='Play' value='-1'>
                      <param name='Loop' value='0'>
                      <param name='Menu' value='-1'>
                      <embed src="<%=FileUploadPath & rs3("picname")%>" width="65" height="65" border=0 align="left" pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
                    </object>
                    <%end if%>
                    <%=rs3("title")%> </a> </td>
                </tr>
              </table>
          <%
			rs3.MoveNext
			wend
			else
			Response.Write "<tr><td width=100% align=center height=18 >暂无</td></tr>>"
			end if
			rs3.close
			set rs3=nothing
			%>
			<%end if %>
        </td>
      </tr>
    </table>
	    </td>	
         <td valign="top" width="570">
	      <table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="25" background="Images/dh_bg.gif" class="daohang">&nbsp;&nbsp;&nbsp;&nbsp;<img src="Images/arrow_dh_1.gif" width="7" height="7">&nbsp;&nbsp;&nbsp;栏目导航&nbsp;&nbsp;<a href="./" class="daohang">网站首页</a>&gt;&gt;<%=E_typename%></td>
      </tr>
      <tr>
        <td height="5"><div align="center">
          <script language=javascript src=./zongg/ad.asp?i=14></script>
        </div></td>
      </tr>
      <tr>
        <td height="25" background="Images/dh_bg.gif">&nbsp;&nbsp;&nbsp;<img src="Images/arrow_dh_1.gif" width="7" height="7">&nbsp;&nbsp;无小类文章区</td>
      </tr>
    </table>
          <!--无大类文章区开始-->
         <% select case mode%>
        <% case 1 %>
        <!--图片模版-->
        <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber5">
          <%set rsnobigclass=server.CreateObject("ADODB.RecordSet")
			if uselevel=1 then		    
				if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then
				rsnobigclass.Source="select  * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &"  and E_BigClassID is null and E_SmallClassID is null and checkked=1) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("key")="" then
					rsnobigclass.Source="select  * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &"  and E_BigClassID is null and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("key")="selfreg" then
					if Request.cookies(eChsys)("reglevel")=3 then
						rsnobigclass.Source="select  * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &"  and E_BigClassID is null and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
					end if
				if Request.cookies(eChsys)("reglevel")=2 then
					rsnobigclass.Source="select  * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &"  and E_BigClassID is null and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("reglevel")=1 then
					rsnobigclass.Source="select  * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &"  and E_BigClassID is null and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				
			end if
			else
			rsnobigclass.Source="select  * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &"  and E_BigClassID is null and E_SmallClassID is null and checkked=1) order by newsid DESC"
			end if
				rsnobigclass.Open rsnobigclass.Source,conn,1,1

if not rsnobigclass.EOF then
rsnobigclass.PageSize     = MyPageSize
MaxPages         = rsnobigclass.PageCount
rsnobigclass.absolutepage = MyPage
total            = rsnobigclass.RecordCount
i = 0

				%>
          <tr>
            <td><table width="90%" border="0" cellspacing="0" cellpadding="3" align="center">
                <!--图片换行显示1-->
                <% do until rsnobigclass.Eof or i = rsnobigclass.PageSize %>
                <tr>
                  <%
					fileExt=lcase(getFileExtName(rsnobigclass("picname")))
					Content=htmlencode4(rsnobigclass("Content"))
					content=replace(content,"[---分页---]","")
					%>
                  <td width=25% align=center valign="top" style="table-layout:fixed; word-break:break-all"><%if not rsnobigclass.EOF then%>
                      <% content=htmlencode4(rsnobigclass("Content"))
					content=replace(content,"[---分页---]","")%>
                      <a class=middle href="E_ReadNews.asp?NewsID=<%=rsnobigclass("NewsID")%>" target=_blank title="<%=CutStr(nohtml(Content),150)%>...">
                      <%if   rsnobigclass("picname")=("") then%>
                      <img  src="IMAGES/flashorno.gif" width="110" height="80" border=1 style=border-color:#000000 align=top>
                      <%else%>
                      <%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
                      <img  src="<%=FileUploadPath & rsnobigclass("picname")%>" width="110" height="80" border=1 style=border-color:#000000 align=top>
                      <%end if%>
                      <%if fileext="swf" then%>
                      <object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="110" height="80" border=0 >
                        <param name=movie value="<%=FileUploadPath & rsnobigclass("picname")%>">
                        <param name=quality value=high>
                        <param name='Play' value='-1'>
                        <param name='Loop' value='0'>
                        <param name='Menu' value='-1'>
                        <embed src="<%=FileUploadPath & rsnobigclass("picname")%>" width="110" height="80" border=0  pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
                      </object>
                      <%end if%>
                      <%end if%>
                      </a> <br>
                      <br>
                    <a class=middle href="E_ReadNews.asp?NewsID=<%=rsnobigclass("NewsID")%>" target=_blank title="<%=rsnobigclass("title")%>"><%=CutStr(rsnobigclass("title"),18)%></a> </td>
                  <%rsnobigclass.movenext
					i = i + 1
					end if %>
                  <!--图片换行显示2-->
                  <%
					fileExt=lcase(getFileExtName(rsnobigclass("picname")))
					Content=htmlencode4(rsnobigclass("Content"))
					content=replace(content,"[---分页---]","")
					%>
                  <td width=25% align=center valign="top" style="table-layout:fixed; word-break:break-all"><%if not rsnobigclass.EOF then%>
                      <%
					content=htmlencode4(rsnobigclass("Content"))
					content=replace(content,"[---分页---]","")%>
                      <a class=middle href="E_ReadNews.asp?NewsID=<%=rsnobigclass("NewsID")%>" target=_blank title="<%=CutStr(nohtml(Content),150)%>">
                      <%if rsnobigclass("picname")=("") then%>
                      <img  src="IMAGES/flashorno.gif" width="110" height="80" border=1 style=border-color:#000000 align=top>
                      <%else%>
                      <%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
                      <img  src="<%=FileUploadPath & rsnobigclass("picname")%>" width="110" height="80" border=1 style=border-color:#000000 align=top>
                      <%end if%>
                      <%if fileext="swf" then%>
                      <object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="110" height="80" border=0 >
                        <param name=movie value="<%=FileUploadPath & rsnobigclass("picname")%>">
                        <param name=quality value=high>
                        <param name='Play' value='-1'>
                        <param name='Loop' value='0'>
                        <param name='Menu' value='-1'>
                        <embed src="<%=FileUploadPath & rsnobigclass("picname")%>" width="110" height="80" border=0  pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
                      </object>
                      <%end if%>
                      <%end if%>
                      </a><br>
                      <br>
                    <a class=middle href="E_ReadNews.asp?NewsID=<%=rsnobigclass("NewsID")%>" target=_blank title="<%=rsnobigclass("title")%>"><%=CutStr(rsnobigclass("title"),18)%></a> </td>
                  <%rsnobigclass.movenext
					i = i + 1
					end if %>
                  <!--图片换行显示3-->
                  <%
					fileExt=lcase(getFileExtName(rsnobigclass("picname")))
					Content=htmlencode4(rsnobigclass("Content"))
					content=replace(content,"[---分页---]","")
					%>
                  <td width=25% align=center valign="top" style="table-layout:fixed; word-break:break-all"><%if not rsnobigclass.EOF then%>
                      <%
					content=htmlencode4(rsnobigclass("Content"))
					content=replace(content,"[---分页---]","")%>
                      <a class=middle href="E_ReadNews.asp?NewsID=<%=rsnobigclass("NewsID")%>" target=_blank title="<%=CutStr(nohtml(Content),150)%>">
                      <%if rsnobigclass("picname")=("") then%>
                      <img  src="IMAGES/flashorno.gif" width="110" height="80" border=1 style=border-color:#000000 align=top>
                      <%else%>
                      <%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
                      <img  src="<%=FileUploadPath & rsnobigclass("picname")%>" width="110" height="80" border=1 style=border-color:#000000 align=top>
                      <%end if%>
                      <%if fileext="swf" then%>
                      <object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="110" height="80" border=0 >
                        <param name=movie value="<%=FileUploadPath & rsnobigclass("picname")%>">
                        <param name=quality value=high>
                        <param name='Play' value='-1'>
                        <param name='Loop' value='0'>
                        <param name='Menu' value='-1'>
                        <embed src="<%=FileUploadPath & rsnobigclass("picname")%>" width="110" height="80" border=0  pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
                      </object>
                      <%end if%>
                      <%end if%>
                      </a> <br>
                      <br>
                    <a class=middle href="E_ReadNews.asp?NewsID=<%=rsnobigclass("NewsID")%>" target=_blank title="<%=rsnobigclass("title")%>"><%=CutStr(rsnobigclass("title"),18)%></a></td>
                  <%rsnobigclass.movenext
				i = i + 1
				end if %>
                  <!--图片换行显示4-->
                  <%
					fileExt=lcase(getFileExtName(rsnobigclass("picname")))
					Content=htmlencode4(rsnobigclass("Content"))
					content=replace(content,"[---分页---]","")
					%>
                  <td width=25% align=center valign="top" style="table-layout:fixed; word-break:break-all"><%if not rsnobigclass.EOF then%>
                      <%content=rsnobigclass("Content")
					content=replace(content,"[---分页---]","")%>
                      <a class=middle href="E_ReadNews.asp?NewsID=<%=rsnobigclass("NewsID")%>" target=_blank title="<%=CutStr(nohtml(Content),150)%>...">
                      <%if rsnobigclass("picname")=("") then%>
                      <img  src="IMAGES/flashorno.gif" width="110" height="80" border=1 style=border-color:#000000 align=top>
                      <%else%>
                      <%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
                      <img  src="<%=FileUploadPath & rsnobigclass("picname")%>" width="110" height="80" border=1 style=border-color:#000000 align=top>
                      <%end if%>
                      <%if fileext="swf" then%>
                      <object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="110" height="80" border=0 >
                        <param name=movie value="<%=FileUploadPath & rsnobigclass("picname")%>">
                        <param name=quality value=high>
                        <param name='Play' value='-1'>
                        <param name='Loop' value='0'>
                        <param name='Menu' value='-1'>
                        <embed src="<%=FileUploadPath & rsnobigclass("picname")%>" width="110" height="80" border=0  pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
                      </object>
                      <%end if%>
                      <%end if%>
                      </a> <br>
                      <br>
                    <a class=middle href="E_ReadNews.asp?NewsID=<%=rsnobigclass("NewsID")%>" target=_blank title="<%=rsnobigclass("title")%>"><%=CutStr(rsnobigclass("title"),18)%></a> </td>
                  <%rsnobigclass.movenext
					i = i + 1
					end if %>
                </tr>
                <!--图片换行结束-->
                <%loop%>
            </table></td>
          </tr>
          <tr>
            <td><table cellspacing=0 cellpadding=3 width="100%" align=center border=0 >
                <tr>
                  <td width="100%" align=center valign="middle">共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 页，每页 <%=MyPageSize%> 条
                    <%
url="E_NobigClass.asp?E_typeid=" & E_typeid & "&E_BigClassID=" & request_E_BigClassID
'request_E_SmallClassID
PageNextSize=int((MyPage-1)/PageShowSize)+1
Pagetpage=int((total-1)/rsnobigclass.PageSize)+1

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
Response.write "<table align=center><tr><td>&nbsp;本大类无小类文章区暂无信息</td></tr></table>"
				
End If
rsnobigclass.close
set rsnobigclass=nothing
%>
                  </td>
                </tr>
            </table></td>
          </tr>
        </table>
      <!--图片模版-->
        <% case 2 %>
        <!--新闻模版-->
		<%set rsnobigclass=server.CreateObject("ADODB.RecordSet")
			if uselevel=1 then
			    
				if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then
				rsnobigclass.Source="select  * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &"  and E_BigClassID is null and E_SmallClassID is null and checkked=1) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("key")="" then
					rsnobigclass.Source="select  * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &"  and E_BigClassID is null and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("key")="selfreg" then
					if Request.cookies(eChsys)("reglevel")=3 then
						rsnobigclass.Source="select  * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &"  and E_BigClassID is null and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
					end if
				if Request.cookies(eChsys)("reglevel")=2 then
					rsnobigclass.Source="select  * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &"  and E_BigClassID is null and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("reglevel")=1 then
					rsnobigclass.Source="select  * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &"  and E_BigClassID is null and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				
			end if
			else
			rsnobigclass.Source="select  * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &"  and E_BigClassID is null and E_SmallClassID is null and checkked=1) order by newsid DESC"
			end if
				rsnobigclass.Open rsnobigclass.Source,conn,1,1

if not rsnobigclass.EOF then
rsnobigclass.PageSize     = MyPageSize
MaxPages         = rsnobigclass.PageCount
rsnobigclass.absolutepage = MyPage
total            = rsnobigclass.RecordCount

i = 0
do until rsnobigclass.Eof or i = rsnobigclass.PageSize

				if showyear=1 then
					newsurl="E_ReadNews.asp?NewsID=" & rsnobigclass("NewsID")
					newswwwurl=rsnobigclass("titleface")
					datetime="<font class=middle>(" & year(rsnobigclass("UpdateTime"))  &"年"& Month(rsnobigclass("UpdateTime"))  &"月"& Day(rsnobigclass("UpdateTime")) &"日)</font>"
				else
					newsurl="E_ReadNews.asp?NewsID=" & rsnobigclass("NewsID")
					newswwwurl=rsnobigclass("titleface")
					datetime="<font class=middle>("& Month(rsnobigclass("UpdateTime"))  &"月"& Day(rsnobigclass("UpdateTime")) &"日)</font>"
				end if
				if rsnobigclass("picnews")=1 then
					img="<font class=pic_word>[图]</font>"
				else
					img=""
				end if
					title=trim(rsnobigclass("title"))
					title=replace(title,"<br>","")
		%>
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td></td>
          </tr>
          <tr>
            <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="69%">&nbsp;&nbsp;<img src="IMAGES/006.gif" width="10" height="10"> <%=img%> <a class=middle href="<%if rsnobigclass("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" title="<%=rsnobigclass("title")%>" target="_blank"> <<%=rsnobigclass("titletype")%>><font color="<%=rsnobigclass("titlecolor")%>"> <%=CutStr(title,46)%></font></<%=rsnobigclass("titletype")%>></a>
                      <!--标题后评论提示-->
                      <% if rsnobigclass("titlesize")>=1 then %>
                      <a class=middle href="<%=path%>review.asp?NewsID=<%=rsnobigclass("NewsID")%>" target="_blank" ><font color=red><b>评</b></font></a>
                      <%end if %>
                      <!--标题后评论提示-->
                  </td>
                  <td width="17%" align="left"><div align="right">
                    <%if showtime="1" then%>
                      <%=datetime%>
                      <%end if%>
                  &nbsp; </div></td>
                  <td width="14%" align="left"><%if year(rsnobigclass("updatetime"))=year(date()) and month(rsnobigclass("updatetime"))=month(date()) and day(rsnobigclass("updatetime"))=day(date()) then%>
                      <img src="images/new.gif">
                      <%end if%>
                      <%if rsnobigclass("goodnews")="1" then%>
                      <img src="images/g.gif" >
                      <%end if%>
                  </td>
                </tr>
            </table></td>
          </tr>
        </table>
      <%
		rsnobigclass.MoveNext
		i = i + 1
		loop
		%>
        <table cellspacing=0 cellpadding=3 width="100%" align=center border=0 >
          <tr>
            <td width="100%" align=center valign="middle">共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 页，每页 <%=MyPageSize%> 条
              <%
url="E_NobigClass.asp?E_typeid=" & E_typeid & "&E_BigClassID=" & request_E_BigClassID
'request_E_SmallClassID
PageNextSize=int((MyPage-1)/PageShowSize)+1
Pagetpage=int((total-1)/rsnobigclass.PageSize)+1

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
Response.write "<table align=center><tr><td>&nbsp;本大类无小类文章区暂无信息</td></tr></table>"
			
End If
rsnobigclass.close
set rsnobigclass=nothing
%>
            </td>
          </tr>
        </table>
      <!--新闻模版-->
        <% case 3 %>
        <!--网址模版-->
        <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber5">
          <%set rsnobigclass=server.CreateObject("ADODB.RecordSet")
		if uselevel=1 then
			if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then
				rsnobigclass.Source="select  * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &"  and E_BigClassID is null and E_SmallClassID is null and checkked=1) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("key")="" then
					rsnobigclass.Source="select  * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &"  and E_BigClassID is null and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("key")="selfreg" then
					if Request.cookies(eChsys)("reglevel")=3 then
						rsnobigclass.Source="select  * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &"  and E_BigClassID is null and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
					end if
				if Request.cookies(eChsys)("reglevel")=2 then
					rsnobigclass.Source="select  * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &"  and E_BigClassID is null and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("reglevel")=1 then
					rsnobigclass.Source="select  * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &"  and E_BigClassID is null and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
			end if
			else
			rsnobigclass.Source="select  * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &"  and E_BigClassID is null and E_SmallClassID is null and checkked=1) order by newsid DESC"
			end if
			rsnobigclass.Open rsnobigclass.Source,conn,1,1

if not rsnobigclass.EOF then
rsnobigclass.PageSize     = MyPageSize
MaxPages         = rsnobigclass.PageCount
rsnobigclass.absolutepage = MyPage
total            = rsnobigclass.RecordCount
i = 0

		%>
          <tr>
            <td><table width="98%" border="1" cellspacing="0" cellpadding="3" bordercolorlight=#cccccc  bordercolordark=#ffffff align="center">
                <!--网址换行显示1-->
                <% do until rsnobigclass.Eof or i = rsnobigclass.PageSize %>
                <tr>
                  <%
						Content=htmlencode4(rsnobigclass("Content"))
					%>
                  <% if not rsnobigclass.EOF then %>
                  <td width=25% align=center valign="middle"><div align="center"> <a class=middle href="<%=rsnobigclass("Original")%>" target=_blank title="<%=CutStr(nohtml(rsnobigclass("Content")),80)%>"><%=CutStr(rsnobigclass("title"),30)%></a> </div></td>
                  <%rsnobigclass.movenext
					i = i + 1
					end if %>
                  <!--网址换行显示2-->
                  <td width=25% align=center valign="middle"><%if not rsnobigclass.EOF then%>
                      <div align="center"> <a class=middle href="<%=rsnobigclass("Original")%>" target=_blank title="<%=CutStr(nohtml(rsnobigclass("Content")),80)%>"><%=CutStr(rsnobigclass("title"),30)%></a> </div></td>
                  <%rsnobigclass.movenext
					i = i + 1
					end if %>
                </tr>
                <!--网址换行结束-->
                <%loop%>
            </table></td>
          <tr>
            <td><table cellspacing=0 cellpadding=3 width="100%" align=center border=0 >
                <tr>
                  <td width="100%" align=center valign="middle">共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 页，每页 <%=MyPageSize%> 条
                    <%
url="E_NobigClass.asp?E_typeid=" & E_typeid & "&E_BigClassID=" & request_E_BigClassID
'request_E_SmallClassID
PageNextSize=int((MyPage-1)/PageShowSize)+1
Pagetpage=int((total-1)/rsnobigclass.PageSize)+1

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
Response.write "<table align=center><tr><td>&nbsp;本大类无小类文章区暂无信息</td></tr></table>"
				
End If

rsnobigclass.close
			set rsnobigclass=nothing
%>
                  </td>
                </table>
            </tr>
        </table>
      <!--网址模版-->
        <% case 4 %>
        <!--软件模版-->
        <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse"  width="100%" id="AutoNumber5">
          <%set rsnobigclass=server.CreateObject("ADODB.RecordSet")
			if uselevel=1 then
				if Request.cookies(eChsys)("key")="super" or Request.cookies(eChsys)("key")="typemaster" or Request.cookies(eChsys)("key")="bigmaster" or Request.cookies(eChsys)("key")="smallmaster" or Request.cookies(eChsys)("key")="check" then
				rsnobigclass.Source="select  * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &"  and E_BigClassID is null and E_SmallClassID is null and checkked=1) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("key")="" then
					rsnobigclass.Source="select  * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &"  and E_BigClassID is null and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("key")="selfreg" then
					if Request.cookies(eChsys)("reglevel")=3 then
						rsnobigclass.Source="select  * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &"  and E_BigClassID is null and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
					end if
				if Request.cookies(eChsys)("reglevel")=2 then
					rsnobigclass.Source="select  * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &"  and E_BigClassID is null and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				if Request.cookies(eChsys)("reglevel")=1 then
					rsnobigclass.Source="select  * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &"  and E_BigClassID is null and E_SmallClassID is null and checkked=1 ) order by newsid DESC"
				end if
				end if
				else
					rsnobigclass.Source="select  * from "& db_EC_News_Table &" where (E_typeid=" & E_typeid &"  and E_BigClassID is null and E_SmallClassID is null and checkked=1) order by newsid DESC"
				end if
			rsnobigclass.Open rsnobigclass.Source,conn,1,1


if not rsnobigclass.EOF then
rsnobigclass.PageSize     = MyPageSize
MaxPages         = rsnobigclass.PageCount
rsnobigclass.absolutepage = MyPage
total            = rsnobigclass.RecordCount

i = 0
do until rsnobigclass.Eof or i = rsnobigclass.PageSize

			    

					if showyear=1 then
						newsurl="E_ReadNews.asp?NewsID=" & rsnobigclass("NewsID")
						newswwwurl=rsnobigclass("titleface")
						datetime="<font class=middle>(" & year(rsnobigclass("UpdateTime"))  &"年"& Month(rsnobigclass("UpdateTime"))  &"月"& Day(rsnobigclass("UpdateTime")) &"日)</font>"
					else
						newsurl="E_ReadNews.asp?NewsID=" & rsnobigclass("NewsID")
						newswwwurl=rsnobigclass("titleface")
						datetime="<font class=middle>("& Month(rsnobigclass("UpdateTime"))  &"月"& Day(rsnobigclass("UpdateTime")) &"日)</font>"
					end if
					if rsnobigclass("picnews")=1 then
						img="<img src='images/news_img.gif' border='0'>"
					else
						img=""
					end if
						title=trim(rsnobigclass("title"))
						title=replace(title,"<br>","")
			%>
          <tr>
            <%
		fileExt=lcase(getFileExtName(rsnobigclass("picname")))
		Content=htmlencode4(rsnobigclass("Content"))
		content=replace(content,"[---分页---]","")%>
            <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr bgcolor="#EFEFEF">
                  <td colspan="2">&nbsp;<img src="images/news_img.gif" width="9" height="9"> <a class=middle href="<%if rsnobigclass("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" title="<%=htmlencode4(title)%>" target="_blank"> <font color="<%=rsnobigclass("titlecolor")%>"><strong> <%=CutStr(title,40)%> </strong></font> </a> </td>
                  <td width="17%" align="right" bgcolor="#EFEFEF"><%if showtime="1" then%>
                      <%=datetime%>
                      <%end if%>
                  &nbsp;</td>
                  <td width="14%" align="center"><%if year(rsnobigclass("updatetime"))=year(date()) and month(rsnobigclass("updatetime"))=month(date()) and day(rsnobigclass("updatetime"))=day(date()) then%>
                      <img src="images/new.gif">
                      <%end if%>
                      <%if rsnobigclass("goodnews")="1" then%>
                      <img src="images/g.gif" >
                      <%end if%>
                  </td>
                </tr>
                <tr>
                  <td width="36%"><a class=middle href="<%if rsnobigclass("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" target=_blank title="<%=htmlencode4(rsnobigclass("title"))%>">
                    <%if   rsnobigclass("picname")=("")   then%>
                    <img  src="IMAGES/softno.gif" width="65" height="65" border=0 align="left">
                    <%else%>
                    <%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
                    <img  src="<%=FileUploadPath & rsnobigclass("picname")%>" width="65" height="65" border=0 align="left">
                    <%end if%>
                    <%if fileext="swf" then%>
                    <object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="65" height="65" border=0 >
                      <param name=movie value="<%=FileUploadPath & rsnobigclass("picname")%>">
                      <param name=quality value=high>
                      <param name='Play' value='-1'>
                      <param name='Loop' value='0'>
                      <param name='Menu' value='-1'>
                      <embed src="<%=FileUploadPath & rsnobigclass("picname")%>" width="6" 5height="65" border=0  pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
                    </object>
                    <%end if%>
                    <%end if%>
                  </a> </td>
                  <td colspan="3" valign="top"><a class=middle href="<%if rsnobigclass("titleface")="无" then %><%=newsurl%><% else %> <%=newswwwurl%><%end if%>" target="_blank" title="<%=htmlencode4(title)%>"><%=CutStr(nohtml(rsnobigclass("Content")),250)%></a> </td>
                </tr>
            </table></td>
          </tr>
          <%
		rsnobigclass.MoveNext
		i = i + 1
		loop
		%>
        </table>
      <table cellspacing=0 cellpadding=3 width="100%" align=center border=0 >
          <tr>
            <td width="100%" align=center valign="middle">共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 页，每页 <%=MyPageSize%> 条
              <%
url="E_NobigClass.asp?E_typeid=" & E_typeid & "&E_BigClassID=" & request_E_BigClassID
'request_E_SmallClassID
PageNextSize=int((MyPage-1)/PageShowSize)+1
Pagetpage=int((total-1)/rsnobigclass.PageSize)+1

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
Response.write "<table align=center><tr><td>&nbsp;本大类无小类文章区暂无信息</td></tr></table>"		
End If
rsnobigclass.close
set rsnobigclass=nothing
%>
            </td>
          </tr>
        </table>
      <!--软件模版-->
        <%end select%>
        <!--无大类文章区结束-->
         </td>
       </tr>
     </table>
<!--#include file="other_bottom.asp" -->
</body>
</html>