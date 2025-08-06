<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="chkuser.asp"-->
<%dim jingyong
E_SmallClassID=ChkRequest(request("E_SmallClassID"),1)	'防注入
E_BigClassID=ChkRequest(request("E_BigClassID"),1)	'防注入
E_typeid=ChkRequest(request("E_typeid"),1)	'防注入
set rs=server.createobject("adodb.recordset")
sql="select * from "& db_User_Table &" where  "& db_User_Name &"='"&request.cookies(eChsys)("username")&"'"
rs.open sql,ConnUser,1,3
jingyong=rs("jingyong")
rs.close
set rs=nothing

set rs8=server.createobject("adodb.recordset")
sql8="select * from "& db_EC_System_Table
rs8.open sql8,conn,1,3
name=rs8("name")

randomize
rannum=int(90000*rnd)+10000
Response.cookies(eChsys)("newsrelated")=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&rannum

if request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="check" or request.cookies(eChsys)("KEY")="bigmaster" or request.cookies(eChsys)("KEY")="smallmaster" or (request.cookies(eChsys)("key")="selfreg" and jingyong=0) or request.cookies(eChsys)("KEY")="typemaster" then%>

<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 添加文章</title>
<script src="ckeditor/ckeditor.js"></script>
<script language="javascript">
<!--
function checkdata()
{
if (document.form1.title.value=="")
	{
	  alert("对不起，请输入文章标题！")
	  document.form1.title.focus()
	  return false
	 }
}

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>

<link rel="stylesheet" type="text/css" href="site.css">
<link rel="STYLESHEET" type="text/css" href="editor.css">
</head>
<body topmargin="0">

<div id=menuDiv style='Z-INDEX: 1000; VISIBILITY: hidden; WIDTH: 1px; POSITION: absolute; HEIGHT: 1px; BACKGROUND-COLOR: #9cc5f8'></div>
<%set rs5=server.CreateObject ("ADODB.RecordSet")
rs5.Source="select * from "& db_EC_System_Table &""
rs5.Open rs5.source,conn,1,1
%>
<div align="center"><center>
<table border="0" cellspacing="0" width="100%" cellpadding="0">
<form name="form1" method="post" action="NewsAdd2.asp" OnSubmit="return checkdata()" onReset="return ResetForm();">
<tr align="center"> 
  <td width="100%"> 
	<table border="1" cellspacing="0" width="100%" cellpadding="0"  bordercolor="<%=border%>" style="border-collapse: collapse" height="463">
	<tr align="center">
<%if E_SmallClassID<>"" then
dim E_typename,E_BigClassName,E_smallclassname
set rs1=server.createobject("adodb.recordset")
sql1="select * from "& db_EC_SmallClass_Table &" where E_SmallClassID="&E_SmallClassID
rs1.open sql1,conn,1,3
E_smallclassname=rs1("E_smallclassname")
set rs2=server.createobject("adodb.recordset")
sql2="select * from "& db_EC_BigClass_Table &" where E_BigClassID="&rs1("E_BigClassID")
rs2.open sql2,conn,1,3
E_BigClassName=rs2("E_BigClassName")
set rs3=server.createobject("adodb.recordset")
sql3="select * from "& db_EC_Type_Table &" where E_typeid="&rs2("E_typeid")
rs3.open sql3,conn,1,3
E_typename=rs3("E_typename")
%>
	<input name="E_typeid" type="hidden" value="<%=rs2("E_typeid")%>">
	<input name="E_BigClassID" type="hidden" value="<%=rs1("E_BigClassID")%>">
	<input name="E_SmallClassID" type="hidden" value="<%=E_SmallClassID%>">
	<td colspan="2" height="25"  valign="middle" class="TDtop">
		<div align="center" >┊ <B>在<b class="unnamed2">[<a href="E_BigClassManage.asp?E_typeid=<%=rs2("E_typeid")%>"><%=E_typename%></a>]-[<a href="E_SmallClassManage.asp?E_BigClassID=<%=rs1("E_BigClassID")%>"><%=E_BigClassName%></a>]-[<a href="ListNews.asp?E_SmallClassID=<%=E_SmallClassID%>"><%=E_smallclassname%></a>]</b>中添 加 文 章</B> ┊</div>
	</td>
<%rs3.close
set rs3=nothing
rs2.close
set rs2=nothing
rs1.close
set rs1=nothing
else%>
<%if E_BigClassID<>"" then
set rs1=server.createobject("adodb.recordset")
sql1="select * from "& db_EC_BigClass_Table &" where E_BigClassID="&E_BigClassID
rs1.open sql1,conn,1,3
E_BigClassName=rs1("E_BigClassName")
bigclasszs=rs1("bigclasszs")
set rs2=server.createobject("adodb.recordset")
sql2="select * from "& db_EC_Type_Table &" where E_typeid="&rs1("E_typeid")
rs2.open sql2,conn,1,3
E_typename=rs2("E_typename")
%>
	<input name="E_typeid" type="hidden" value="<%=rs1("E_typeid")%>">
	<input name="E_BigClassID" type="hidden" value="<%=E_BigClassID%>">
	<td colspan="2"  height="29"  class="TDtop" valign="middle" bgcolor="<%=m_top%>" width="100%">在<b class="unnamed2">
		<a href="E_BigClassManage.asp?E_typeid=<%=rs2("E_typeid")%>"><%=E_typename%></a>-<a href="E_SmallClassManage.asp?E_BigClassID=<%=rs1("E_BigClassID")%>"><%=E_BigClassName%></a></b>中添 加 文 章
	</td>
<%rs2.close
set rs2=nothing
rs1.close
set rs1=nothing
else
%>

<%
if E_typeid="" then
response.write "<script>alert('请至少选择总栏！');history.go(-1);</Script>"
response.end
else
set rs1=server.createobject("adodb.recordset")
sql1="select * from "& db_EC_Type_Table &" where E_typeid="&E_typeid
rs1.open sql1,conn,1,3
E_typename=rs1("E_typename")
mode=rs1("mode")
%>
	<input name="E_typeid" type="hidden" value="<%=E_typeid%>">
	
	<td colspan="2"  height="29"  class="TDtop" valign="middle" bgcolor="<%=m_top%>" width="100%">在<b class="unnamed2">
		<a href="E_BigClassManage.asp?E_typeid=<%=rs1("E_typeid")%>"><%=E_typename%></a></b>中添 加 文 章
	</td>
<%
rs1.close
set rs1=nothing
end if 
end if
end if

%>

	</tr>
<%if mode<>5 and bigclasszs<>5 then%>
	<tr> 
	<td width="10%" align="right" class="unnamed2" valign="middle" bgcolor="#FFFFFF" >所属专题：</td>
	<td width="90%" bgcolor="#FFFFFF" height="25"> &nbsp; 
	<select name="SpecialID" size="1" onMouseOver="window.status='点击这里选择文章专题';return true;" onMouseOut="window.status='';return true;">
                    <%			
set rs3=server.CreateObject ("ADODB.RecordSet")
rs3.Source="select * from "& db_EC_Special_Table &""
rs3.Open rs3.source,conn,1,1			
%>
	<option value=0>不属于任何专题</option>
                    <%if rs3.EOF then %>
	<option value=0>暂无任何专题</option>
                    <%else
while not rs3.EOF
%>
	<option value="<%=rs3("SpecialID")%>"><%=trim(rs3("E_SpecialName"))%></option>
                    <%
rs3.MoveNext
wend
end if
rs3.close	
%>
	</select>
&nbsp;第二专题：<select name="SpecialID2" size="1" onMouseOver="window.status='点击这里选择文章专题';return true;" onMouseOut="window.status='';return true;">
                    <%			
set rs3=server.CreateObject ("ADODB.RecordSet")
rs3.Source="select * from "& db_EC_Special_Table &""
rs3.Open rs3.source,conn,1,1			
%>
	<option value=0>不属于任何专题</option>
                    <%if rs3.EOF then %>
	<option value=0>暂无任何专题</option>
                    <%else
while not rs3.EOF
%>
	<option value="<%=rs3("SpecialID")%>"><%=trim(rs3("E_SpecialName"))%></option>
                    <%
rs3.MoveNext
wend
end if
rs3.close	
%>
	</select>
	</td>
	</tr>
	<%end if%>
	<tr height="25"> 
	<td align="right" class="unnamed2" valign="middle" bgcolor="#FFFFFF">文章标题：</td>
	<td bgcolor="#FFFFFF"> <span style='cursor:hand' title='缩短对话框' onClick='if (me.size>10)me.size=me.size-2'>-</span> 
	<input name="title" id=me type="TEXT" size=60 maxlength=100 style=" height:22;background-color:ffffff;border: 1 double; border-color:#BDD0FD;"  onMouseOver="window.status='在这里输入要添加的文章标题，必填';return true;" onMouseOut="window.status='';return true;" title='在这里输入要添加的文章标题，必填'>
	<span style='cursor:hand' title='加长对话框' onClick='if (me.size<102)me.size=me.size+2'>+</span> <font color="#FF0000">*</font>
	<SELECT size=1 id=title_color name=title_color LANGUAGE=javascript onChange="return title_color_onchange()" style="width:55">
	<OPTION value="" selected>颜色</OPTION>
	<OPTION value="">缺省</OPTION>
	<OPTION value="#000000" style="background-color:#000000"></OPTION>
	<OPTION value="#FFFFFF" style="background-color:#FFFFFF"></OPTION>
	<OPTION value="#008000" style="background-color:#008000"></OPTION>
	<OPTION value="#800000" style="background-color:#800000"></OPTION>
	<OPTION value="#808000" style="background-color:#808000"></OPTION>
	<OPTION value="#000080" style="background-color:#000080"></OPTION>
	<OPTION value="#800080" style="background-color:#800080"></OPTION>
	<OPTION value="#808080" style="background-color:#808080"></OPTION>
	<OPTION value="#FFFF00" style="background-color:#FFFF00"></OPTION>
	<OPTION value="#00FF00" style="background-color:#00FF00"></OPTION>
	<OPTION value="#00FFFF" style="background-color:#00FFFF"></OPTION>
	<OPTION value="#FF00FF" style="background-color:#FF00FF"></OPTION>
	<OPTION value="#FF0000" style="background-color:#FF0000"></OPTION>
	<OPTION value="#0000FF" style="background-color:#0000FF"></OPTION>
	<OPTION value="#008080" style="background-color:#008080"></OPTION>
	</SELECT>

	<SELECT size=1 id=title_type name=title_type LANGUAGE=javascript onChange="return title_type_onchange()" style="width:55">
	<OPTION value="0" selected>类型</OPTION>
	<OPTION value="0">普通</OPTION>
	<OPTION value="b">加粗</OPTION>
	<OPTION value="i">倾斜</OPTION>
	<OPTION value="u">下划线</OPTION>
	</SELECT>

	</SELECT>
	<SELECT size=1 id=title_size name=title_size LANGUAGE=javascript onChange="return title_size_onchange()" style="width:55">
	<OPTION value="0" selected>不评</OPTION>
	<OPTION value="0">不评</OPTION>
	<OPTION value="1">要评</OPTION>
	</SELECT>

	</td>
	</tr>
<%if mode<>5 and bigclasszs<>5 then%>
	<tr height="25"> 
	<td align="right" class="unnamed2" valign="middle" bgcolor="#FFFFFF" >链接文章：</td>
	<td bgcolor="#FFFFFF" > <span style='cursor:hand' title='缩短对话框' onClick='if (title_face.size>10)title_face.size=title_face.size-2'>-</span> 
	<input name="title_face" id=title_face type="TEXT" size=30 maxlength=100 style=" height:22;background-color:ffffff;border: 1 double; border-color:#BDD0FD;" onMouseOver="window.status='文章为直接链接文章时，请填写网址';return true;" onMouseOut="window.status='';return true;" title='文章为直接链接文章时，请填写网址'>
	<span style='cursor:hand' title='加长对话框' onClick='if (title_face.size<102)title_face.size=title_face.size+2'>+</span> 
	(文章为直接链接文章时，请填写网址)</td>
	</tr>
	<tr height="25"> 
                <td align="right" class="unnamed2" valign="middle" bgcolor="#FFFFFF" >文章来源：</td>
                <td bgcolor="#FFFFFF" > <span style='cursor:hand' title='缩短对话框' onClick='if (message.size>10)message.size=message.size-2'>-</span> 
                  <input name="Original" id=message type="TEXT" size=30 maxlength=100 style=" height:22;background-color:ffffff;border: 1 double; border-color:#BDD0FD;" onMouseOver="window.status='如果有，在这里输入文章来源';return true;" onMouseOut="window.status='';return true;" title='如果有，在这里输入文章来源'>
                  <span style='cursor:hand' title='加长对话框' onClick='if (message.size<102)message.size=message.size+2'>+</span> 
                (网页模版时，请填写网址)</td>
              </tr>
              <tr height="25"> 
                <td align="right" class="unnamed2" valign="middle" bgcolor="#FFFFFF" >文章作者：</td>
                <td bgcolor="#FFFFFF" ><span style='cursor:hand' title='缩短对话框' onClick='if (mess.size>10)mess.size=mess.size-2'>-</span> 
                  <input value="" name="Author" id=mess type="TEXT" size=30 maxlength=100 style=" height:22;background-color:ffffff;border: 1 double; border-color:#BDD0FD;" onMouseOver="window.status='如果有，在这里输入文章作者';return true;" onMouseOut="window.status='';return true;" title='如果有，在这里输入文章作者'>
                  <span style='cursor:hand' title='加长对话框' onClick='if (mess.size<102)mess.size=mess.size+2'>+</span> 
                </td>
              </tr>
<%end if%>
              <tr> 
                <%
    '此逻辑替代了 TextBox.asp
    Dim strContentForEditor
    Dim p_newsID
    p_newsID = ChkRequest(Request.QueryString("newsID"),1)

    If p_newsID <> "" Then
        set rs_content=server.createobject("adodb.recordset")
        sql_content="select * from "& db_EC_News_Table &" where newsid=" & p_newsID
        rs_content.open sql_content,conn,1,1
        If Not rs_content.Eof Then
            strContentForEditor=rs_content("Content")
            '图片路径替换，与原始 TextBox.asp 中的逻辑相同
            strContentForEditor=replace(strContentForEditor,"="&chr(34)&"uploadfile/","="&chr(34)&FileUploadPath,1,-1,1) 
        End If
        rs_content.close
        set rs_content = nothing
    Else
        ' 对于新文章，如果需要，可以从 cookie 加载，或留空
        strContentForEditor = request.cookies(eChsys)("content") 
    End If
%>

<td align="right" valign="top" class="unnamed2" bgcolor="#FFFFFF" height="145">
    文章内容：<br><font color="#FF0000">*</font>
</td>
<td bgcolor="#FFFFFF">
    <%-- 用于 CKEditor 的新文本域 --%>
    <textarea name="content" id="content" rows="10" cols="80">
        <%=strContentForEditor%>
    </textarea>
</td>
              </tr>
			  <%if mode<>5 and bigclasszs<>5 then%>
              <tr height="25"> 
                <td align="right" class="unnamed2" valign="middle" bgcolor="#FFFFFF" >相关文章：</td>
                <td bgcolor="#FFFFFF" ><span style='cursor:hand' title='缩短对话框' onclick='if (ss.size>10)ss.size=ss.size-2'>-</span> 
                  <INPUT NAME="about" id=ss TYPE="TEXT" SIZE=30 maxlength=100 style=" height:22;background-color:ffffff;border: 1 double; border-color:#BDD0FD;" onMouseOver="window.status='如果有，在这里输入相关文章的关键字或完整标题';return true;" onMouseOut="window.status='';return true;" title='如果有，在这里输入相关文章的关键字或完整标题'>
                  <span style='cursor:hand' title='加长对话框' onclick='if (ss.size<102)ss.size=ss.size+2'>+</span> 
                  填入关键字或完整标题 </td>
              </tr><tr> 
                <td align="right" class="unnamed2" valign="middle" bgcolor="#FFFFFF">首页图片：</td>
                <td bgcolor="#FFFFFF">&nbsp;&nbsp;<input name="PicUrl" type="text" id="PicUrl" size="30" maxlength="200" style=" height:22;background-color:ffffff;border: 1 double; border-color:#BDD0FD;">
              用于在首页的图片文章处显示或者直接从上传图片中选择：<select name="PicList" id="PicList" onChange="PicUrl.value=this.value;">
                <option selected>不指定首页图片</option>
              </select>
              </tr>
			  <tr height="25"> 
                <td align="right" valign="middle" class="unnamed2" bgcolor="#FFFFFF" >保存图片：</td>
                <td valign="middle" bgcolor="#FFFFFF" >&nbsp;
				<input type="radio" value="1" name="getphoto" onMouseOver="window.status='点击这里，将保存远程图片';return true;" onMouseOut="window.status='';return true;" title='点击这里，将保存远程图片'>
                  是 
                  <input type="radio" value="0" checked name="getphoto">
                  否&nbsp;&nbsp;保存远程图片&nbsp;&nbsp;(从其它网站上复制内容，并且内容中包含有图片)</td>
              </tr>

              <%if request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="bigmaster" or request.cookies(eChsys)("key")="typemaster" then%>
              <tr height="25"> 
                <td align="right" valign="middle" class="unnamed2" bgcolor="#FFFFFF" >推荐文章：</td>
                <td valign="middle" bgcolor="#FFFFFF" >&nbsp; 
                  <input type="radio" value="1" name="goodnews" onMouseOver="window.status='点击这里，将生成推荐文章';return true;" onMouseOut="window.status='';return true;" title='点击这里，将生成推荐文章'>
                  是 
                  <input type="radio" value="0" checked name="goodnews">
                  否&nbsp;&nbsp;生成推荐文章</td>
              </tr>
			    <tr height="25"> 
                <td align="right" valign="middle" class="unnamed2" bgcolor="#FFFFFF" >大标条头：</td>
                <td valign="middle" bgcolor="#FFFFFF" >&nbsp; 
                  <input type="radio" value="1" name="daodu" onMouseOver="window.status='点击这里，将生成大标题导读';return true;" onMouseOut="window.status='';return true;" title='点击这里，将生成大标题导读'>
                  是 
                  <input type="radio" value="0" checked name="daodu">
                  否&nbsp;&nbsp;生成大标条头</td>
              </tr>
<tr height="25"> 
                <td align="right" valign="middle" class="unnamed2" bgcolor="#FFFFFF" >文章固顶：</td>
                <td valign="middle" bgcolor="#FFFFFF" >&nbsp; 
                  <input type="radio" value="1" name="istop" onMouseOver="window.status='点击这里，将生成固顶文章';return true;" onMouseOut="window.status='';return true;" title='点击这里，将生成固顶文章'>
                  是 
                  <input type="radio" value="0" checked name="istop">
                  否&nbsp;&nbsp;生成固顶文章</td>
              </tr>
<tr height="25"> 
                <td align="right" valign="middle" class="unnamed2" bgcolor="#FFFFFF" >自动缩进：</td>
                <td valign="middle" bgcolor="#FFFFFF" >&nbsp; 
                  <input type="radio" value="1" name="backtype" onMouseOver="window.status='点击这里，文章首行缩进排版';return true;" onMouseOut="window.status='';return true;" title='点击这里，文章首行缩进排版'>
                  是 
                  <input type="radio" value="0" checked name="backtype">
                  否&nbsp;&nbsp;文章首行缩进排版</td>
</tr>
<tr height="25"> 
                <td align="right" valign="middle" class="unnamed2" bgcolor="#FFFFFF" >内部文章：</td>
                <td valign="middle" bgcolor="#FFFFFF" >&nbsp; 
                  <input type="radio" value="1" name="backtype1" onMouseOver="window.status='点击这里，设为内部浏览文章';return true;" onMouseOut="window.status='';return true;" title='点击这里，设为内部浏览文章'>
                  是 
                  <input type="radio" value="0" checked name="backtype1">
                  否&nbsp;&nbsp;内部浏览文章</td>
</tr>

<%if uselevel=1 then%>
<tr height="25"> 
                <td align="right" bgcolor="#FFFFFF" height="30">文章级别：</td>
                <td bgcolor="#FFFFFF" height="30"> &nbsp; <select size="1" name="newslevel" onMouseOver="window.status='在这里选择您要添加的文章分级浏览级别';return true;" onMouseOut="window.status='';return true;">
<%for i=0 to 3%>
                    <option value="<%=i%>"><%=i%></option>
<%next%>
                    </select>文章级别指可浏览该文章会员的级别。0为游客，1为普通会员，2为高级会员，3为特级会员。
                  </td>
              </tr>
              <%end if%><%end if%><%end if%>
<tr align="center" height="25"> 
<td colspan="2" height="25" width="100%" clacc="TDtop">
	<input name="newsrelated" type="hidden" value="<%=Request.cookies(eChsys)("newsrelated")%>">
	<input name="checkked" type="hidden" value="<%if request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="check" or request.cookies(eChsys)("key")="bigmaster" or request.cookies(eChsys)("key")="typemaster" then%>1<%else%><%if request.cookies(eChsys)("key")="selfreg" and fabiaocheck="1" then%>1<%else%><%if request.cookies(eChsys)("key")="smallmaster" and rs5("usecheck")="1" then%>1<%else%>0<%end if%><%end if%><%end if%>"><%rs5.close%>
	<input name="encode" type="hidden" value="html">
	<input name="editor" type="hidden" value="<%=request.cookies(eChsys)("username")%>">
	<input type="button" value=" 返 回 " onClick="javascript:history.go(-1)" class="unnamed5" >&nbsp;&nbsp;
	<input type="submit" value=" 添 加 " name="Submit" class="unnamed5" >&nbsp; 
	<input type="reset" value=" 清 除 " name="Reset" class="unnamed5" >
 <%rs8.close
set rs8=nothing%>
</td>
</tr>
</table>
</td>
</tr>

</form>
</table>
</center>
</div>
<script>
    CKEDITOR.replace('content', {
        height: 350 // 你可以在这里设置高度
    });
</script>
</body>
</html>
<%end if%>
<!--#include file="Admin_Bottom.asp"-->