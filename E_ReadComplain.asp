<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<%
ComplainID=Request.QueryString("ComplainID")%>
<%
    ComplainID=ChkRequest(request("ComplainID"),1)	'防注入
	if ComplainID="" then
	Show_Err("请输入您要要查看的投诉举报ID!<br><br>－－－<a href='javascript:history.back()'>回去重来</a>－－－")
	response.end
	end if	
	
	set rs=server.CreateObject("ADODB.RecordSet")
	rs.Source="select * from "& db_EC_Complain_Table &" where  ComplainID="&ComplainID	
	rs.Open rs.Source,conn,1,1	
		
	if rs.EOF then
	Show_Err("非法投诉举报ID，请确认ComplainID是否存在!<br><br>－<a href='javascript:history.back()'>回去重来</a>－")
	Response.End
	end if
		%>
<html>
<head>
<title>阅读投诉举报__<%=eChSys%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="news.css" rel=stylesheet type=text/css>
</head>
<body bgcolor="#FFFFFF" text="#000000" topmargin="0">
<!--#include file="other_Top.asp"-->
<table width="780" height="417" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td width="780" height="417" valign="top"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="9%" height="53" valign="middle" background="Images/location_bg.gif" align="center" >&nbsp;<img src="Images/point_location.gif" width="24" height="24"></td>
        <td width="91%" valign="middle" background="Images/location_bg.gif" class="daohang">&nbsp;您的位置：&nbsp;<a class=daohang href="./" >网站首页</a>&gt;&gt;<a href="Complain.asp" class="daohang">投诉举报</a>&gt;&gt;查看投诉</td>
      </tr>
    </table>
        <table width="98%" height="364" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="349" valign="top"><div align="center">
                <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#E6D9A4">
                  <tr class="TDtop">
                    <td height="27" colspan="2" bgcolor="#FCFDEB" class="TDtop"><div align="center" class="top1">查看投诉内容</div></td>
                  </tr>
                  <tr>
                    <td width="14%" height="25" bgcolor="#FCFDEB"><div align="right" class="top1">投诉日期：</div></td>
                    <td width="86%" bgcolor="#FCFDEB"><%=year(rs("UpdateTime")) %>-<%=Month(rs("UpdateTime")) %>-<%=Day(rs("UpdateTime")) %></td>
                  </tr>
                  <tr>
                    <td height="23" bgcolor="#FCFDEB"><div align="right" class="top1">您的姓名：</div></td>
                    <td bgcolor="#FCFDEB"><%=trim(rs("YourName"))%></td>
                  </tr>
                  <tr>
                    <td height="25" bgcolor="#FCFDEB"><div align="right" class="top1">联系邮箱：</div></td>
                    <td bgcolor="#FCFDEB"><a class=middle href='mailto:<%=trim(rs("email"))%>'><%=trim(rs("email"))%></a></td>
                  </tr>
                  <tr>
                    <td height="24" bgcolor="#FCFDEB"><div align="right" class="top1">联系电话：</div></td>
                    <td bgcolor="#FCFDEB"><%=trim(rs("TelePhone"))%></td>
                  </tr>
                  <tr>
                    <td height="24" bgcolor="#FCFDEB"><div align="right" class="top1">联系地址：</div></td>
                    <td bgcolor="#FCFDEB"><%=trim(rs("Address"))%></td>
                  </tr>
                  <tr>
                    <td height="24" bgcolor="#FCFDEB"><div align="right" class="top1">类&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;型：</div></td>
                    <td bgcolor="#FCFDEB"><%if rs("ClassID")=0 then%>
                        <font color=red>投诉</font>
                        <%else%>
                        <font color=red>举报</font>
                        <%end if%></td>
                  </tr>
                  <tr>
                    <td height="24" bgcolor="#FCFDEB"><div align="right" class="top1">投诉标题：</div></td>
                    <td bgcolor="#FCFDEB"><%=trim(rs("Topic"))%></td>
                  </tr>
                  <tr>
                    <td height="59" bgcolor="#FCFDEB"><div align="right" class="top1">投诉内容：</div></td>
                    <td bgcolor="#FCFDEB"><%if rs("Display")=0 then%>
					                      <font color=red>*来信者不希望公开投诉举报内容*</font>
					                      <%else%>
					                      <%=trim(rs("content"))%>
										  <%end if%></td>
                  </tr>
                  <tr>
                    <td height="27" bgcolor="#FCFDEB"><div align="right" class="top1">回复时间：</div></td>
                    <td bgcolor="#FCFDEB">
					<%if trim(rs("ReplyContent"))<>"" then %>
					<%=trim(rs("ReplyTime"))%>
					<%else%>
					<font color=red >请耐心等待回复！</font>
					<%end if%>
					</td>
                  </tr>
                  <tr>
                    <td height="55" bgcolor="#FCFDEB"><div align="right" class="top1">回复内容：</div></td>
                    <td bgcolor="#FCFDEB">
					<%if trim(rs("ReplyContent"))<>"" then %>
					<%=trim(rs("ReplyContent"))%>
					<%else%>
					<font color=red >请耐心等待回复！</font>
					<%end if%>
					</td>
                  </tr>
                </table>
            </div></td>
          </tr>
          <tr>
            <td height="13">&nbsp;</td>
          </tr>
    </table></td>
  </tr>
</table>
<%
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		%>
<!--#include file="other_bottom.asp"-->
