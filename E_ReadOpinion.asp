<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="function.asp"-->
<%
OpinionID=Request.QueryString("OpinionID")%>
<%
    OpinionID=ChkRequest(request("OpinionID"),1)	'��ע��
	if OpinionID="" then
	Show_Err("��������ҪҪ�鿴�ķ���ID!<br><br>������<a href='javascript:history.back()'>��ȥ����</a>������")
	response.end
	end if	
	
	set rs=server.CreateObject("ADODB.RecordSet")
	rs.Source="select * from "& db_EC_Opinion_Table &" where  OpinionID="&OpinionID	
	rs.Open rs.Source,conn,1,1	
		
	if rs.EOF then
	Show_Err("�Ƿ�����ID����ȷ��OpinionID�Ƿ����!<br><br>��<a href='javascript:history.back()'>��ȥ����</a>��")
	Response.End
	end if
		%>
<html>
<head>
<title>�Ķ�����__<%=eChSys%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="news.css" rel=stylesheet type=text/css>
</head>
<body bgcolor="#FFFFFF" text="#000000" topmargin="0">
<!--#include file="other_Top.asp"-->
<table width="780" height="450" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td width="780" valign="top"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="9%" height="53" valign="middle" background="Images/location_bg.gif" align="center" >&nbsp;<img src="Images/point_location.gif" width="24" height="24"></td>
        <td width="91%" valign="middle" background="Images/location_bg.gif" class="daohang">&nbsp;����λ�ã�&nbsp;<a class=daohang href="./" >��վ��ҳ</a>&gt;&gt;<a href="E_Opinion.asp" class="daohang">�����ײ�</a>&gt;&gt;�鿴����</td>
      </tr>
    </table>
        <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#D3E3ED">
          <tr class="TDtop">
            <td height="27" colspan="2" bgcolor="#FFFFFF" class="TDtop"><div align="center" class="top1">�鿴��������</div></td>
          </tr>
          <tr>
            <td width="14%" height="25" bgcolor="#FFFFFF"><div align="right" class="top1">�������ڣ�</div></td>
            <td width="86%" bgcolor="#FFFFFF"><%=year(rs("UpdateTime")) %>-<%=Month(rs("UpdateTime")) %>-<%=Day(rs("UpdateTime")) %></td>
          </tr>
          <tr>
            <td height="23" bgcolor="#FFFFFF"><div align="right" class="top1">����������</div></td>
            <td bgcolor="#FFFFFF"><%=trim(rs("YourName"))%></td>
          </tr>
          <tr>
            <td height="25" bgcolor="#FFFFFF"><div align="right" class="top1">��ϵ���䣺</div></td>
            <td bgcolor="#FFFFFF"><a class=middle href='mailto:<%=trim(rs("email"))%>'><%=trim(rs("email"))%></a></td>
          </tr>
          <tr>
            <td height="24" bgcolor="#FFFFFF"><div align="right" class="top1">��ϵ�绰��</div></td>
            <td bgcolor="#FFFFFF"><%=trim(rs("TelePhone"))%></td>
          </tr>
          <tr>
            <td height="24" bgcolor="#FFFFFF"><div align="right" class="top1">��ϵ��ַ��</div></td>
            <td bgcolor="#FFFFFF"><%=trim(rs("Address"))%></td>
          </tr>
          <tr>
            <td height="24" bgcolor="#FFFFFF"><div align="right" class="top1">�������ͣ�</div></td>
            <td bgcolor="#FFFFFF"><%if rs("ClassID")=0 then%>
                <font color=red>��ϣ��������ϣ���ظ������䣡</font>
                <%else%>
                <font color=red>���Թ���</font>
                <%end if%></td>
          </tr>
          <tr>
            <td height="24" bgcolor="#FFFFFF"><div align="right" class="top1">���԰�飺</div></td>
            <td  bgcolor="#FFFFFF"><%if rs("OpinionClass")=1 then%>
                <font color=red>���㼺��</font>
              <%end if%>
                <%if rs("OpinionClass")=2 then%>
              <font color=red>��չ����</font>
              <%end if%>
                <%if rs("OpinionClass")=3 then%>
              <font color=red>��������</font>
              <%end if%>
                <%if rs("OpinionClass")=4 then%>
              <font color=red>�ش��轨��</font>
              <%end if%>
            </td>
          </tr>
          <tr>
            <td height="24" bgcolor="#FFFFFF"><div align="right" class="top1">���Ա��⣺</div></td>
            <td bgcolor="#FFFFFF"><%=trim(rs("Topic"))%></td>
          </tr>
          <tr>
            <td height="59" bgcolor="#FFFFFF"><div align="right" class="top1">�������ݣ�</div></td>
            <td bgcolor="#FFFFFF"><%if rs("ClassID")=0 then%>
                <font color=red>*�����߲�ϣ��������������*</font>
                <%else%>
                <%=rs("content")%>
                <%end if%></td>
          </tr>
          <tr>
            <td height="27" bgcolor="#FFFFFF"><div align="right" class="top1">�ظ�ʱ�䣺</div></td>
            <td bgcolor="#FFFFFF"><%if trim(rs("ReplyContent"))<>"" then %>
                <%=trim(rs("ReplyTime"))%>
                <%else%>
                <font color=red >�����ĵȴ��ظ���</font>
                <%end if%>
            </td>
          </tr>
          <tr>
            <td height="55" bgcolor="#FFFFFF"><div align="right" class="top1">�ظ����ݣ�</div></td>
            <td bgcolor="#FFFFFF"><%if trim(rs("ReplyContent"))<>"" then %>
                <%=trim(rs("ReplyContent"))%>
                <%else%>
                <font color=red >�����ĵȴ��ظ���</font>
                <%end if%>
            </td>
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
