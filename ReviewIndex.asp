<!--#include file="conn.asp"-->
<!--#include file="function.asp"-->
<!--'#include file="char.inc"-->
<%
dim shu,order,path,ss,oo,tt
shu=request.querystring("shu")
order=request("order")
title=request.querystring("title")'
Path = "./"
if shu<>"" then   '��ʾ��Ŀ������Ϊ6
	ss=shu
else
	ss=6
end if
if order="reviewid" then  '����ʽ��clickΪ�������
	oo="UpdateTime"
else
	oo="reviewid"
end if

if title<>"" then   '��ʾ������Ŀ������Ϊ10
	nn=title
else
	nn=50
end if %>
<table width="98%">
<%set rs=conn.execute("select top "&ss&" * from "& db_EC_Review_Table &" where passed=1 order by "&oo&" desc")
if rs.eof and rs.bof then
response.Write("<tr><td align=center ></td></tr>")
else
do while not rs.eof
title=nohtml(CheckStr(trim(rs("Content"))))'��ȡ����������ֶ����ݲ��滻��ʽ��
if rs("NewsID")>0 and rs("passed")=1 then%>
<tr><td align="left">		
			����<A class=blacklink href="E_ReadNews.asp?NewsID=<%=rs("NewsID")%>" title="<%=CutStr(title,34)%>" target="_blank"><%=CutStr(title,nn)%></A><%if tt=1 then%><%end if%><%if date()-rs("updatetime")<2 then%><FONT color="#ff0000">��!</FONT><%end if%></td></tr>
<tr>
<td height="1" colspan="2" background="IMAGES/bg_line.gif"></td>
</tr>
			
<%else%>
		
<%if rs("NewsID")<0 and rs("passed")=1 then%>
<tr><td align="left"> 			
            �ԡ�<A class=blacklink href="E_GuestBook.asp?ReViewID=<%=rs("ReViewID")%>" title"������Ա�����Ļ����ǹ���Ա���ܲ鿴!" target="_blank"><font color=green>���Ļ�!</font></A>
			<%if date()-rs("updatetime")<2 then%><FONT color="#ff0000">��!</FONT><%end if%>
</td></tr>
                <tr>
                  <td height="1" colspan="2" background="IMAGES/bg_line.gif"></td>
                </tr>
			<%else%>
<tr><td align="left">				
            �ԡ�<A class=blacklink href="E_GuestBook.asp?ReViewID=<%=rs("ReViewID")%>" title="<%=CutStr(title,34)%>" target="_blank"><%=CutStr(title,nn)%></A><%if date()-rs("updatetime")<2 then%><FONT color="#ff0000">��!</FONT><%end if%>
</td></tr>
                <tr>
                  <td height="1" colspan="2" background="IMAGES/bg_line.gif"></td>
                </tr>
<%end if%>
<%end if
rs.movenext 
loop 
end if
Rs.Close
set Rs=nothing
%>
</table>