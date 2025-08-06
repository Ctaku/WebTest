<%@ Language=VBScript%>
<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>网站地图__<%=eChSys%></title>
<LINK href=news.css rel=stylesheet>
</head>

<body>
<!--#include file="other_Top.asp"-->
<table width="780" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr valign="top">
    <td rowspan="2">
	<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="F1F7FC">
      <tr>
        <td height="25" background="Images/dh_bg.gif" class="daohang">&nbsp;&nbsp;&nbsp;<img src="Images/arrow_dh_1.gif" width="7" height="7">&nbsp;&nbsp;<span class="daohang">&nbsp;栏目导航&nbsp;&nbsp;当前位置：<a class="daohang" href="./" >网站首页</a>&gt;&gt;网站地图</span></td>
      </tr>   
      <tr>
        <td height="25">
          <table width="99%" height="126" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" id="AutoNumber3" style="border-collapse: collapse">
            <tr>
              <td height="10" valign=top>&nbsp;</td>
            </tr>
            <tr>
              <td height="113" valign=top><table width="100%" border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td width="20%" height="20" colspan="5">
                      <%
dim menuid1
dim menuname1
dim menucontent1
dim menuview1
set rs22=server.CreateObject("ADODB.RecordSet")
rs22.Source="select * from " & db_EC_Type_Table & " where E_typeview=1 order by E_typeorder"
rs22.Open rs22.Source,conn,1,1
i=1

while not rs22.EOF
RecordCount=rs22.RecordCount
menuid1=rs22("E_typeid")
menuname1=rs22("E_typename")
menucontent1=rs22("typecontent")
menuview1=rs22("E_typeview")

%>
                      <img src="images/m.gif" width="13" height="11"><a class=middle target="_top" href='E_Type.asp?E_typeid=<%=menuid1%> ' title=<%=menucontent1%> ><%=menuName1 %></a><br>
                      <%  
E_typeid=rs22("E_typeid")
	set rs21=conn.execute("SELECT * FROM " & db_EC_BigClass_Table & " where E_typeid="&E_typeid&" order by E_bigclassorder") 
	do while not rs21.eof
%>
                      <%if not Rs21.eof then%>
                      <table width="100%" border="0" cellpadding="3" cellspacing="0">
                        <tr>
                          <td width="18%">&nbsp;&nbsp;<img src="images/m1.gif" width="13" height="11"><a target="_top" class=middle href="E_BigClass.asp?E_typeid=<%=E_typeid%>&E_BigClassID=<%=Rs21("E_BigClassID")%>"><%=Rs21("E_BigClassName")%></a></td>
                          <td width="82%" background="images/h.gif">
                            <%set nrs=conn.execute("SELECT * FROM "& db_EC_SmallClass_Table &" where E_BigClassID="&cstr(rs21("E_BigClassID"))&" order by smallclassorder")%>
                            <%do while not nrs.eof%>
                            <%if not nRs.eof then%>
                            <a target="_top" class=class href="E_SmallClass.asp?E_typeid=<%=E_typeid%>&E_BigClassID=<%=nrs("E_BigClassID")%>&E_SmallClassID=<%=nrs("E_SmallClassID")%>"><%=nRs("E_smallclassname")%></a>
                            <%nrs.movenext   
		     end if %>
                            <%if not nRs.eof then%>
                            <a target="_top" class=class href="E_SmallClass.asp?E_typeid=<%=E_typeid%>&E_BigClassID=<%=nrs("E_BigClassID")%>&E_SmallClassID=<%=nrs("E_SmallClassID")%>"><%=nRs("E_smallclassname")%></a>
                            <%nRs.movenext
             end if%>
                            <%loop
            nRs.Close
           set nRs=nothing
            %>                          </td>
                        </tr>
                      </table>
                      <%rs21.movenext   
	end if
	loop  
	rs21.close
%>
                      <br>
                      <%


i=i+1

rs22.MoveNext
wend
rs22.close
set rs22=nothing
%>
                    </td>
                  </tr>
                  <tr>
                    <td height="10" valign=top>&nbsp;</td>
                  </tr>
              </table></td>
            </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
 
</table>
<!--#include file="other_bottom.asp" -->
</body>
</html>
