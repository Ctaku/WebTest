<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkURL.asp"-->

<%IF request.cookies(eChsys)("KEY")<>"super" THEN
	Show_Err("�Բ������ĺ�̨����Ȩ�޲���������")
	response.end
else
	if ggmanage="1" or (request.cookies(eChsys)("ManagePurview")="99999" and request.cookies(eChsys)("purview")="99999") then

	inuse=request.form("inuse")
	if inuse="" then
		Show_Err("�Բ�������ѡ�񹫸��<br><br><a href=history.go(-1)>����</a>")
		response.end
	end if
	set rs=server.createobject("adodb.recordset")
	sql="Select inuse from "& db_EC_Board_Table &" where inuse=1"
	rs.open sql,conn,1,3
	if not rs.eof then
		do while not rs.eof
			rs("inuse")=0
		rs.movenext
		loop
	end if
	rs.close
	
	sql="Select * from "& db_EC_Board_Table &" where ID="&inuse
	rs.open sql,conn,1,3
	if not rs.EOF then
		do while not rs.EOF 
			rs("inuse")=1
			rs.MoveNext 
		loop
	end if
	rs.close
	set rs=nothing
	dim title
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_EC_Board_Table &" where inuse=1"
	rs.open sql,conn,1,1
	title=rs("title")
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.redirect "E_BoardManage.asp"
	else
		Show_Err("�Բ��𣬸ù����Ѿ�������ϵͳ����Ա�رգ���û��Ȩ�޲�����")
		response.end
	end if
end if%>