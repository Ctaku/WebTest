<!--#INCLUDE FILE="Conn.asp" -->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkURL.asp"-->
<%
IF request.cookies(eChsys)("KEY")="" THEN
	Show_Err("�Բ������ĺ�̨����Ȩ�޲���������")
	response.end
else
	if request("action")="update" then
		dim smallclassorder,smallclassma,E_smallclassview,E_smallclassname,E_BigClassID,smallclasszs,classurl
		for i=1 to request.form("E_SmallClassID").count
			E_SmallClassID=CheckStr(request.form("E_SmallClassID")(i))
			smallclassorder=CheckStr(request.form("smallclassorder")(i))
			smallclassma=CheckStr(request.form("smallclassma")(i))
			E_smallclassview=CheckStr(request.form("E_smallclassview")(i))
			E_smallclassname=CheckStr(request.form("E_smallclassname")(i))
			E_BigClassID=CheckStr(request.form("E_BigClassID")(i))
			smallclasszs=CheckStr(request.form("smallclasszs")(i))
			classurl=CheckStr(request.form("url")(i))
			if CheckStr(request.form("smallclassorder")(i))="" then
				Show_Err("����дС������<br><br><a href=history.go(-1)>����</a>")
				response.end
			end if
			if CheckStr(request.form("smallclassma")(i))="" then
				smallclassma="��"
			else
				master=split(CheckStr(request.form("smallclassma")(i)),"|")
			 	for k=0 to ubound(master)
					set rs=server.createobject("adodb.recordset")
					sql="Select * from "& db_User_Table &" where oskey='smallmaster' and  "& db_User_Name &"='"&trim(master(k))&"'"
					rs.open sql,ConnUser,1,3
					if trim(master(k))<>"��" then
						if rs.eof then
							Show_Err("С�����Ա���ޡ�& trim(master(k)) &���û���������ѡ���С��Ĺ���Ա��<br><br><a href='javascript:history.back()'>����</a>")
							Response.End
						end if
					else
						smallclassma="��"
					end if
					rs.close
					set rs=nothing
				next
			end if
			conn.execute("update "& db_EC_SmallClass_Table &" set smallclassorder="&smallclassorder&",smallclassma='"&smallclassma&"',E_smallclassview="&E_smallclassview&",E_smallclassname='"&E_smallclassname&"',E_BigClassID="&E_BigClassID&",smallclasszs='"&smallclasszs&"',url='"&classurl&"' where E_SmallClassID="&E_SmallClassID)
			Set nrs = Server.CreateObject("ADODB.Recordset")
			sqln="select * from "& db_EC_News_Table &" where E_SmallClassID="&E_SmallClassID
			nrs.open sqln,conn,3,3
			while not nrs.EOF
				nrs("E_BigClassID")=E_BigClassID
				nrs.MoveNext
			wend
			nrs.close
			set nrs=nothing
		next
	end if

	if request("action")="add" then
		function changechr(str) 
			changechr=replace(replace(replace(replace(str,"<","&lt;"),">","&gt;"),chr(13),"<br>")," "," ") 
			changechr=replace(changechr,"'","&acute;")
			changechr=replace(changechr,mid(" "" ",2,1),"&quot;")
		end function
					
		smallclasszs=request.form("smallclasszs")
		E_smallclassname=changechr(trim(request("E_smallclassname")))
		smallclassorder=request.form("smallclassorder")
		smallclassma=request.form("smallclassma")
		E_smallclassview=request.form("E_smallclassview")
		E_BigClassID=request.form("E_BigClassID")
		E_typeid=request.form("E_typeid")
		classurl=request.Form("url")
		
		If E_smallclassname="" Then
			Show_Err("С�����Ʋ���Ϊ�գ���<a href=javascript:history.go(-1)>����������д</a>��")
			response.end
		end if
		If smallclassorder="" Then
			Show_Err("С��������Ϊ�գ���<a href=javascript:history.go(-1)>����������д</a>��")
			response.end
		end if
		
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql="select * from "& db_EC_SmallClass_Table &""
		rs.open sql,conn,3,3
	 	rs.addnew
		rs("E_smallclassname")=E_smallclassname
		rs("smallclasszs")=smallclasszs
		rs("E_BigClassID")=E_BigClassID
		rs("E_typeid")=E_typeid
		rs("url")=classurl
		if smallclassorder="" then
			rs("smallclassorder")=0
		else
			rs("smallclassorder")=smallclassorder
		end if
		rs("E_smallclassview")=E_smallclassview
		if smallclassma="" then
			rs("smallclassma")="��"
		else
			rs("smallclassma")=smallclassma
		end if
		rs.update
		E_BigClassID=rs("E_BigClassID")
		rs.close
		set rs=nothing
	end if
		
	conn.close
	set conn=nothing
	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=E_SmallClassManage.asp?E_BigClassID="&E_BigClassID&""">"
	Show_Message("<p align=center><font color=red>��ϲ��!���óɹ�!<br><br>"&freetime&"���Ӻ󷵻���ҳ!</font>")
	response.end
end if
%>