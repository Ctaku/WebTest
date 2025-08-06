<%@ Language=VBScript %>
<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="function.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp"-->
<%
if not(request.cookies(eChsys)("ManageKEY")="super" or request.cookies(eChsys)("ManageKEY")="typemaster" or request.cookies(eChsys)("ManageKEY")="bigmaster") then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	dim jingyong,delnews
	delnews=request.form("delnews")
	if delnews="1" then
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_User_Table &" where  "& db_User_Name &"='"&request.cookies(eChsys)("username")&"'"
		rs.open sql,ConnUser,1,3
		jingyong=rs("jingyong")
		rs.close
		set rs=nothing

		if request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="check" or request.cookies(eChsys)("key")="bigmaster" or request.cookies(eChsys)("key")="smallmaster" or (request.cookies(eChsys)("key")="selfreg" and jingyong=0) or request.cookies(eChsys)("key")="typemaster" then%>
			<%dim username,E_typeid,E_typename,E_BigClassID,E_BigClassName,E_SmallClassID,E_smallclassname
			NewsID=ChkRequest(Request.Form("NewsID"),1)	'防注入
			set rs=server.CreateObject("ADODB.RecordSet")
			rs.Source="select * from "& db_EC_News_Table &" where NewsID=" & NewsID
			rs.Open rs.source,conn,1,1
			image=rs("image")
			username=rs("editor")

			dim title
			title=rs("title")
			newsrelated=rs("newsrelated")
			E_typeid=rs("E_typeid")
			E_BigClassID=rs("E_BigClassID")
			E_SmallClassID=rs("E_SmallClassID")
			rs.close
			set rs=nothing
			set rs=server.CreateObject("ADODB.RecordSet")
			rs.Source="select * from "& db_EC_Type_Table &" where E_typeid=" & E_typeid
			rs.Open rs.source,conn,1,1
			E_typename=rs("E_typename")
			rs.close
			set rs=Nothing
			if E_SmallClassID<>"" then
			set rs=server.CreateObject("ADODB.RecordSet")
			rs.Source="select * from "& db_EC_BigClass_Table &" where E_BigClassID=" & E_BigClassID
			rs.Open rs.source,conn,1,1
			E_BigClassName=rs("E_BigClassName")
			rs.close
			set rs=Nothing
			end if
			if E_SmallClassID<>"" then
				set rs=server.CreateObject("ADODB.RecordSet")
				rs.Source="select * from "& db_EC_SmallClass_Table &" where E_SmallClassID=" & E_SmallClassID
				rs.Open rs.source,conn,1,1
				E_smallclassname=rs("E_smallclassname")
				rs.close
				set rs=nothing
			end if
			button_value=trim(Request.Form("alert_button"))
			dim dep1,dep2
			if button_value="是" then
				if request.Form("del")<>"1" then
					set rs2=server.createobject("adodb.recordset")
					sql2="select * from "& db_User_Table &" where  "& db_User_Name &"='"&username&"'"
					rs2.open sql2,ConnUser,1,3
					dep1=rs2("E_DepName")
		            dep2=rs2("E_DepType")
					rs2("number")=rs2("number")-1
					rs2.update
					rs2.close
					set rs2=nothing

					set rs5=server.createobject("adodb.recordset")
					sql5="select * from "& db_EC_Dep_Table &" where  ( E_DepName='"&dep1&"'  and E_DepType='"&dep2&"' ) "
					rs5.open sql5,Conn,1,3
					If(not rs5.eof) or (not rs5.bof) Then
						rs5("depnumber")=rs5("depnumber")-1
						rs5.update
					End if
					rs5.close
					set rs5=nothing
				end if
				conn.execute("delete from "& db_EC_News_Table &" where NewsID=" & NewsID)
				conn.execute("delete from "& db_EC_Review_Table &" where NewsID=" & NewsID)
				set rs2=server.createobject("adodb.recordset")
				sql2="select * from "& db_EC_Attach_Table &" where NewsID=" & NewsID
				rs2.open sql2,conn,1,3
				do while not rs2.EOF
					Set fso=Server.CreateObject("Scripting.FileSystemObject")
					If fso.FileExists(Server.Mappath(FileUploadPath & rs2("filename")))=true Then
						fso.DeleteFile Server.MapPath(FileUploadPath & rs2("filename") )
					End If
					Set fso=Nothing
					rs2.movenext
				loop
				rs2.close
				set rs2=nothing
				conn.execute("delete from "& db_EC_Attach_Table &" where NewsID=" & NewsID)
				if E_SmallClassID<>"" then
					response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=listnews.asp?E_SmallClassID="&E_SmallClassID&""">"
					Show_Message("<p align=center><font color='000000'>总栏“"&E_typename&"”下大类“"&E_BigClassName&"”下的小类“"&E_smallclassname&"”的文章“"&title&"”已经被删除！<br><br>"&freetime&"秒钟后返回上页！</font>")
					response.end
				else if E_BigClassID<>"" then
					response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=E_SmallNO.asp?E_BigClassID="&E_BigClassID&""">"
					Show_Message("<p align=center><font color='000000'>总栏“"&E_typename&"”下大类“"&E_BigClassName&"”的文章“"&title&"”已经被删除！<br><br>"&freetime&"秒钟后返回上页！</font>")
					response.end
					else
						response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=E_bigclassNO.asp?E_typeid="&E_typeid&""">"
						Show_Message("<p align=center><font color='000000'>总栏“"&E_typename&"”下的文章“"&title&"”已经被删除！<br><br>"&freetime&"秒钟后返回上页！</font>")
						response.end
					end if			
				end if
			else
				if E_SmallClassID<>"" then
					response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=listnews.asp?E_SmallClassID="&E_SmallClassID&""">"
				else 
					if E_BigClassID<>"" then
					response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=E_SmallNO.asp?E_BigClassID="&E_BigClassID&""">"
					else 
						response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=E_bigclassNO.asp?E_typeid="&E_typeid&""">"
					end if
				end if
				Show_Err("您没有执行删除操作！<br><br>"&freetime&"秒钟后返回上页！</a>")
				response.end
			end if
		end if
	else
		Show_Err("您无权删除此项目！<br><br><a href='javascript:history.back()'>返回</a>")
		response.end
	end if
end if%>