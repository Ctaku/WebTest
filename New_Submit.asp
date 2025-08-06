<%@ Language=VBScript %>
<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp" -->
<%
if not((request.cookies(eChsys)("ManageKEY")="super" ) or (request.cookies(eChsys)("KEY")="super" )) then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	Set rs9 = Server.CreateObject("ADODB.Recordset")
	sql9 ="SELECT * From "& db_EC_System_Table &" Order By id DESC"
	RS9.open sql9,Conn,1,1

	if rs9("init")="1" or (request.cookies(eChsys)("ManagePurview")="99999" and request.cookies(eChsys)("purview")="99999") then

		if (request.form("action")="del") then
			if request.form("selectdel")<>"" then
				For i = 1 To Request.Form("selectdel").Count
					SelectTable=Request.Form("selectdel")(i)
					select case SelectTable
						'初始化部门表
						case db_EC_Dep_Table conn.execute "delete from "& db_EC_Dep_Table
						'初始化总栏表
						case db_EC_Type_Table conn.execute "delete from "& db_EC_Type_Table
						'初始化大栏表
						case db_EC_BigClass_Table conn.execute "delete from "& db_EC_BigClass_Table
						'初始化小栏表
						case db_EC_SmallClass_Table conn.execute "delete from "& db_EC_SmallClass_Table
						'初始化专题表
						case db_EC_Special_Table conn.execute "delete from "& db_EC_Special_Table
						'初始化文章表
						case db_EC_News_Table conn.execute "delete from "& db_EC_News_Table
						'初始化临时表
						case db_NewsFile_Table conn.execute "delete from "& db_NewsFile_Table
						'初始化上传临时表
						case db_UploadPic_Table
								set rs2=server.createobject("adodb.recordset")
								sql2="select * from "& db_UploadPic_Table
								rs2.open sql2,conn,1,3
								do while not rs2.EOF
									Set fso=Server.CreateObject("Scripting.FileSystemObject")
									If fso.FileExists(Server.Mappath(FileUploadPath & rs2("picname")))=true Then
										fso.DeleteFile Server.MapPath(FileUploadPath & rs2("picname"))
									End If
									Set fso=Nothing
									rs2.movenext
								loop
								rs2.close
								set rs2=nothing
								conn.execute "delete from "& db_UploadPic_Table
						'初始化投票表
						case db_EC_Vote_Table conn.execute "delete from "& db_EC_Vote_Table
						'初始化友情链接表
						case db_EC_Link_Table conn.execute "delete from "& db_EC_Link_Table
						'初始化公告表
						case db_EC_Board_Table conn.execute "delete from "& db_EC_Board_Table
						'初始化附件表
						case db_EC_Attach_Table
								set rs1=server.createobject("adodb.recordset")
								sql1="select * from "& db_EC_Attach_Table
								rs1.open sql1,conn,1,3
								do while not rs1.EOF
									Set fso=Server.CreateObject("Scripting.FileSystemObject")
									If fso.FileExists(Server.Mappath(FileUploadPath & rs1("filename")))=true Then
										fso.DeleteFile Server.MapPath(FileUploadPath & rs1("filename") )
									End If
									Set fso=Nothing
									rs1.movenext
								loop
								rs1.close
								set rs1=nothing
						 		conn.execute "delete from "& db_EC_Attach_Table
						'初始化评论表
						case (db_EC_Review_Table&"1") conn.execute "delete from "& db_EC_Review_Table &" where NewsId>0"
						'初始化留言表
						case (db_EC_Review_Table&"0") conn.execute "delete from "& db_EC_Review_Table &" where NewsId<1"
						'初始化用户表
						case db_User_Table ConnUser.execute "delete from "& db_User_Table &" where purview<>99999"
						'初始化管理员表
						case db_EC_Manage_Table conn.execute "delete from "& db_EC_Manage_Table &" where purview<>99999"
					end select
				Next
				ConnUser.close
				set ConnUser=nothing
				conn.close
				set conn=nothing
				response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=New.asp"">"
				Show_Message("<p align=center><font color=red>系统表"& (request.form("selectdel")) &"初始化完成!<br><br>"&freetime&"秒钟后返回初始页!</font>")
				response.end
			else
				Show_Err("对不起，您并没有选择任何操作！<br><br>－－－<a href='javascript:history.back()'>回去重来</a>－－－")
				response.end
			end if
		else
			SpecialID=Request.Form("SpecialID")
			button_value=trim(Request.Form("alert_button"))
			if button_value="确定" then
				conn.execute("delete from "& db_EC_BigClass_Table &"")
				conn.execute("delete from "& db_EC_SmallClass_Table &"")
				conn.execute("delete from "& db_EC_Special_Table &"")
				conn.execute("delete from "& db_EC_News_Table &"")
				conn.execute("delete from "& db_EC_Review_Table &"")
				conn.execute("delete from "& db_EC_Vote_Table &"")
				conn.execute("delete from "& db_EC_Link_Table &"")
				conn.execute("delete from "& db_EC_Type_Table &"")
				conn.execute("delete from "& db_EC_Dep_Table &"")
				conn.execute("delete from "& db_EC_Board_Table &"")

				set rs1=server.createobject("adodb.recordset")
				sql1="select * from "& db_EC_Attach_Table
				rs1.open sql1,conn,1,3
				do while not rs1.EOF
					Set fso=Server.CreateObject("Scripting.FileSystemObject")
					If fso.FileExists(Server.Mappath(FileUploadPath & rs1("filename")))=true Then
						fso.DeleteFile Server.MapPath(FileUploadPath & rs1("filename") )
					End If
					Set fso=Nothing
					rs1.movenext
				loop
				rs1.close
				set rs1=nothing
				conn.execute("delete from "& db_EC_Attach_Table &"")
				conn.execute("delete from "& db_NewsFile_Table &"")

				set rs2=server.createobject("adodb.recordset")
				sql2="select * from "& db_UploadPic_Table
				rs2.open sql2,conn,1,3
				do while not rs2.EOF
					Set fso=Server.CreateObject("Scripting.FileSystemObject")
					If fso.FileExists(Server.Mappath(FileUploadPath & rs2("picname")))=true Then
						fso.DeleteFile Server.MapPath(FileUploadPath & rs2("picname"))
					End If
					Set fso=Nothing
					rs2.movenext
				loop
				rs2.close
				set rs2=nothing
				conn.execute("delete from "& db_UploadPic_Table &"")

				if UserTableType = "EChuang" then		'不整合则
					ConnUser.execute("delete from "& db_User_Table &" where purview<>99999")
				end if
				response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=admin_index.asp"">"
				Show_Message("<p align=center><font color=red>系统已经被初始化!<br><br>"&freetime&"秒钟后返回首页!</font>")
				response.end
			else
				Response.Redirect "admin_index.asp"
			end if
		end if
		rs9.close
		set rs9=nothing
		conn.close
		set conn=nothing
	else
		Show_Err("对不起，该功能已经被超级系统管理员关闭，您没有权限操作！")
		response.end
	end if
end if%>