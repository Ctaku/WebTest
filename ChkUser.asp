<!--#include file="SafeRequest.asp"-->
<!--#include file="char.inc"-->

<%
if ((Request.cookies(eChsys)("UserName")="") or (Request.cookies(eChsys)("passwd")="") or (Request.cookies(eChsys)("KEY")="")) then
	if IsDebug="1" then
		Show_Err("cookies(UserName)(passwd)(KEY)中有为空值。<br><br><a href='javascript:history.back()'>返回</a>")
	else
		response.redirect "login.asp"
	end if
	response.end
else
	if not(Request.cookies(eChsys)("KEY")="super" or Request.cookies(eChsys)("KEY")="check" or Request.cookies(eChsys)("KEY")="typemaster" or Request.cookies(eChsys)("KEY")="bigmaster" or Request.cookies(eChsys)("KEY")="smallmaster" or Request.cookies(eChsys)("KEY")="selfreg") THEN
		if IsDebug="1" then
				Show_Err("cookies(KEY)值不是[super][check][typemaster][bigmaster][smallmaster][selfreg]中的一个。<br><br><a href='javascript:history.back()'>返回</a>")
		else
			response.redirect "login.asp"
		end if
		response.end
	else
		UserName=CheckStr(trim(Request.cookies(eChsys)("username")))
		PassWD=trim(Request.cookies(eChsys)("passwd"))
		Key=trim(request.cookies(eChsys)("KEY"))
		
		set urs=server.createobject("adodb.recordset")
		sql="select * from "& db_User_Table &" where "& db_User_Name &"='"& UserName &"'"
		urs.open sql,ConnUser,1,3
		if urs.bof or urs.eof then
			if IsDebug="1" then
					Show_Err("以cookies(username)的返回值，查找不到用户表中的相对应记录。<br><br><a href='javascript:history.back()'>返回</a>")
			else
				response.redirect "login.asp"
			end if
			response.end
		else
			if urs("jingyong")=0 or (Request.cookies(eChsys)("KEY")="super") then
				if PassWD<>urs(db_User_Password) then
					if IsDebug="1" then
						Show_Err("cookies(passwd)的返回值与用户表中的相对应记录的Password值不符。<br><br><a href='javascript:history.back()'>返回</a>")
					else
						response.redirect "login.asp"
					end if
					response.end
				else
					if key<>urs("OSKEY") then
						if IsDebug="1" then
							Show_Err("cookies(KEY)的返回值与用户表中的相对应记录的OSKEY值不符。<br><br><a href='javascript:history.back()'>返回</a>")
						else
							response.redirect "login.asp"
						end if
						response.end
					end if
				end if
			else
				Show_Err("对不起，您的用户状态为不可用，请联系管理员解决！<br><br><a href='javascript:history.back()'>返回</a>")
				response.end
			end if
		end if
		urs.close
		set urs=nothing
	end if
end if
%>