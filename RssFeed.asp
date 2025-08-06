<%@ Language=VBScript%>
<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="function.asp"-->
<!--#include file="char.inc"-->
<%
dim iType,iBigClass,iSmallClass,iTopNews,iContentLength,sShowEmail,iCreatXML,sTitle,fdata,sAuthorEamil,sSaveFileName
iType=CheckStr(trim(Request.QueryString("E_typeid")))
iBigClass=CheckStr(trim(Request.QueryString("E_BigClassID")))
iSmallClass=CheckStr(trim(Request.QueryString("E_SmallClassID")))
iTopNews=CheckStr(trim(Request.QueryString("TopNews")))
iContentLength=CheckStr(trim(Request.QueryString("ContentLength")))
sShowEmail=trim(Request.QueryString("Email"))
iCreatXML=CheckStr(trim(Request.QueryString("iCreatXML")))
sTitle=""
if iContentLength<1 or iContentLength="" then
	iContentLength=100
end if

if iTopNews<0 or iTopNews="" then
	iTopNews=20
end if

if sShowEmail="" then
	sShowEmail="hide"
end if

sql="select top "& iTopNews &" * from "& db_EC_News_Table &" where checkked=1 order by newsid DESC"
if iType<>"" then
	sql="select top "& iTopNews &" * from "& db_EC_News_Table &" where checkked=1 and E_typeid="& iType &" order by newsid DESC"

	set rst=server.CreateObject("ADODB.RecordSet")
	sqlt="select * from "& db_EC_Type_Table &" where E_typeid="&iType
	rst.Open sqlt,conn,1,1
	if not rst.eof then
		sTitle="_"& rst("E_typename")
	end if
end if
if iBigClass<>"" then
	sql="select top "& iTopNews &" * from "& db_EC_News_Table &" where checkked=1 and E_BigClassID="& iBigClass &" order by newsid DESC"

	set rst=server.CreateObject("ADODB.RecordSet")
	sqlt="select * from "& db_EC_BigClass_Table &" where E_BigClassID="&iBigClass
	rst.Open sqlt,conn,1,1
	if not rst.eof then
		sTitle="_"& rst("E_BigClassName")
	end if
end if
if iSmallClass<>"" then
	sql="select top "& iTopNews &" * from "& db_EC_News_Table &" where checkked=1 and E_SmallClassID="& iSmallClass &" order by newsid DESC"

	set rst=server.CreateObject("ADODB.RecordSet")
	sqlt="select * from "& db_EC_SmallClass_Table &" where E_SmallClassID="&iSmallClass
	rst.Open sqlt,conn,1,1
	if not rst.eof then
		sTitle="_"& rst("E_smallclassname")
	end if
	rst.close
	set rst=nothing
end if
set sqlt=nothing

Response.ContentType = "text/XML"
'在网址后加(/)
if right(xpurl,1)<>"/" then
	xpurl=xpurl &"/"
end if
if left(xpurl,7)<>"http://" then
	xpurl="http://"& xpurl
end if

fdata="<?xml version=""1.0"" encoding=""gb2312"" ?>"&chr(10)
fdata=fdata & "<rss version=""2.0"">"&chr(10)
fdata=fdata & "<channel>"&chr(10)
fdata=fdata & "	<title>"& eChSys & sTitle &"</title>"&chr(10)
fdata=fdata & "	<description>"& eChSys &"</description>"&chr(10)
fdata=fdata & "	<link>"& xpurl &"</link>"&chr(10)
fdata=fdata & "	<language>zh-cn</language>"&chr(10)
fdata=fdata & "	<generator>Forecast News V1.1</generator>"&chr(10)
fdata=fdata & "	<managingEditor>"& email &"</managingEditor>"&chr(10)
fdata=fdata & "	<webMaster>"& email &"</webMaster>"&chr(10)
fdata=fdata & "	<image>"&chr(10)
fdata=fdata & "	<link>"& xpurl &"</link>" &chr(10)
fdata=fdata & "	<url>"& xpurl & logo &"</url>" &chr(10)
fdata=fdata & "	<title>"& eChSys &"</title>" &chr(10)
fdata=fdata & "	</image>"&chr(10)

set rs=server.CreateObject("ADODB.RecordSet")
rs.Open sql,conn,1,1
if not rs.eof then
	while not rs.EOF
		fdata=fdata & "	<item>"&chr(10)
		if rs("picnews")=1 then
			fdata=fdata & "	<title>[图]  "& rs("title") &"</title>"&chr(10)
		else
			fdata=fdata & "	<title>"& rs("title") &"</title>"&chr(10)
		end if
		fdata=fdata & "	<link>"& xpurl &"E_ReadNews.asp?NewsID="& rs("NewsID") &"</link>"&chr(10)
		if sShowEmail="hide" then		'不在作者栏显示EMAIL
			fdata=fdata & "	<author>(" & rs("editor") &")</author>" &chr(10)
		else
			set rsuser=server.CreateObject("ADODB.RecordSet")
			sqluser="select * from "& db_User_Table &" where "& db_User_Name &" like '"& trim(rs("editor")) &"'"
			rsuser.Open sqluser,ConnUser,1,1
			if not rsuser.eof then
				if rsuser(db_User_Email)<>"" then
					sAuthorEamil= rsuser(db_User_Email)
				else
					sAuthorEamil= "guest@e-cweb.com"
				end if
			end if
			rsuser.close
			set rsuser=nothing
			set sqluser=nothing
			fdata=fdata & "	<author>(" & rs("editor") &") " & sAuthorEamil & "</author>" &chr(10)
		end if
		fdata=fdata & "	<pubDate>" & rs("updatetime") & "</pubDate>" &chr(10)
		fdata=fdata & "	<description><![CDATA[" & CutStr(nohtmlcode(rs("Content")),iContentLength) & "<font color=red>『<a href='"& xpurl &"E_ReadNews.asp?NewsID="& rs("NewsID") &"'>要浏览详细内容，请点击标题</a>』</font>]]></description>" &chr(10)
		fdata=fdata & "	</item>"&chr(10)&chr(10)
		rs.MoveNext
	wend
else
		fdata=fdata & "	<item>"&chr(10)
		fdata=fdata & "	<title>无内容</title>"&chr(10)
		fdata=fdata & "	<link>"& xpurl &"</link>"&chr(10)
		fdata=fdata & "	<author>无内容</author>" &chr(10)
		fdata=fdata & "	<pubDate>" & now() & "</pubDate>" &chr(10)
		fdata=fdata & "	</item>"&chr(10)&chr(10)
end if
rs.close
conn.close
set rs=nothing
set conn=nothing
set sql=nothing
fdata=fdata & "</channel>"&chr(10)
fdata=fdata & "</rss>"&chr(10)
response.write fdata


if iCreatXML=1 then
	sSaveFileName="index.xml"
	if iType<>"" then
		sSaveFileName="index_type"& iType &".xml"
	end if
	if iBigClass<>"" then
		sSaveFileName="index_bigclass"& iBigClass &".xml"
	end if
	if iSmallClass<>"" then
		sSaveFileName="index_smallclass"& iSmallClass &".xml"
	end if
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set objCountFile=objFSO.CreateTextFile(Server.MapPath(sSaveFileName),True)
	objCountFile.Write fdata
	objCountFile.Close
	Set objCountFile=Nothing
	Set objFSO = Nothing
end if
%>