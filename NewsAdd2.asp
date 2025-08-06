<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp"-->
<!--#include file="inc/config2.asp"-->
<%
Server.ScriptTimeout=999
dim SavePath,SaveSecondPath,FileExtPath

SaveSecondPath=year(now)&"-"&month(now)
FileExtPath="gethttppic"				'偷图存放的二级分类目录
SavePath = FileUploadPath   '存放上传文件的目录
if right(SavePath,1)<>"/" then SavePath=SavePath&"/" '在目录后加(/)

Dim strData
Dim intFieldCount
Dim i

content = Request.Form("content")

'删除被添加的本站WEB路径,即将引用本站资源的绝对改为相对路径
if Request.ServerVariables("SERVER_PORT")="80" then
	weburl="http://"& Cstr(Request.ServerVariables("SERVER_NAME")) & Cstr(Request.ServerVariables("url"))
else
	weburl="http://"& Cstr(Request.ServerVariables("SERVER_NAME")) &":"& Request.ServerVariables("SERVER_PORT") & Cstr(Request.ServerVariables("url"))
end if

'注意,下一行中被替换字符串应为实际的本文件名
weburl=replace(weburl,"NewsAdd2.asp","",1,-1,1)
content=replace(content,weburl,"",1,-1,1)
'替换系统设定的weburl
content=replace(content,"http://"&xpurl&"/","",1,-1,1)

PicUrl=Request.Form("PicUrl")

Set objRegExp = New Regexp
objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.Pattern = "<img.+?>"
strs=trim(content)

'是否偷图
if request.Form("getphoto")="1" then
	Set Matches =objRegExp.Execute(strs)
	For Each Match in Matches
		RetStr = RetStr &getimgs( Match.Value )
	Next 
end if 

function getimgs(str)
	Dim strChkExtFileName, ii

	getimgs=""
	Set objRegExp1 = New Regexp
	objRegExp1.IgnoreCase = True
	objRegExp1.Global = True
	objRegExp1.Pattern = "http://.+?"""
	set mm=objRegExp1.Execute(str)
	For Each Match1 in mm
		strChkExtFileName = Trim(left(Match1.Value,len(Match1.Value)-1))
		Forumupload = Split(UpFileType,"|")		'UpFileType为inc/config2.asp文件中相关值

		For ii=0 To ubound(Forumupload)
			If GetHttpfileExt(strChkExtFileName) = Trim(Forumupload(ii)) Then
				getimgs = getimgs & "||" & strChkExtFileName	'文件名累加
				Exit For
			End if
		Next
	next
end function

function GetHttpfileExt(FullPath)
	If FullPath <> "" Then
		GetHttpfileExt = trim(mid(FullPath, InStrRev(FullPath, ".")+1))
	Else
		GetHttpfileExt = ""
	End If
End function

function getHTTPPage(url)
	on error resume next
	dim http

	ServePath=server.mappath(SavePath)
	Set fso=server.createobject("scripting.filesystemobject")

	if fso.FolderExists(ServePath) then
		'检查Config.asp中设置的上传目录,无则自动建立
	else
    Set f = fso.CreateFolder(ServePath)
    set f=nothing
	End if
	
	if fso.FolderExists(ServePath &"\"& FileExtPath) then
		'检查上传目录有没有上传文件类型(扩展名)目录,无则自动建立
	else
    Set f = fso.CreateFolder(ServePath &"\"& FileExtPath)
    set f=nothing
	End if

	if fso.FolderExists(ServePath &"\"& FileExtPath &"\"& SaveSecondPath) then
		'检查上传目录有没有本年月目录,无则自动建立
	else
    Set f = fso.CreateFolder(ServePath &"\"& FileExtPath &"\"& SaveSecondPath)
    set f=nothing
	End if
	set fso=nothing

	set http=server.createobject("MSXML2.XMLHTTP")
	Http.open "GET",url,false
	Http.send()
	if Http.readystate<>4 then 
		exit function
	end if
	getHTTPPage=Http.responseBody
	set http=nothing
	if err.number<>0 then err.Clear 
end function

arrimg=split(retstr,"||")
allimg=""
newimg=""
for i=1 to ubound(arrimg)
	if arrimg(i)<>"" and instr(allimg,arrimg(i))<1 then
		randomize
		ranNum=int(900*rnd)+100
		SaveFileName = year(now) & month(now) & day(now) & hour(now) & minute(now) & second(now) & ranNum & cstr(i & mid(arrimg(i), instrrev(arrimg(i), ".")))

		fname=SavePath & FileExtPath &"/"& SaveSecondPath &"/"& SaveFileName

		dim geturl,objStream,imgs
		geturl=trim(arrimg(i))
		imgs=gethttppage(geturl)
		Set objStream = Server.CreateObject("ADODB.Stream")
		objStream.Type =1
		objStream.Open
		objstream.write imgs
		objstream.SaveToFile server.mappath(fname),2
		objstream.Close()
		set objstream=nothing
		allimg=allimg&"||"&arrimg(i)
		newimg=newimg&"||"&fname
	end if
next
arrnew=split(newimg,"||")
arrall=split(allimg,"||")
for i=1 to ubound(arrnew)
	arrnew(i)=replace(arrnew(i),""& FileUploadPath,FileUploadPath)
	strs=replace(strs,arrall(i),arrnew(i))
	arrnew(i)=replace(arrnew(i),FileUploadPath,"")
next
content=strs
if left(Picurl,4)="http" then
	
	strChkExtFileName = Trim(Picurl)
	Forumupload = Split(UpFileType,"|")		'UpFileType为inc/config2.asp文件中相关值

	For ii=0 To ubound(Forumupload)
		If GetHttpfileExt(strChkExtFileName) = Trim(Forumupload(ii)) Then	'由Picurl传过来的文件类型检测

			randomize
			ranNum=int(900*rnd)+100
			SaveFileName=year(now) & month(now) & day(now) & hour(now) & minute(now) & second(now) & ranNum & cstr(i & mid(Picurl, instrrev(Picurl, ".")))
			fname=SavePath & FileExtPath &"/"& SaveSecondPath &"/"& SaveFileName

			dim imgsa
			imgsa=gethttppage(Picurl)
			Set objStream = Server.CreateObject("ADODB.Stream")
			objStream.Type =1
			objStream.Open
			objstream.write imgsa
			objstream.SaveToFile server.mappath(fname),2
			objstream.Close()
			set objstream=nothing
			aassss=Picurl
			Picurl=fname
			Picurl=replace(Picurl,FileUploadPath,"")
			aassss1=Picurl
			content=replace(content,aassss,FileUploadPath & Picurl)

			Exit For	'退出FOR循环检索扩展名
		End if
	Next
end if
if Content="" then
	Show_Err("请输入文章内容！<br><br><a href='javascript:history.back()'>返回</a>")
	response.end
end if
if request.cookies(eChsys)("key")="super" then
	if request("viewhtml")<>"" then 
		Show_Err("请取消查看HTML源代码后再添加！<br><br><a href='javascript:history.back()'>返回</a>")
		response.end
	end if
end if
E_SmallClassID=Request.Form("E_SmallClassID")
E_BigClassID=Request.Form("E_BigClassID")
E_typeid=Request.Form("E_typeid")
title=CheckStr(trim(request.form("title")))


if title="" then
	Show_Err("请填写文章标题！<br><br><a href='javascript:history.back()'>返回</a>")
	response.end
end if

if len(title)>100 then
	Show_Err("文章标题过长！<br><br><a href='javascript:history.back()'>返回</a>")
	response.end
end if

if request.Form("goodnews")="1" then
	goodnews=1
else
	goodnews=0
end if

if request.Form("daodu")="1" then
	daodu=1
else
	daodu=0
end if

if request.Form("istop")="1" then
	istop=1
else
	istop=0
end if

if request.Form("backtype")="1" then
	backtype=1
else
	backtype=0
end if

if request.Form("backtype1")="1" then
	backtype1=1
else
	backtype1=0
end if

SpecialID=Request.Form("SpecialID")
SpecialID2=Request.Form("SpecialID2")
Author=CheckStr(trim(Request.Form("Author")))
Original=CheckStr(trim(Request.Form("Original")))
about=CheckStr(trim(Request.Form("about")))

if request.Form("PicUrl")="" then
	picnews=0
else
	picnews=1
end if

Picurl=replace(Picurl,FileUploadPath,"")
editor=request.form("editor")

'判断除小类以外管理员权限
if request.cookies(eChsys)("key")="super" or request.cookies(eChsys)("key")="check" or request.cookies(eChsys)("key")="bigmaster" or request.cookies(eChsys)("key")="typemaster" then
	checkked=1
Else
	'判断注册用户总权限
	if request.cookies(eChsys)("key")="selfreg" and fabiaocheck="1" then
		checkked=1	
	else
		'判断小类管理员权限
		if request.cookies(eChsys)("key")="smallmaster" And usecheck="1" then
			checkked=1
		Else
			checkked=0
			if request.cookies(eChsys)("shenhe")="0" then
				checkked=1
			else
				checkked=0
			end If
		end if	
	end if
end if

EnCode=Request.Form("EnCode")
newsrelated=Request.Form("newsrelated")
newslevel=Request.Form("newslevel")
title_color=Request.Form("title_color")
title_type=Request.Form("title_type")
title_face=Request.Form("title_face")
title_size=Request.Form("title_size")

set rs=server.createobject("adodb.recordset")
sql="select * from "& db_EC_News_Table
rs.open sql,conn,1,3
rs.addnew
rs("title")=title
rs("Author")=Author

''修改上传文件路径标记为[="uploadfile/]
content=replace(content,"="&chr(34)&FileUploadPath,"="&chr(34)&"uploadfile/",1,-1,1)
rs("content")=content
rs("Original")=Original
rs("goodnews")=goodnews
rs("daodu")=daodu
rs("picnews")=picnews
if newslevel="" then
	rs("newslevel")=0
else
	rs("newslevel")=newslevel
end if
rs("istop")=istop
rs("backtype")=backtype
rs("backtype1")=backtype1
rs("editor")=editor
rs("checkked")=checkked
rs("E_BigClassID")=E_BigClassID
rs("E_SmallClassID")=E_SmallClassID
rs("SpecialID")=SpecialID
rs("SpecialID2")=SpecialID2
rs("EnCode")=EnCode
rs("newsrelated")=newsrelated
rs("about")=about
rs("picname")=PicUrl
if title_type="" or title_type="0" then
	rs("titletype")="l"
else
	rs("titletype")=title_type
end if
rs("titlecolor")=title_color

if title_face="" then
	rs("titleface")="无"
else
	rs("titleface")=CheckStr(title_face)
end if

rs("titlesize")=title_size
rs("E_typeid")=E_typeid
rs("UpdateTime")=now()
rs.update
newsid=rs("newsid")

dim username,dep1,dep2
username=request.cookies(eChsys)("username")
set rs2=server.createobject("adodb.recordset")
sql2="select * from "& db_User_Table &" where  "& db_User_Name &"='"&username&"'"
rs2.open sql2,ConnUser,1,3
dep1=rs2("E_DepName")
dep2=rs2("E_DepType")
rs2("number")=rs2("number")+1
rs2.update
rs2.close
set rs2=nothing

set rs9=server.createobject("adodb.recordset")
sql9="select * from "& db_EC_Dep_Table &" where  ( E_DepName='"&dep1&"'  and E_DepType='"&dep2&"' ) "
		rs9.open sql9,Conn,1,3
If(not rs9.eof) or not rs9.bof Then
	rs9("depnumber")=rs9("depnumber")+1
	rs9.update
End if
rs9.close
set rs9=nothing

rs.Close
set rs1=nothing

set rs3=server.createobject("adodb.recordset")
sql3="select * from "& db_UploadPic_Table &" where username='"&request.cookies(eChsys)("username")&"'" 
rs3.open sql3,conn,1,3
do while not rs3.eof
	set rs4=server.createobject("adodb.recordset")
	sql4="select * from "& db_EC_Attach_Table &"" 
	rs4.open sql4,conn,1,3
	filename=rs3("picname")
	rs4.addnew
	rs4("filename")=filename
	rs4("newsid")=newsid
	rs4.update
	rs4.close
	set rs4=nothing
	RS3.MoveNext
loop
conn.execute("delete from "& db_UploadPic_Table &" where username='"&request.cookies(eChsys)("username")&"'")
rs3.close
set rs3=nothing
 
for i=1 to ubound(arrnew)
	set rs4=server.createobject("adodb.recordset")
	sql4="select * from "& db_EC_Attach_Table &"" 
	rs4.open sql4,conn,1,3
	filename=arrnew(i)
	rs4.addnew
	rs4("filename")=filename
	rs4("newsid")=newsid
	rs4.update
	rs4.close
	set rs4=nothing
next

set rs5=server.createobject("adodb.recordset")
sql5="select * from "& db_EC_Attach_Table &"" 
rs5.open sql5,conn,1,3
filename=aassss1
rs5.addnew
rs5("filename")=filename
rs5("newsid")=newsid
rs5.update
rs5.close
set rs4=nothing

Response.cookies(eChsys)("content")=""
Response.cookies(eChsys)("newsrelated")=""

if request.cookies(eChsys)("key")="typemaster" or request.cookies(eChsys)("key")="bigmaster" or request.cookies(eChsys)("key")="smallmaster" then
	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=E_TypeManage.asp"">"
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
end if
Show_Message("<p align=center><font color=red>恭喜您，文章“"&title&"”已经成功添加!<br><br>"&freetime&"秒钟后返回上页!</font>")
response.end
%>