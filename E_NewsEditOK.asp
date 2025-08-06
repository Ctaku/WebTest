<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkURL.asp"-->
<!--#include file="inc/config2.asp"-->
<%
Server.ScriptTimeout=999
if not(request.cookies(eChsys)("KEY")="super" or request.cookies(eChsys)("KEY")="typemaster" or request.cookies(eChsys)("KEY")="bigmaster" or request.cookies(eChsys)("key")="smallmaster" or request.cookies(eChsys)("key")="selfreg" or (request.cookies(eChsys)("KEY")="check" and checkmod="1"))  then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	NewsID=ChkRequest(Request.QueryString("NewsID"),1)	'防注入
	title=trim(request.form("title"))
	if Title="" then
		Show_Err("请填写文章标题！<br><br><a href='javascript:history.back()'>返回</a>")
		response.end
	end if
	if request.cookies(eChsys)("key")="super" then
		if request("viewhtml")<>"" then 
			Show_Err("请取消查看HTML源代码后再添加！<br><br><a href='javascript:history.back()'>返回</a>")
			response.end
		end if
	end if
		
	Author=CheckStr(trim(Request.Form("Author")))
	Original=CheckStr(trim(Request.Form("Original")))
	about=CheckStr(trim(Request.Form("about")))

	dim SavePath,SaveSecondPath,FileExtPath
	SaveSecondPath=year(now)&"-"&month(now)
	FileExtPath="gethttppic"				'偷图存放的二级分类目录
	SavePath = FileUploadPath   '存放上传文件的目录
	if right(SavePath,1)<>"/" then SavePath=SavePath&"/" '在目录后加(/)

	Dim strData
	Dim intFieldCount
	Dim i
	
	intFieldCount = Request.Form("hdnCount")
	
	For i=1 To intFieldCount
		content = content & Request.Form("hdnBigfield" & i)
	Next

	'删除被添加的本站WEB路径,即将引用本站资源的绝对改为相对路径
	if Request.ServerVariables("SERVER_PORT")="80" then
		weburl="http://"& Cstr(Request.ServerVariables("SERVER_NAME")) & Cstr(Request.ServerVariables("url"))
	else
		weburl="http://"& Cstr(Request.ServerVariables("SERVER_NAME")) &":"& Request.ServerVariables("SERVER_PORT") & Cstr(Request.ServerVariables("url"))
	end if
	'注意,下一行中被替换字符串应为实际的本文件名
	weburl=replace(weburl,"E_NewsEditOK.asp","",1,-1,1)
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
			SaveFileName=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&cstr(i&mid(arrimg(i),instrrev(arrimg(i),".")))

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
		arrnew(i)=replace(arrnew(i),FileUploadPath,FileUploadPath)
		strs=replace(strs,arrall(i),arrnew(i))
		arrnew(i)=replace(arrnew(i),FileUploadPath,"")
		if PicUrl=arrall(i) then
			PicUrl=arrnew(i)
		end if
	next
	content=strs
	
	if left(Picurl,4)="http" then
		strChkExtFileName = Trim(Picurl)
		Forumupload = Split(UpFileType,"|")		'UpFileType为inc/config2.asp文件中相关值

		For ii=0 To ubound(Forumupload)
			If GetHttpfileExt(strChkExtFileName) = Trim(Forumupload(ii)) Then	'由Picurl传过来的文件类型检测

				randomize
				ranNum=int(900*rnd)+100
				SaveFileName=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&cstr(i&mid(Picurl,instrrev(Picurl,".")))
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
		
	if request.Form("PicUrl")="" then
		picnews=0
	else
		picnews=1
	end if
	Picurl=replace(Picurl,FileUploadPath,"")
	
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
	
	checkked=request.form("checkked")

		if request.cookies(eChsys)("key")="smallmaster" then		


			if request.cookies(eChsys)("shenhe")="0" then
				checkked=1
			else
				checkked=0
			End if
		
		
		end If

	SpecialID=Request.Form("SpecialID")
	SpecialID2=Request.Form("SpecialID2")
	EnCode=trim(Request.Form("EnCode"))
	newslevel=Request.Form("newslevel")
	title_color=Request.Form("title_color")
	title_type=Request.Form("title_type")
	title_face=Request.Form("title_face")
	title_size=Request.Form("title_size")

	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_EC_News_Table &" where NewsID=" & NewsID
	rs.open sql,conn,3,3

	rs("title")=title
	rs("Author")=Author

	'修改上传文件路径标记为[="uploadfile/]
	content=replace(content,"="&chr(34)&FileUploadPath,"="&chr(34)&"uploadfile/",1,-1,1)
	rs("content")=content
	rs("Original")=Original
	rs("goodnews")=goodnews
	rs("daodu")=daodu
	rs("istop")=istop
	rs("backtype")=backtype
	rs("backtype1")=backtype1
	rs("picnews")=picnews
	rs("checkked")=checkked
	rs("SpecialID")=SpecialID
	rs("SpecialID2")=SpecialID2
	rs("EnCode")=EnCode
	if newslevel="" then
		rs("newslevel")=0
	else
		rs("newslevel")=newslevel
	end if
	rs("about")=about
	rs("picname")=PicUrl
	'rs("UpdateTime")=now()
	dim E_BigClassID,E_SmallClassID,E_typeid
	E_typeid=rs("E_typeid")
	E_BigClassID=rs("E_BigClassID")
	E_SmallClassID=rs("E_SmallClassID")
	if title_type="" or title_type="0" then
		rs("titletype")="l"
	else
		rs("titletype")=title_type
	end if
	rs("titlecolor")=title_color
	
	if title_face="" then
		rs("titleface")="无"
	else
		rs("titleface")=title_face
	end if
	
	rs("titlesize")=title_size
	rs.update
	rs.Close
	set rs=nothing
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
	set rs5=nothing

	if E_SmallClassID<>"" then
		response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=listnews.asp?E_SmallClassID="&E_SmallClassID&""">"
		Show_Message("<p align=center><font color=red>恭喜您，文章“"&title&"”已经成功修改[1]!<br><br>"&freetime&"秒钟后返回上页!</font>")
		response.end
	else
		if E_BigClassID<>"" then
		response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=E_SmallNO.asp?E_BigClassID="&E_BigClassID&""">"
		Show_Message("<p align=center><font color=red>恭喜您，文章“"&title&"”已经成功修改[2]!<br><br>"&freetime&"秒钟后返回上页!</font>")
		response.end
		else
			response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=E_bigclassNO.asp?E_typeid="&E_typeid&""">"
			Show_Message("<p align=center><font color=red>恭喜您，文章“"&title&"”已经成功修改[2]!<br><br>"&freetime&"秒钟后返回上页!</font>")
			response.end
			end if
	end if
end if%>