<% @language="vbscript" %>
<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp"-->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp"-->
<%
IF request.cookies(eChsys)("ManageKEY")<>"super" then
	Show_Err("对不起，您的后台管理权限不够！")
	response.end
else

	Set rs9 = Server.CreateObject("ADODB.Recordset")
	sql9 ="SELECT * From "& db_EC_System_Table &" Order By id DESC"
	RS9.open sql9,Conn,1,1
	if rs9("system")="1" or request.cookies(eChsys)("ManagePurview")="99999" then

	xpurl=request("xpurl")
	email=request("email")
	QQ=request("QQ")
	copyright=request("copyright")
	version=request("version")
	ver=request("ver")
	Address=request("Address")
	zip=request("zip")
	tel=request("tel")
	gd1=request("gd1")
	gd2=request("gd2")
	B_BG=request("B_BG")
	
	top_news=request("top_news")
	L_BG=request("L_BG")
	bigclassshownum=request("bigclassshownum")
	top_sp=request("top_sp")
	top_txt=request("top_txt")
	top_img=request("top_img")
	newsnum=request("newsnum")
	picnum=request("picnum")
	topuser=request("topuser")
	linkshownum=request("linkshownum")
	reviewnum=request("reviewnum")
	
	logo=request("logo")
	logourl=request("logourl")
	TopBanner=request("TopBanner")
	TopBannerurl=request("TopBannerurl")
	OtherTopBanner=request("OtherTopBanner")
	OtherTopBannerurl=request("OtherTopBannerurl")
	moveurl=request("moveurl")
	E_announceview=request("E_announceview")


	
	gg1=server.htmlencode(request("gg1"))
	gg1=CheckStr(gg1)
	SiteName=server.htmlencode(request("SiteName"))
	SiteName=CheckStr(SiteName)
	Address=server.htmlencode(request("Address"))
	Address=CheckStr(Address)
	CopyRight=server.htmlencode(request("CopyRight"))
	CopyRight=CheckStr(CopyRight)
	
	if DbType="ACCESS" then
		version=""
	end if
	
	if DbType="MSSQL" then
		version=""
	end if
	
	if DbType="MYSQL" then
		version=""
	end if
	
	ver=""
	R_BG=request("R_BG")
	AD_Show=request("AD_Show")
	ad_class=request("ad_class")
	R_TOP=request("R_TOP")
	R_MAIN=request("R_MAIN")
	L_MAIN=request("L_MAIN")
	
	If SiteName=""  or xpurl="" or email="" Then
		Show_Err("对不起,看看哪里忘记填写了?<br><br><a href='javascript:history.back()'>回去重来</a>")
		response.end
	end if
	
	if len(gg1)>=65535 then
		Show_Err("本站公告太长了，将会破坏页面的整体风格<br><br><a href='javascript:history.back()'>回去重来</a>")
		response.end
	end if
	
	if len(url)>=100 then
		Show_Err("你的主页地址太长了<br><br><a href='javascript:history.back()'>回去重来</a>")
		response.end
	end if
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	
	sql="update "& db_EC_System_Table &" set SiteName='"&SiteName&"',xpurl='"&xpurl&"',bigclassshownum='"&bigclassshownum&"',reviewnum='"&reviewnum&"',linkshownum='"&linkshownum&"',topuser='"&topuser&"',email='"&email&"',QQ='"&QQ&"',copyright='"&copyright&"',Address='"&Address&"',Zip='"&zip&"',Tel='"&Tel&"',gd1='"&gd1&"',gd2='"&gd2&"',B_BG='"&B_BG&"',top_news="&top_news&",L_BG="&L_BG&",top_sp="&top_sp&",R_BG="&R_BG&",AD_Show="&AD_Show&",ad_class="&ad_class&",top_txt="&top_txt&",top_img="&top_img&",gg1='"&gg1&"',R_TOP='"&R_TOP&"',R_MAIN='"&R_MAIN&"',L_MAIN='"&L_MAIN&"',newsnum='"&newsnum&"',picnum='"&picnum&"',ver='"&ver&"',version='"&version&"',logo='"&logo&"',logourl='"&logourl&"',Topbanner='"&Topbanner&"',Topbannerurl='"&Topbannerurl&"',OtherTopbanner='"&OtherTopbanner&"',OtherTopbannerurl='"&OtherTopbannerurl&"',moveurl='"&moveurl&"',E_announceview='"&E_announceview&"'"
	conn.execute(sql)
	conn.close
	
	set conn=nothing
	set fso=Server.CreateObject("Scripting.FileSystemObject")
	set hf=fso.CreateTextFile(Server.mappath("inc/config.asp"),true)
	hf.write "<" & "%" & vbcrlf
	hf.write "Const MaxFileSize=" & chr(34) & trim(request("MaxFileSize")) & chr(34) & "        '上传文件大小限制" & vbcrlf
	hf.write "Const SaveUpFilesPath=" & chr(34) & "UploadFile" & chr(34) & "        		'存放上传文件的目录" & vbcrlf
	hf.write "Const UpFileType=" & chr(34) & trim(request("UpFileType")) & chr(34) & "        '允许的上传文件类型" & vbcrlf
	hf.write "Const byteType=" & chr(34) & trim(request("byteType")) & chr(34) & "        '留言本过滤词语" & vbcrlf
    hf.write "Const byteipType=" & chr(34) & trim(request("byteipType")) & chr(34) & "        '网站恶意ＩＰ地址留言屏蔽" & vbcrlf
	hf.write "Const bytezfType=" & chr(34) & trim(request("bytezfType")) & chr(34) & "        '网站恶意广告留言字符屏蔽" & vbcrlf
	hf.write "Const newsipType=" & chr(34) & trim(request("newsipType")) & chr(34) & "        '浏览文章IP限制设置(填写为允许、发文需要相应设置)" & vbcrlf
	hf.write "Const echuangmenu=" & chr(34) & trim(request("echuangmenu")) & chr(34) & "        '自定义菜单" & vbcrlf
	hf.write "Const BOTTOMmenu=" & chr(34) & trim(request("BOTTOMmenu")) & chr(34) & "	'自定义底部菜单" & vbcrlf
	hf.write "%" & ">"
	hf.close
	set hf=nothing
	set fso=nothing
	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=E_System.asp"">"
	Show_Message("<p align=center><font color=red>已经成功添加到数据库!<br><br>"&freetime&"秒钟后返回上页!</font>")
	response.end

	rs9.close
		set rs9=nothing
		conn.close
		set conn=nothing
	else
		Show_Err("对不起，该功能已经被超级系统管理员关闭，您没有权限操作！")
		response.end
	end if
end if%>