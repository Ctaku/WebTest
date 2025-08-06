<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<%
    '****************************************************************************
    '' @功能说明: 根据所提供的classid值和newsid及查找方向来返回相同类下的与当前文
	''            章相邻文章的newsid
    '' @参数说明:  fx>0 查找下一条，否则查找上一条
    '' @返回值:   NewsID，找不到或空记录为0
    '****************************************************************************
Function getSideNewsID(ClassId, NewsId, FX)
    Set rs = server.CreateObject("adodb.recordset")
    sql = "select newsid from EC_News where E_BigClassID=" & ClassId
    rs.Open sql, Conn, 1, 1
    If rs.EOF And rs.BOF Then
        getSideNewsID = 0
        Exit Function
    End If
    
    rs.MoveFirst
    rs.Find "newsid=" & NewsId
    
    If FX > 0 Then
        rs.MoveNext
    Else
        rs.MovePrevious
    End If
    
    If rs.EOF Or rs.BOF Then
        getSideNewsID = 0
    Else
        getSideNewsID = rs(0)
    End If
    
    rs.Close
    Set rs = Nothing
End Function

    '****************************************************************************
    '' @功能说明: 获取指定newsid的标题
    '' @参数说明:  
    '' @返回值:   
    '****************************************************************************
Function getNewsTitle(NewsId)
    If NewsId = 0 Then
        getNewsTitle = ""
        Exit Function
    Else
        Set rs = server.CreateObject("adodb.recordset")
        sql = "select title from EC_News where newsid=" & NewsId
        rs.Open sql, conn, 1, 1
        getNewsTitle = rs(0)
        rs.Close
        Set rs = Nothing
    End If
End Function

	
    '****************************************************************************
    '' @功能说明: 获得文件扩展名
    '' @参数说明:  
    '' @返回值:   
    '****************************************************************************

function getFileExtName(fileName)
	dim pos
	pos=instrrev(filename,".")
	if pos>0 then
		getFileExtName=mid(fileName,pos+1)
	else
		getFileExtName=""
	end if
end function

function gotTopic(str,strlen)
	dim l,t,c
	l=len(str)
	t=0
	for i=1 to l
		c=Abs(Asc(Mid(str,i,1)))
		if c>255 then
			t=t+2
		else
			t=t+1
		end if
		if t>=strlen then
			gotTopic=left(str,i)&""
			exit for
		else
			gotTopic=str&""
		end if
	next
end function

    '****************************************************************************
    '' @功能说明: 计算源字符串Str的长度(一个中文字符为2个字节长)
    '' @参数说明:  - str [string]: 源字符串
    '' @返回值:   - [Int] 源字符串的长度
    '****************************************************************************
 Public Function strLen(Str)
  If Trim(Str)="" Or IsNull(str) Then 
   strlen=0
  else
   Dim P_len,x
   P_len=0
   StrLen=0
   P_len=Len(Trim(Str))
   For x=1 To P_len
    If Asc(Mid(Str,x,1))<0 Then
     StrLen=Int(StrLen) + 2
    Else
     StrLen=Int(StrLen) + 1
    End If
   Next
  end if
 End Function

    '****************************************************************************
    '' @功能说明: 截取源字符串Str的前LenNum个字符(一个中文字符为2个字节长)
    '' @参数说明:  - str [string]: 源字符串
    '' @参数说明:  - LenNum [int]: 截取的长度
    '' @返回值:   - [string]: 转换后的字符串
    '****************************************************************************
 Public Function CutStr(Str,LenNum)
  Dim P_num
  Dim I,X
  If StrLen(Str)<=LenNum Then
   Cutstr=Str
  Else
   P_num=0
   X=0
   Do While Not P_num > LenNum-2
    X=X+1
    If Asc(Mid(Str,X,1))<0 Then
     P_num=Int(P_num) + 2
    Else
     P_num=Int(P_num) + 1
    End If
    Cutstr=Left(Trim(Str),X)&".."
   Loop
  End If
 End Function


    '****************************************************************************
    '' @功能说明: 将字符串中的str中的HTML代码进行过滤
    '' @参数说明:  - str [string]: 源字符串
    '' @返回值:   - [string] 转换后的字符串
    '****************************************************************************

function nohtml(str)
    dim re
    If(isnull(str)) Then
	nohtml=str
    else
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	re.Pattern="(\<.[^\<]*\>)"
	str=re.replace(str," ")
	re.Pattern="(\<\/[^\<]*\>)"
	str=re.replace(str," ")
	nohtml=str
    End if
    set re=nothing
end function

'替换一切 HTML 标记
function nohtmlcode(str)
	lsstart=instr(1,str,"<",1)
	while lsstart>0
		lsstart=instr(1,str,"<",0)
		if lsstart>0 then
			lsend=instr(lsstart,str,">",0)+1
			lstemp=mid(str,lsstart,lsend-lsstart)
			str=replace(str,lstemp,"",1,-1,1)	'
		end if
	wend

	str=replace(str,"&nbsp;","",1,-1,1) '空格
	str=replace(str,vbcrlf,"",1,-1,1)	'换行符
	str=replace(str,chr(13),"",1,-1,1)	'回车
	str=replace(str,chr(32),"",1,-1,1)	'空格
	str=replace(str,chr(34),"",1,-1,1)	'双引号
	str=replace(str,chr(39),"",1,-1,1)	'单引号
	str=replace(str,"'","",1,-1,1)	'单引号
	str=replace(str,"=","",1,-1,1)	'双引号
	str=replace(str,"/","",1,-1,1)	'单引号
	str=replace(str,"&copy;","",1,-1,1)	'版权号
	str=replace(str,"&reg;","",1,-1,1)	'
	str=replace(str,"&amp;","",1,-1,1)	'
	str=replace(str,"&","",1,-1,1)	'
	str=replace(str,"#","",1,-1,1)	'
	str=replace(str,"<BR>"," ",1,-1,1)	'
	str=replace(str,"<p>"," ",1,-1,1)	'
	str=replace(str,"</p>","",1,-1,1)	'
	str=replace(str,"<B>","",1,-1,1)	'
	str=replace(str,"</B>","",1,-1,1)	'
	str=replace(str,"\n","",1,-1,1)	'
	str=replace(str,"<","",1,-1,1)	'
	str=replace(str,">","",1,-1,1)	'
	nohtmlcode=str
end function
'****************************************************************************
'' @功能说明: 获取本系统的http安装路径,包括获取非默认端口
'' @参数说明:  - str [strFileName]: 调用该定义的程序文件名
'' @返回值:   - str [ServerHttpUrl]: http路径，如 http://www.xxx.com/news/
'****************************************************************************
Function ServerHttpUrl(strFileName)
	dim weburl
	if Request.ServerVariables("SERVER_PORT")="80" then
		weburl="http://"& Cstr(Request.ServerVariables("SERVER_NAME")) & Cstr(Request.ServerVariables("url"))
	else
		weburl="http://"& Cstr(Request.ServerVariables("SERVER_NAME")) &":"& Request.ServerVariables("SERVER_PORT") & Cstr(Request.ServerVariables("url"))
	end if
	'注意,下一行中被替换字符串应为实际的本文件名
	weburl=replace(weburl,strFileName,"",1,-1,1)
	ServerHttpUrl=weburl
end Function
%>