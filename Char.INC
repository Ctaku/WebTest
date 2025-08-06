<%

'html字符转换过滤，完全模式

function htmlencode(str)
	dim result
	dim l
	if isNULL(str) then 
		htmlencode=""
		exit function
	end if
	l=len(str)
	result=""
	dim i
	for i = 1 to l
		select case mid(str,i,1)
			case "<"
				result=result+"&lt;"
			case ">"
				result=result+"&gt;"
			case chr(13)
				result=result+"<br>"
			case chr(34)
				result=result+"&quot;"
			case "&"
				result=result+"&amp;"
			case chr(32)
				'result=result+"&nbsp;"
				if i+1<=l and i-1>0 then
					if mid(str,i+1,1)=chr(32) or mid(str,i+1,1)=chr(9) or mid(str,i-1,1)=chr(32) or mid(str,i-1,1)=chr(9)  then	                      
						result=result+"&nbsp;"
					else
						result=result+" "
					end if
				else
					result=result+"&nbsp;"
				end if
			case chr(9)
				result=result+"    "
			case else
				result=result+mid(str,i,1)
		end select
	next 
	htmlencode=result
end function

'html字符转换过滤，模式1

function htmlencode1(fString)
if fString<>"" and not isnull(fString) then

	fString = replace(fString, "&gt;", ">")
	fString = replace(fString, "&lt;", "<")
	fString = Replace(fString, "&nbsp;", chr(32))
	fString = Replace(fString, "</P><P>", CHR(10) & CHR(10))
	fString = Replace(fString, "<BR>", CHR(10))

	htmlencode1=fString
else
	htmlencode1=""
end if
end function

'html字符转换过滤，模式2

function htmlencode2(fString)
if fString<>"" and not isnull(fString) then

	fString = Replace(fString, chr(32), "&nbsp;")
	fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
	fString = Replace(fString, CHR(10), "<BR>")
	htmlencode2=fString
else
	htmlencode2=""
end if
end function

'html字符转换过滤，模式3

function htmlencode3(fString)
if fString<>"" and not isnull(fString) then

	fString = replace(fString, "&gt;", ">")
	fString = replace(fString, "&lt;", "<")

	htmlencode3=fString
else
	htmlencode3=""
end if
end function

'html字符转换过滤，模式4，不能过滤<>，因为文章标题有字体效果

function htmlencode4(fString)
if fString<>"" and not isnull(fString) then

	fString = replace(fString, "<br>", "")

	htmlencode4=fString
else
	htmlencode4=""
end if
end function



'函数

function sustainhtml(str)
    dim result
    dim l
    if isNULL(str) then 
       sustainhtml=""
       exit function
    end if
    l=len(str)
    result=""
	dim i
	for i = 1 to l
	    select case mid(str,i,1)
	           case chr(13)
	                result=result+"<br>"
	           case chr(34)
	                result=result+"&quot;"
                   case chr(32)	           
	                'result=result+"&nbsp;"
	                if i+1<=l and i-1>0 then
	                   if mid(str,i+1,1)=chr(32) or mid(str,i+1,1)=chr(9) or mid(str,i-1,1)=chr(32) or mid(str,i-1,1)=chr(9)  then	                      
	                      result=result+"&nbsp;"
	                   else
	                      result=result+" "
	                   end if
	                else
	                   result=result+"&nbsp;"	                    
	                end if
	           case "&"
	                result=result+"&amp;"
	           case chr(9)
	                result=result+"    "
	           case else
	                result=result+mid(str,i,1)
         end select
       next 
       sustainhtml=result
   end function

'检查sql字符串中是否有单引号，有则进行转化，支持部分HTML的时候采用，过滤危险js代码

   function CheckStr(str)
        dim tstr,l,i,ch
    l=len(str)
    for i=1 to l
        ch=mid(str,i,1)
        if ch="'" then
        tstr=tstr+"'"
     end if
        if ch="(" then
        tstr=tstr+"("
     end if
     tstr=tstr+ch
                if ch="[" then
        tstr=tstr+"["
     end if
                if ch=" " then
        tstr=tstr+""
     end if
     
     
    next
    CheckStr=tstr
   end function

'检查email合法性

function IsValidEmail(email)

dim names, name, i, c

'Check for valid syntax in an email address.

IsValidEmail = true
names = Split(email, "@")
if UBound(names) <> 1 then
   IsValidEmail = false
   exit function
end if
for each name in names
   if Len(name) <= 0 then
     IsValidEmail = false
     exit function
   end if
   for i = 1 to Len(name)
     c = Lcase(Mid(name, i, 1))
     if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
       IsValidEmail = false
       exit function
     end if
   next
   if Left(name, 1) = "." or Right(name, 1) = "." then
      IsValidEmail = false
      exit function
   end if
next
if InStr(names(1), ".") <= 0 then
   IsValidEmail = false
   exit function
end if
i = Len(names(1)) - InStrRev(names(1), ".")
if i <> 2 and i <> 3 then
   IsValidEmail = false
   exit function
end if
if InStr(email, "..") > 0 then
   IsValidEmail = false
end if

end function

'SQL注入过滤
'--- 传入参数 ---
'ParaName:参数名称-字符型
'ParaType:参数类型-数字型(1表示以上参数是数字，0表示以上参数为字符)
 Public Function ChkRequest(ParaName,ParaType)
	Dim ParaValue
	If ParaType=1 then
		ParaValue=ParaName
		 If Not (ParaValue="") then
			If not isNumeric(ParaValue)  then
				ParaValue=""
				Show_Err("参数" & ParaName & "必须为数字型！")
				Response.end
			End if
		End if
	Else
		ParaValue=trim(ParaName)
		ParaValue=replace(ParaValue,"'","''")
	End if
	ChkRequest=ParaValue
End function

function Show_Err(str)%>

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%></title>
</head>
<body topmargin="0">
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFDFBF" width="100%" id="AutoNumber1">
<tr> 
	<td width="100%" height="5" >&nbsp;</td>
</tr>
<tr> 
	<td width="100%" height="100" align="center"><font face="楷体_GB2312" size="5" color=red><b><%if str="" then%>对不起，你没有执行该操作的权限！<%else%><%=str%><%end if%></td></font>
	</td>
</tr>
<tr> 
	<td width="100%" height="25" align="center">如发现程序中存在问题，请及时和我们联系，谢谢！</td>
</tr>
</table>
</body>
</html>
<%
end function

function Show_Message(str)%>

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%></title>
</head>
<body topmargin="0">
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFDFBF" width="100%" id="AutoNumber1">
<tr> 
	<td width="100%" height="5" >&nbsp;</td>
</tr>
<tr> 
	<td width="100%" height="100" align="center"><font face="楷体_GB2312" size="5" color=red><b><%if str="" then%>对不起，你没有执行该操作的权限！<%else%><%=str%><%end if%></td></font>
	</td>
</tr>
<tr> 
	<td width="100%" height="25" align="center">如发现程序中存在问题，请及时和我们联系，谢谢！</td>
</tr>
</table>
</body>
</html>
<%end function%>