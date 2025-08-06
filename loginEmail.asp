<script language="JavaScript">
<!--
/*本代码由:先飞电脑技术网(http://www.xfbbs.com)整理完善!*/
function check(input){                                                                                             
if(input.mailSelect.options.selectedIndex==0){                                                                                             
alert("请选择您的电子邮箱！");                                                                                             
input.mailSelect.focus();                                                                                             
return false;}                                                                                             
if(input.name.value==""){                                                                                             
alert("请输入您的帐号！");                                                                                             
input.name.focus();                                                                                             
return false;}                                                                                             
if(input.password.value=="" || input.password.value.length<3){                                                                                             
alert("请正确输入您的密码！");                                                                                             
input.password.focus();                                                                                             
return false;}                                                                                             
else{go();                                                                                             
return false;}}                                                                                             
function makeURL(){                                                                                             
var objForm=document.mailForm;                                                                                             
var intIndex=objForm.mailSelect.options.selectedIndex;                                                                                             
var varInfo=objForm.mailSelect.options[intIndex].value; /*获取的表单中邮件服务器及用户账号和密码信息*/                                                                                             
var arrayInfo=varInfo.split(';'); /*将以上获取的信息进行分割，并赋给数组变量*/                                                                                             
var strName=objForm.name.value,varPasswd=objForm.password.value;                                                                                             
var length=arrayInfo.length,strProvider=arrayInfo[0],strIdName=arrayInfo[1],varPassName=arrayInfo[2];                                                                                             
if(length==3){                                                                                             
var strUrl='http://'+strProvider+'?'+strIdName+'='+strName+'&'+varPassName+'='+varPasswd; /*合并字符串，得到形如“http://mail.sina.com.cn/cgi-bin/log...”的字符串型URL*/                                                                                             
}                                                                                             
else{                                                                                             
var strUrl='<form name="tmpForm" action="http://'+strProvider+'" method="post"><input type="hidden" name="'+strIdName+'" value="'+strName+'"><input type="hidden" name="'+varPassName+'" value="'+varPasswd+'">';                                                                                             
if(arrayInfo[4]=='hidden') strUrl+='<input type="hidden" name="'+arrayInfo[5]+'" value="'+arrayInfo[6]+'">'                                                                                             
strUrl+='</form>';                                                                                             
}                                                                                             
return strUrl;                                                                                             
}                                                                                             
function go(){                                                                                             
var strLocation=makeURL();                                                                                             
if(strLocation.indexOf('<form name="tmpForm"')!=-1){/*对于只能用“post”来获取表单数据的邮箱使用自动提交的临时表单*/                                                                                             
outWin=window.open('','','scrollbars=yes,menubar=yes,toolbar=yes,location=yes,status=yes,resizable=yes');                                                                                             
doc=outWin.document;                                                                                             
doc.open('text/html');                                                                                             
doc.write('<html><head><meta http-equiv="Content-Type" content="text/html; charset=gb2312"><title>邮箱登录</title></head><body onload="document.tmpForm.submit()">');                                                                                             
doc.write('<p align="center"><br></br><br></br>欢迎您使用电子邮局快速登录系统，<br></br><br></br>请 稍 候…………<br></br><br></br>http://www.XFBBS.Com 论坛提供</p>'+strLocation+'</body></html>');                                                                                             
doc.close();                                                                                             
}                                                                                             
else window.open(strLocation,'','scrollbars=yes,menubar=yes,toolbar=yes,location=yes,status=yes,resizable=yes');                                                                                             
}                                                                                             
-->     
</script>
<TABLE border=0 align=center>
<form method=post name=mailForm target=＿blank onsubmit='return check(this)'>
<TR>
<td>
<select name=mailSelect size=1>
<option>免费电子邮箱登录</option>
<option value=login.mail.sohu.com/chkpwd.php;UserName;Password;post>搜狐sohu.com</option>
<option value=edit.bjs.yahoo.com/config/login;login;passwd;post>中文雅虎</option>
<option value=reg4.163.com/in.jsp?url=http://reg4.163.com/EnterEmail.jsp?username=window.document.mailForm.name.value;username;password;post>网易163.com</option>
<option value=bjweb.163.net/cgi/login;user;pass>163.net电子邮局</option>
<option value=web.netease.com/cgi/login;user;pass;post>网易netease.com</option>
<option value=www.126.com/cgi/login;user;pass;post>网易126.com</option>
<option value=mail.sina.com.cn/cgi-bin/login;u;psw>新浪(sina.com)</option>
<option value=reg4.163.com/CheckUser.jsp;username;password;post>网易通行证</option>
<option value=webmail.21cn.com/NULL/NULL/NULL/NULL/NULL/SignIn.gen;LoginName;passwd;post>21CN.COM</option>
<option value=freemail.inhe.net/cgi-bin/login;LoginName;Password;post>河北银河inhe.net</option>
<option value=mail.china.com/extend/gb/NULL/NULL/NULL/SignIn.gen;LoginName;passwd;post>中华网china.com</option>
<option value=login.chinaren.com/zhs/servlet/Login;username;password;post;hidden;url;http://mail.chinaren.com>
中国人chinaren.com</option>
<option value=lc1.law13.hotmail.passport.com/cgi-bin/dologin;login;passwd;post>hotmail.com</option>
<option value=mw1.elong.com/cgi-bin/weblogon.cgi;username;password;post>e龙elong.com</option>
<option value=login.etang.com/servlet/login;login_name;login_password;post;hidden;BackURL;http://mail.etang.com/cgi/door>亿唐etang.com</option>
<option value=mail.fm365.com/cgi-bin/legend/wmaila;username;password;post>FM365.com</option>
<option value=epost.soim.com/member_info.jsp;username;passwd;post>索易soim.com</option>
<option value=mail.2911.net/cgi-bin/mail/main.pl;USERNAME;PASSWORD;post>2911中文邮局</option>
<option value=02.106.186.230/extend/newgb1/NULL/NULL/NULL/SignIn.gen;LoginName;passwd;post;hidden;DomainName;email.com.cn>@email.com.cn</option>
<option value=www.eyou.com/cgi-bin/login;LoginName;Password;post>@eyou.com</option>
<option value=bjweb.mail.tom.com/cgi/login2;user;pass>TOM.Com</option>
</select>
</td>
</tr>
<tr>
<td>
<table border=0 cellpadding=0 cellspacing=0>
<tr>
<td><font color=black>帐号:        
<input type=text size=8 name=name onfocus='this.select()' style=" BORDER-BOTTOM: #679FC2 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #ffffff 1px solid; FONT-SIZE: 9pt">
</font> </td>
<td>
<p align=right> <font color=#800000 size=2> <input name="Submit" type=image id="Submit2" src="images/button_dl.gif" height="23"  width="46" style="CURSOR: hand" > 
</font> </td>
</tr>
<tr>
<td height=10></td>
</tr>
<tr>
<td><font color=black>
密码:</font><font color=#800000> 
<input type=password size=8 name=password onfocus='this.select()' style=" BORDER-BOTTOM: #679FC2 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #ffffff 1px solid; FONT-SIZE: 9pt">
</font></td>
<td height=23> <p align=right><font color=#800000 size=2> <IMG style="CURSOR: hand" onclick=reset() src="images/button_cx.gif" height=23  width=46>
</font></td>
</tr>
</table></td>
</tr>
</form>
</TABLE>