<script language="JavaScript">
<!--
/*��������:�ȷɵ��Լ�����(http://www.xfbbs.com)��������!*/
function check(input){                                                                                             
if(input.mailSelect.options.selectedIndex==0){                                                                                             
alert("��ѡ�����ĵ������䣡");                                                                                             
input.mailSelect.focus();                                                                                             
return false;}                                                                                             
if(input.name.value==""){                                                                                             
alert("�����������ʺţ�");                                                                                             
input.name.focus();                                                                                             
return false;}                                                                                             
if(input.password.value=="" || input.password.value.length<3){                                                                                             
alert("����ȷ�����������룡");                                                                                             
input.password.focus();                                                                                             
return false;}                                                                                             
else{go();                                                                                             
return false;}}                                                                                             
function makeURL(){                                                                                             
var objForm=document.mailForm;                                                                                             
var intIndex=objForm.mailSelect.options.selectedIndex;                                                                                             
var varInfo=objForm.mailSelect.options[intIndex].value; /*��ȡ�ı����ʼ����������û��˺ź�������Ϣ*/                                                                                             
var arrayInfo=varInfo.split(';'); /*�����ϻ�ȡ����Ϣ���зָ�������������*/                                                                                             
var strName=objForm.name.value,varPasswd=objForm.password.value;                                                                                             
var length=arrayInfo.length,strProvider=arrayInfo[0],strIdName=arrayInfo[1],varPassName=arrayInfo[2];                                                                                             
if(length==3){                                                                                             
var strUrl='http://'+strProvider+'?'+strIdName+'='+strName+'&'+varPassName+'='+varPasswd; /*�ϲ��ַ������õ����硰http://mail.sina.com.cn/cgi-bin/log...�����ַ�����URL*/                                                                                             
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
if(strLocation.indexOf('<form name="tmpForm"')!=-1){/*����ֻ���á�post������ȡ�����ݵ�����ʹ���Զ��ύ����ʱ��*/                                                                                             
outWin=window.open('','','scrollbars=yes,menubar=yes,toolbar=yes,location=yes,status=yes,resizable=yes');                                                                                             
doc=outWin.document;                                                                                             
doc.open('text/html');                                                                                             
doc.write('<html><head><meta http-equiv="Content-Type" content="text/html; charset=gb2312"><title>�����¼</title></head><body onload="document.tmpForm.submit()">');                                                                                             
doc.write('<p align="center"><br></br><br></br>��ӭ��ʹ�õ����ʾֿ��ٵ�¼ϵͳ��<br></br><br></br>�� �� �򡭡�����<br></br><br></br>http://www.XFBBS.Com ��̳�ṩ</p>'+strLocation+'</body></html>');                                                                                             
doc.close();                                                                                             
}                                                                                             
else window.open(strLocation,'','scrollbars=yes,menubar=yes,toolbar=yes,location=yes,status=yes,resizable=yes');                                                                                             
}                                                                                             
-->     
</script>
<TABLE border=0 align=center>
<form method=post name=mailForm target=��blank onsubmit='return check(this)'>
<TR>
<td>
<select name=mailSelect size=1>
<option>��ѵ��������¼</option>
<option value=login.mail.sohu.com/chkpwd.php;UserName;Password;post>�Ѻ�sohu.com</option>
<option value=edit.bjs.yahoo.com/config/login;login;passwd;post>�����Ż�</option>
<option value=reg4.163.com/in.jsp?url=http://reg4.163.com/EnterEmail.jsp?username=window.document.mailForm.name.value;username;password;post>����163.com</option>
<option value=bjweb.163.net/cgi/login;user;pass>163.net�����ʾ�</option>
<option value=web.netease.com/cgi/login;user;pass;post>����netease.com</option>
<option value=www.126.com/cgi/login;user;pass;post>����126.com</option>
<option value=mail.sina.com.cn/cgi-bin/login;u;psw>����(sina.com)</option>
<option value=reg4.163.com/CheckUser.jsp;username;password;post>����ͨ��֤</option>
<option value=webmail.21cn.com/NULL/NULL/NULL/NULL/NULL/SignIn.gen;LoginName;passwd;post>21CN.COM</option>
<option value=freemail.inhe.net/cgi-bin/login;LoginName;Password;post>�ӱ�����inhe.net</option>
<option value=mail.china.com/extend/gb/NULL/NULL/NULL/SignIn.gen;LoginName;passwd;post>�л���china.com</option>
<option value=login.chinaren.com/zhs/servlet/Login;username;password;post;hidden;url;http://mail.chinaren.com>
�й���chinaren.com</option>
<option value=lc1.law13.hotmail.passport.com/cgi-bin/dologin;login;passwd;post>hotmail.com</option>
<option value=mw1.elong.com/cgi-bin/weblogon.cgi;username;password;post>e��elong.com</option>
<option value=login.etang.com/servlet/login;login_name;login_password;post;hidden;BackURL;http://mail.etang.com/cgi/door>����etang.com</option>
<option value=mail.fm365.com/cgi-bin/legend/wmaila;username;password;post>FM365.com</option>
<option value=epost.soim.com/member_info.jsp;username;passwd;post>����soim.com</option>
<option value=mail.2911.net/cgi-bin/mail/main.pl;USERNAME;PASSWORD;post>2911�����ʾ�</option>
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
<td><font color=black>�ʺ�:        
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
����:</font><font color=#800000> 
<input type=password size=8 name=password onfocus='this.select()' style=" BORDER-BOTTOM: #679FC2 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #ffffff 1px solid; FONT-SIZE: 9pt">
</font></td>
<td height=23> <p align=right><font color=#800000 size=2> <IMG style="CURSOR: hand" onclick=reset() src="images/button_cx.gif" height=23  width=46>
</font></td>
</tr>
</table></td>
</tr>
</form>
</TABLE>