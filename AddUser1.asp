
	<html>
	<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<title>会员注册</title>
	<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
	<link rel="stylesheet" type="text/css" href="site.css">
	<script LANGUAGE="javascript">
	<!--
	function FrmAddLink_onsubmit() {
	var i, n;
	if (document.FrmAddLink.username.value=="")
		{
		  alert("对不起，请输入您的用户名！")
		  document.FrmAddLink.username.focus()
		  return false
		 }
	else if (document.FrmAddLink.username.value.length < 2)
		{
		  alert("您的用户名能不能长一点！")
		  document.FrmAddLink.username.focus()
		  return false
		 }
	else if (document.FrmAddLink.username.value.length > 30)
		{
		  alert("您的用户名太长了吧！")
		  document.FrmAddLink.username.focus()
		  return false
		 }
	else if (document.FrmAddLink.username.value.indexOf('`') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('~') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('!') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('@') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('#') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('$') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('%') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('^') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('&') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('*') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('(') >= 0 ||
	          document.FrmAddLink.username.value.indexOf(')') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('+') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('{') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('}') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('|') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('[') >= 0 ||
	          document.FrmAddLink.username.value.indexOf(']') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('\\') >= 0 ||
	          document.FrmAddLink.username.value.indexOf(';') >= 0 ||
	          document.FrmAddLink.username.value.indexOf(':') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('>') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('<') >= 0 ||
	          document.FrmAddLink.username.value.indexOf(',') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('?') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('/') >= 0 || 
	          document.FrmAddLink.username.value.indexOf('\'') >= 0 || 
	          document.FrmAddLink.username.value.indexOf('"') >= 0 || 
	          document.FrmAddLink.username.value.indexOf(' ') >= 0 || 
	          document.FrmAddLink.username.value.indexOf('=') >= 0 || 
	          document.FrmAddLink.username.value.indexOf('%') >= 0
	         )  
	          {
	            alert("用户名中包含无效字符，请重新选择用户名！"); 
	            document.FrmAddLink.username.focus();
	            return false;
	          }
	else if (document.FrmAddLink.passwd.value=="")
		{
		  alert("对不起，请您输入密码！")
		  document.FrmAddLink.passwd.focus()
		  return false
		 }
	else if (document.FrmAddLink.passwd.value.length < 4)
		{
		  alert("为了安全，您的密码应该长一点！")
		  document.FrmAddLink.passwd.focus()
		  return false
		 }
	else if (document.FrmAddLink.passwd.value.length > 16)
		{
		  alert("您的密码太长了吧！")
		  document.FrmAddLink.passwd.focus()
		  return false
		 }
	else if (document.FrmAddLink.username.value==document.FrmAddLink.passwd.value)
		{
		  alert("为了安全，用户名与密码不应该相同！")
		  document.FrmAddLink.passwd.focus()
		  return false
		 }
	else if (document.FrmAddLink.passwd2.value=="")
		{
		  alert("对不起，请您输入验证密码！")
		  document.FrmAddLink.passwd2.focus()
		  return false
		 }
	else if (document.FrmAddLink.passwd2.value !== document.FrmAddLink.passwd.value)
		{
		  alert("对不起，您两次输入的密码不一致！")
		  document.FrmAddLink.passwd2.focus()
		  return false
		 }
	else if (document.FrmAddLink.question.value=="")
		{
		  alert("对不起，请您输入提示问题！")
		  document.FrmAddLink.question.focus()
		  return false
		 }
	else if (document.FrmAddLink.answer.value=="")
		{
		  alert("对不起，请您输入问题答案！")
		  document.FrmAddLink.answer.focus()
		  return false
		 }
	else if (document.FrmAddLink.question.value==document.FrmAddLink.answer.value)
		{
		  alert("为了安全，提示问题与问题答案不应该相同！")
		  document.FrmAddLink.answer.focus()
		  return false
		 }
	else if (document.FrmAddLink.fullname.value=="")
		{
		  alert("对不起，请输入您的真实姓名！")
		  document.FrmAddLink.fullname.focus()
		  return false
		 }
	else if (document.FrmAddLink.depid.value=="")
		{
		  alert("对不起，请选择您的工作单位！")
		  document.FrmAddLink.depid.focus()
		  return false
		 }
	else if (document.FrmAddLink.sex.value=="")
		{
		  alert("对不起，请选择您的性别！")
		  document.FrmAddLink.sex.focus()
		  return false
		 }
	else if (document.FrmAddLink.tel.value=="")
		{
		  alert("对不起，请输入您的联系电话！")
		  document.FrmAddLink.tel.focus()
		  return false
		 }
	else if (document.FrmAddLink.email.value=="")
		{
		  alert("对不起，请输入您的电子邮件！")
		  document.FrmAddLink.email.focus()
		  return false
		 }
	else if (document.FrmAddLink.email.value.indexOf("@",0)== -1||document.FrmAddLink.email.value.indexOf(".",0)==-1)
		  {
		  alert("对不起，您输入的电子邮件有误！")
		  document.FrmAddLink.email.focus()
		  return false
		 }
	}
	
	//Function to open pop up window
	function openWin(theURL,winName,features) {
	  	window.open(theURL,winName,features);
	}
	//-->
	</script>
	
	<//生日选择日期处理开始>
<SCRIPT language=JavaScript src="inc/User_Info_Modify.js"></SCRIPT>
	<//生日选择日期处理结束>
	</head>
	<body topmargin="0">    
	<form align="center" method="post" name="FrmAddLink" LANGUAGE="javascript" onsubmit="return FrmAddLink_onsubmit()" action="saveuser.asp">
  <table width="500" border="1" align=center cellpadding="0" cellspacing="0" bordercolor="0" bgcolor="#FFFFFF" style="border-collapse: collapse">
    <tr> 
      <td height="20" colspan="2" align="center" valign="middle"><span class="smallFont">┊<strong> 会 员 注 册 </strong>┊</span></td>
    </tr>
    <tr> 
      <td height="20" colspan="2" align="center" valign="middle">为了使您能正常使用本系统，请详细填写每一项资料</td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right" valign="middle"> <div align="center"><span class="smallFont">用 
          户 名：</span></div></td>
      <td width="264" height="20"> <div align="left"> 
          <input name="username" size="26" class="smallInput" maxlength="30" style="font-family: 宋体; font-size: 9pt" title="请在这里填写您的用户名，注册成功后将用此用户名登录本系统。">
        </div></td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">密&nbsp;&nbsp;&nbsp;码：</span></div></td>
      <td width="264" height="20"> <div align="left"> 
          <input name="passwd" size="26" class="smallInput" maxlength="15" style="font-family: 宋体; font-size: 9pt" type="password"  title="请在这里填写您的登录密码，在登录时本系统将验证您的密码。">
        </div></td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">验证密码：</span></div></td>
      <td width="264" height="20"> <div align="left"> 
          <input name="passwd2" size="26" class="smallInput" maxlength="15" style="font-family: 宋体; font-size: 9pt" type="password" title="请在这里填写您的验证密码，必须与上面的密码一致，主要是防止您的错误输入。">
        </div></td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">提示问题：</span></div></td>
      <td width="264" height="20"> <div align="left"> 
          <input name="question" size="26" class="smallInput" maxlength="50" style="font-family: 宋体; font-size: 9pt" type="text" title="请在这里填写您的提示问题，如果您忘记了密码，可以利用此功能来找回您的密码。">
        </div></td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">问题答案：</span></div></td>
      <td width="264" height="20"> <div align="left"> 
          <input name="answer" size="26" class="smallInput" maxlength="50" style="font-family: 宋体; font-size: 9pt" type="text" title="请在这里填写您的提示问题的答案，如果您忘记了密码，可以利用此功能来找回您的密码。">
        </div></td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">真实姓名：</span></div></td>
      <td width="264" height="20"> <div align="left"> 
          <input name="fullname" size="26" class="smallInput" maxlength="10" style="font-family: 宋体; font-size: 9pt" title="请在这里填写您的真实姓名。">
        </div></td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">工作单位：</span></div></td>
      <td width="264" height="20"> 
        
        <select name="depid" style="font-family: 宋体; font-size: 9pt" title="请在这里选择您的工作单位。">
          <option value="">请选择工作单位</option>
          
          <option value="1"></option>
          
          <option value="2"></option>
          
          <option value="3"></option>
          
          <option value="4"></option>
          
        </select> </td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">性&nbsp;&nbsp;&nbsp; 
          别：</span></div></td>
      <td width="264" height="20"> <select size="1" name="sex" style="font-family: 宋体; font-size: 9pt" title="请在这里填写您的性别。">
          <option selected value="">请选择性别</option>
          
          <option value="先生">先生</option>
          <option value="女士">女士</option>
          <option value="保密">保密</option>
          
        </select> </td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">生&nbsp;&nbsp;&nbsp;日：</span></div></td>
      <td width="264" height="40"> 
        
        <div align="left"> 
          <select size="1" name="birthyear" style="font-family: 宋体; font-size: 9pt">
            
            <option value="1950">1950</option>
            
            <option value="1951">1951</option>
            
            <option value="1952">1952</option>
            
            <option value="1953">1953</option>
            
            <option value="1954">1954</option>
            
            <option value="1955">1955</option>
            
            <option value="1956">1956</option>
            
            <option value="1957">1957</option>
            
            <option value="1958">1958</option>
            
            <option value="1959">1959</option>
            
            <option value="1960">1960</option>
            
            <option value="1961">1961</option>
            
            <option value="1962">1962</option>
            
            <option value="1963">1963</option>
            
            <option value="1964">1964</option>
            
            <option value="1965">1965</option>
            
            <option value="1966">1966</option>
            
            <option value="1967">1967</option>
            
            <option value="1968">1968</option>
            
            <option value="1969">1969</option>
            
            <option value="1970">1970</option>
            
            <option value="1971">1971</option>
            
            <option value="1972">1972</option>
            
            <option value="1973">1973</option>
            
            <option value="1974">1974</option>
            
            <option value="1975">1975</option>
            
            <option value="1976">1976</option>
            
            <option value="1977">1977</option>
            
            <option value="1978">1978</option>
            
            <option value="1979">1979</option>
            
            <option value="1980">1980</option>
            
            <option value="1981">1981</option>
            
            <option value="1982">1982</option>
            
            <option value="1983">1983</option>
            
            <option value="1984">1984</option>
            
            <option value="1985">1985</option>
            
            <option value="1986">1986</option>
            
            <option value="1987">1987</option>
            
            <option value="1988">1988</option>
            
            <option value="1989">1989</option>
            
            <option value="1990">1990</option>
            
            <option value="1991">1991</option>
            
            <option value="1992">1992</option>
            
            <option value="1993">1993</option>
            
            <option value="1994">1994</option>
            
            <option value="1995">1995</option>
            
            <option value="1996">1996</option>
            
            <option value="1997">1997</option>
            
            <option value="1998">1998</option>
            
            <option value="1999">1999</option>
            
            <option value="2000">2000</option>
            
            <option value="2001">2001</option>
            
            <option value="2002">2002</option>
            
            <option value="2003">2003</option>
            
            <option value="2004">2004</option>
            
          </select>
          年 
          <select size="1" name="birthmonth" style="font-family: 宋体; font-size: 9pt">
            
            <option value="1">1</option>
            
            <option value="2">2</option>
            
            <option value="3">3</option>
            
            <option value="4">4</option>
            
            <option value="5">5</option>
            
            <option value="6">6</option>
            
            <option value="7">7</option>
            
            <option value="8">8</option>
            
            <option value="9">9</option>
            
            <option value="10">10</option>
            
            <option value="11">11</option>
            
            <option value="12">12</option>
            
          </select>
          月 
          <select size="1" name="birthday" style="font-family: 宋体; font-size: 9pt">
            
            <option value="1">1</option>
            
            <option value="2">2</option>
            
            <option value="3">3</option>
            
            <option value="4">4</option>
            
            <option value="5">5</option>
            
            <option value="6">6</option>
            
            <option value="7">7</option>
            
            <option value="8">8</option>
            
            <option value="9">9</option>
            
            <option value="10">10</option>
            
            <option value="11">11</option>
            
            <option value="12">12</option>
            
            <option value="13">13</option>
            
            <option value="14">14</option>
            
            <option value="15">15</option>
            
            <option value="16">16</option>
            
            <option value="17">17</option>
            
            <option value="18">18</option>
            
            <option value="19">19</option>
            
            <option value="20">20</option>
            
            <option value="21">21</option>
            
            <option value="22">22</option>
            
            <option value="23">23</option>
            
            <option value="24">24</option>
            
            <option value="25">25</option>
            
            <option value="26">26</option>
            
            <option value="27">27</option>
            
            <option value="28">28</option>
            
            <option value="29">29</option>
            
            <option value="30">30</option>
            
            <option value="31">31</option>
            
          </select>
          日 </div>
        
      </td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">联系电话：</span></div></td>
      <td width="264" height="20" valign="middle"> <input name="tel" size="26" class="smallInput" maxlength="100" style="font-family: 宋体; font-size: 9pt" title="请在这里填写您的联系电话，以便我们与您联系。"></td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">电子邮件：</span></div></td>
      <td width="264" height="20" valign="middle"> <input name="email" size="26" class="smallInput" maxlength="100" style="font-family: 宋体; font-size: 9pt" title="请在这里填写您的电子邮件地址。"></td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">个人照片：</span></div></td>
      <td width="264" height="20" valign="middle"> 
        
        <input id="photo" name="photo" value="images/nopic.gif" onChange="document.all.imag.src=this.value" size="26" class="smallInput" maxlength="255" style="font-family: 宋体; font-size: 9pt" title="个人照片。您可以上传自己的照片，也可以直接填写您的网上照片的地址。"> 
				<select onChange="document.all.imag.src=options[selectedIndex].value;document.all.photo.value=options[selectedIndex].value" >
					<option select value="images/nopic.gif">默认</option> 
					 
					<option select value="images/Image1.gif">头像1</option> 
					 
					<option select value="images/Image2.gif">头像2</option> 
					 
					<option select value="images/Image3.gif">头像3</option> 
					 
					<option select value="images/Image4.gif">头像4</option> 
					 
					<option select value="images/Image5.gif">头像5</option> 
					 
					<option select value="images/Image6.gif">头像6</option> 
					 
					<option select value="images/Image7.gif">头像7</option> 
					 
					<option select value="images/Image8.gif">头像8</option> 
					 
					<option select value="images/Image9.gif">头像9</option> 
					 
					<option select value="images/Image10.gif">头像10</option> 
					 
					<option select value="images/Image11.gif">头像11</option> 
					 
					<option select value="images/Image12.gif">头像12</option> 
					 
					<option select value="images/Image13.gif">头像13</option> 
					 
					<option select value="images/Image14.gif">头像14</option> 
					 
					<option select value="images/Image15.gif">头像15</option> 
					 
					<option select value="images/Image16.gif">头像16</option> 
					 
					<option select value="images/Image17.gif">头像17</option> 
					 
					<option select value="images/Image18.gif">头像18</option> 
					 
					<option select value="images/Image19.gif">头像19</option> 
					 
					<option select value="images/Image20.gif">头像20</option> 
					 
					<option select value="images/Image21.gif">头像21</option> 
					 
					<option select value="images/Image22.gif">头像22</option> 
					 
					<option select value="images/Image23.gif">头像23</option> 
					 
					<option select value="images/Image24.gif">头像24</option> 
					 
					<option select value="images/Image25.gif">头像25</option> 
					 
					<option select value="images/Image26.gif">头像26</option> 
					 
					<option select value="images/Image27.gif">头像27</option> 
					 
					<option select value="images/Image28.gif">头像28</option> 
					 
					<option select value="images/Image29.gif">头像29</option> 
					 
					<option select value="images/Image30.gif">头像30</option> 
					 
					<option select value="images/Image31.gif">头像31</option> 
					 
					<option select value="images/Image32.gif">头像32</option> 
					 
					<option select value="images/Image33.gif">头像33</option> 
					 
					<option select value="images/Image34.gif">头像34</option> 
					 
					<option select value="images/Image35.gif">头像35</option> 
					 
				</select>
        <p align=center><img src="images/nopic.gif" name="imag" border="0" onload="javascript:if(this.width>screen.width-550)this.width=screen.width-550">
        </P>
        
      </td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">自我介绍：</span></div></td>
      <td width="264" height="32" valign="middle"> <textarea rows="5" name="content" cols="29" style="font-family: 宋体; font-size: 9pt" title="请在这里填写您的个人介绍。"></textarea> 
      </td>
    </tr>
  </table>
		<div align="center">
				<br>
					<input type="submit" value=" 确 定 " name="cmdOk" class="buttonface" style="font-family: 宋体; font-size: 9pt;">
					&nbsp; 
					<input type="reset" value=" 重 填 " name="cmdReset" class="buttonface" style="font-family: 宋体; font-size: 9pt;" >
		</div>
	</form>
	</body>
	</html>
