
	<html>
	<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<title>��Աע��</title>
	<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
	<link rel="stylesheet" type="text/css" href="site.css">
	<script LANGUAGE="javascript">
	<!--
	function FrmAddLink_onsubmit() {
	var i, n;
	if (document.FrmAddLink.username.value=="")
		{
		  alert("�Բ��������������û�����")
		  document.FrmAddLink.username.focus()
		  return false
		 }
	else if (document.FrmAddLink.username.value.length < 2)
		{
		  alert("�����û����ܲ��ܳ�һ�㣡")
		  document.FrmAddLink.username.focus()
		  return false
		 }
	else if (document.FrmAddLink.username.value.length > 30)
		{
		  alert("�����û���̫���˰ɣ�")
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
	            alert("�û����а�����Ч�ַ���������ѡ���û�����"); 
	            document.FrmAddLink.username.focus();
	            return false;
	          }
	else if (document.FrmAddLink.passwd.value=="")
		{
		  alert("�Բ��������������룡")
		  document.FrmAddLink.passwd.focus()
		  return false
		 }
	else if (document.FrmAddLink.passwd.value.length < 4)
		{
		  alert("Ϊ�˰�ȫ����������Ӧ�ó�һ�㣡")
		  document.FrmAddLink.passwd.focus()
		  return false
		 }
	else if (document.FrmAddLink.passwd.value.length > 16)
		{
		  alert("��������̫���˰ɣ�")
		  document.FrmAddLink.passwd.focus()
		  return false
		 }
	else if (document.FrmAddLink.username.value==document.FrmAddLink.passwd.value)
		{
		  alert("Ϊ�˰�ȫ���û��������벻Ӧ����ͬ��")
		  document.FrmAddLink.passwd.focus()
		  return false
		 }
	else if (document.FrmAddLink.passwd2.value=="")
		{
		  alert("�Բ�������������֤���룡")
		  document.FrmAddLink.passwd2.focus()
		  return false
		 }
	else if (document.FrmAddLink.passwd2.value !== document.FrmAddLink.passwd.value)
		{
		  alert("�Բ�����������������벻һ�£�")
		  document.FrmAddLink.passwd2.focus()
		  return false
		 }
	else if (document.FrmAddLink.question.value=="")
		{
		  alert("�Բ�������������ʾ���⣡")
		  document.FrmAddLink.question.focus()
		  return false
		 }
	else if (document.FrmAddLink.answer.value=="")
		{
		  alert("�Բ���������������𰸣�")
		  document.FrmAddLink.answer.focus()
		  return false
		 }
	else if (document.FrmAddLink.question.value==document.FrmAddLink.answer.value)
		{
		  alert("Ϊ�˰�ȫ����ʾ����������𰸲�Ӧ����ͬ��")
		  document.FrmAddLink.answer.focus()
		  return false
		 }
	else if (document.FrmAddLink.fullname.value=="")
		{
		  alert("�Բ���������������ʵ������")
		  document.FrmAddLink.fullname.focus()
		  return false
		 }
	else if (document.FrmAddLink.depid.value=="")
		{
		  alert("�Բ�����ѡ�����Ĺ�����λ��")
		  document.FrmAddLink.depid.focus()
		  return false
		 }
	else if (document.FrmAddLink.sex.value=="")
		{
		  alert("�Բ�����ѡ�������Ա�")
		  document.FrmAddLink.sex.focus()
		  return false
		 }
	else if (document.FrmAddLink.tel.value=="")
		{
		  alert("�Բ���������������ϵ�绰��")
		  document.FrmAddLink.tel.focus()
		  return false
		 }
	else if (document.FrmAddLink.email.value=="")
		{
		  alert("�Բ������������ĵ����ʼ���")
		  document.FrmAddLink.email.focus()
		  return false
		 }
	else if (document.FrmAddLink.email.value.indexOf("@",0)== -1||document.FrmAddLink.email.value.indexOf(".",0)==-1)
		  {
		  alert("�Բ���������ĵ����ʼ�����")
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
	
	<//����ѡ�����ڴ���ʼ>
<SCRIPT language=JavaScript src="inc/User_Info_Modify.js"></SCRIPT>
	<//����ѡ�����ڴ������>
	</head>
	<body topmargin="0">    
	<form align="center" method="post" name="FrmAddLink" LANGUAGE="javascript" onsubmit="return FrmAddLink_onsubmit()" action="saveuser.asp">
  <table width="500" border="1" align=center cellpadding="0" cellspacing="0" bordercolor="0" bgcolor="#FFFFFF" style="border-collapse: collapse">
    <tr> 
      <td height="20" colspan="2" align="center" valign="middle"><span class="smallFont">��<strong> �� Ա ע �� </strong>��</span></td>
    </tr>
    <tr> 
      <td height="20" colspan="2" align="center" valign="middle">Ϊ��ʹ��������ʹ�ñ�ϵͳ������ϸ��дÿһ������</td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right" valign="middle"> <div align="center"><span class="smallFont">�� 
          �� ����</span></div></td>
      <td width="264" height="20"> <div align="left"> 
          <input name="username" size="26" class="smallInput" maxlength="30" style="font-family: ����; font-size: 9pt" title="����������д�����û�����ע��ɹ����ô��û�����¼��ϵͳ��">
        </div></td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">��&nbsp;&nbsp;&nbsp;�룺</span></div></td>
      <td width="264" height="20"> <div align="left"> 
          <input name="passwd" size="26" class="smallInput" maxlength="15" style="font-family: ����; font-size: 9pt" type="password"  title="����������д���ĵ�¼���룬�ڵ�¼ʱ��ϵͳ����֤�������롣">
        </div></td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">��֤���룺</span></div></td>
      <td width="264" height="20"> <div align="left"> 
          <input name="passwd2" size="26" class="smallInput" maxlength="15" style="font-family: ����; font-size: 9pt" type="password" title="����������д������֤���룬���������������һ�£���Ҫ�Ƿ�ֹ���Ĵ������롣">
        </div></td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">��ʾ���⣺</span></div></td>
      <td width="264" height="20"> <div align="left"> 
          <input name="question" size="26" class="smallInput" maxlength="50" style="font-family: ����; font-size: 9pt" type="text" title="����������д������ʾ���⣬��������������룬�������ô˹������һ��������롣">
        </div></td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">����𰸣�</span></div></td>
      <td width="264" height="20"> <div align="left"> 
          <input name="answer" size="26" class="smallInput" maxlength="50" style="font-family: ����; font-size: 9pt" type="text" title="����������д������ʾ����Ĵ𰸣���������������룬�������ô˹������һ��������롣">
        </div></td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">��ʵ������</span></div></td>
      <td width="264" height="20"> <div align="left"> 
          <input name="fullname" size="26" class="smallInput" maxlength="10" style="font-family: ����; font-size: 9pt" title="����������д������ʵ������">
        </div></td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">������λ��</span></div></td>
      <td width="264" height="20"> 
        
        <select name="depid" style="font-family: ����; font-size: 9pt" title="��������ѡ�����Ĺ�����λ��">
          <option value="">��ѡ������λ</option>
          
          <option value="1"></option>
          
          <option value="2"></option>
          
          <option value="3"></option>
          
          <option value="4"></option>
          
        </select> </td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">��&nbsp;&nbsp;&nbsp; 
          ��</span></div></td>
      <td width="264" height="20"> <select size="1" name="sex" style="font-family: ����; font-size: 9pt" title="����������д�����Ա�">
          <option selected value="">��ѡ���Ա�</option>
          
          <option value="����">����</option>
          <option value="Ůʿ">Ůʿ</option>
          <option value="����">����</option>
          
        </select> </td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">��&nbsp;&nbsp;&nbsp;�գ�</span></div></td>
      <td width="264" height="40"> 
        
        <div align="left"> 
          <select size="1" name="birthyear" style="font-family: ����; font-size: 9pt">
            
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
          �� 
          <select size="1" name="birthmonth" style="font-family: ����; font-size: 9pt">
            
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
          �� 
          <select size="1" name="birthday" style="font-family: ����; font-size: 9pt">
            
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
          �� </div>
        
      </td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">��ϵ�绰��</span></div></td>
      <td width="264" height="20" valign="middle"> <input name="tel" size="26" class="smallInput" maxlength="100" style="font-family: ����; font-size: 9pt" title="����������д������ϵ�绰���Ա�����������ϵ��"></td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">�����ʼ���</span></div></td>
      <td width="264" height="20" valign="middle"> <input name="email" size="26" class="smallInput" maxlength="100" style="font-family: ����; font-size: 9pt" title="����������д���ĵ����ʼ���ַ��"></td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">������Ƭ��</span></div></td>
      <td width="264" height="20" valign="middle"> 
        
        <input id="photo" name="photo" value="images/nopic.gif" onChange="document.all.imag.src=this.value" size="26" class="smallInput" maxlength="255" style="font-family: ����; font-size: 9pt" title="������Ƭ���������ϴ��Լ�����Ƭ��Ҳ����ֱ����д����������Ƭ�ĵ�ַ��"> 
				<select onChange="document.all.imag.src=options[selectedIndex].value;document.all.photo.value=options[selectedIndex].value" >
					<option select value="images/nopic.gif">Ĭ��</option> 
					 
					<option select value="images/Image1.gif">ͷ��1</option> 
					 
					<option select value="images/Image2.gif">ͷ��2</option> 
					 
					<option select value="images/Image3.gif">ͷ��3</option> 
					 
					<option select value="images/Image4.gif">ͷ��4</option> 
					 
					<option select value="images/Image5.gif">ͷ��5</option> 
					 
					<option select value="images/Image6.gif">ͷ��6</option> 
					 
					<option select value="images/Image7.gif">ͷ��7</option> 
					 
					<option select value="images/Image8.gif">ͷ��8</option> 
					 
					<option select value="images/Image9.gif">ͷ��9</option> 
					 
					<option select value="images/Image10.gif">ͷ��10</option> 
					 
					<option select value="images/Image11.gif">ͷ��11</option> 
					 
					<option select value="images/Image12.gif">ͷ��12</option> 
					 
					<option select value="images/Image13.gif">ͷ��13</option> 
					 
					<option select value="images/Image14.gif">ͷ��14</option> 
					 
					<option select value="images/Image15.gif">ͷ��15</option> 
					 
					<option select value="images/Image16.gif">ͷ��16</option> 
					 
					<option select value="images/Image17.gif">ͷ��17</option> 
					 
					<option select value="images/Image18.gif">ͷ��18</option> 
					 
					<option select value="images/Image19.gif">ͷ��19</option> 
					 
					<option select value="images/Image20.gif">ͷ��20</option> 
					 
					<option select value="images/Image21.gif">ͷ��21</option> 
					 
					<option select value="images/Image22.gif">ͷ��22</option> 
					 
					<option select value="images/Image23.gif">ͷ��23</option> 
					 
					<option select value="images/Image24.gif">ͷ��24</option> 
					 
					<option select value="images/Image25.gif">ͷ��25</option> 
					 
					<option select value="images/Image26.gif">ͷ��26</option> 
					 
					<option select value="images/Image27.gif">ͷ��27</option> 
					 
					<option select value="images/Image28.gif">ͷ��28</option> 
					 
					<option select value="images/Image29.gif">ͷ��29</option> 
					 
					<option select value="images/Image30.gif">ͷ��30</option> 
					 
					<option select value="images/Image31.gif">ͷ��31</option> 
					 
					<option select value="images/Image32.gif">ͷ��32</option> 
					 
					<option select value="images/Image33.gif">ͷ��33</option> 
					 
					<option select value="images/Image34.gif">ͷ��34</option> 
					 
					<option select value="images/Image35.gif">ͷ��35</option> 
					 
				</select>
        <p align=center><img src="images/nopic.gif" name="imag" border="0" onload="javascript:if(this.width>screen.width-550)this.width=screen.width-550">
        </P>
        
      </td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">���ҽ��ܣ�</span></div></td>
      <td width="264" height="32" valign="middle"> <textarea rows="5" name="content" cols="29" style="font-family: ����; font-size: 9pt" title="����������д���ĸ��˽��ܡ�"></textarea> 
      </td>
    </tr>
  </table>
		<div align="center">
				<br>
					<input type="submit" value=" ȷ �� " name="cmdOk" class="buttonface" style="font-family: ����; font-size: 9pt;">
					&nbsp; 
					<input type="reset" value=" �� �� " name="cmdReset" class="buttonface" style="font-family: ����; font-size: 9pt;" >
		</div>
	</form>
	</body>
	</html>
