	  <select name="action" style="border:1 solid #CCD5F9; font-size:9pt; background-color:#FEEBE4"　size="1">
      <option selected value="">全部内容</option>
      <option value="title">按标题</option>
      <option value="content">按内容</option>
	  <option value="editor">按作者</option>
	  <option value="about">按关键字</option>
	  </select>
      <input type="text" name="keyword" style="font-size:9pt; background-color:#FEEBE4" size=10 value="关键字"onfocus="if (value =='关键字'){value =''}"onblur="if (value ==''){value='关键字'}" maxlength="50">
      <input type="submit" name="Submit" value="搜索"  >