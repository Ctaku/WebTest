<%
Const MaxFileSize="1000000"        '上传文件大小限制
Const SaveUpFilesPath="UploadFile"        		'存放上传文件的目录
Const UpFileType="doc|ppt|xls|rar|zip|mdb|mp3|pdf|gsp|DWF|txt|wmv|exe"        '允许的上传文件类型
Const byteType="操|法轮功|做|垃圾|狗屎|日|婊子|妈的|贱"        '留言本过滤词语
Const byteipType=""        '网站恶意ＩＰ地址留言屏蔽
Const bytezfType=""        '网站恶意广告留言字符屏蔽
Const newsipType=""        '浏览文章IP限制设置(填写为允许、发文需要相应设置)
Const echuangmenu="&nbsp;&nbsp;&nbsp;&nbsp;<a href='index.asp' class=top1>  首 页</a>&nbsp;   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;     <a href='E_Special.asp' class=top> 专题报道 </a>   &nbsp;&nbsp;&nbsp;&nbsp;   <a href='E_PhotoNews.asp' class=top target='_blank'> 图片新闻 </a>&nbsp; &nbsp;&nbsp;"        '自定义菜单
Const BOTTOMmenu="<a href='E_Special.asp' class=bottom> 专题报道 </a>|<a href='E_PhotoNews.asp' class=bottom target='_blank'> 图片新闻 </a>|<a href='./' class=bottom>&nbsp;网站地图&nbsp;</a>|<a href='mailto:@cbrc.gov.cn' class=bottom> 联系我们 </a>|"	'自定义底部菜单
%>