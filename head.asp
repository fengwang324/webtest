<meta name="google-translate-customization" content="be6c1c0ea43c5fe2-22aedaa640b13498-gd3aa91603a37bec1-e"></meta>
<META http-equiv=Content-Type content="text/html; charset=utf-8">
<link rel="shortcut icon" href="favicon.ico">
<style type="text/css">
.menucontainer ul#topnav a.products {
/*	background: url(images/nav_products.png) no-repeat;
*/	width: <%=kindwidth%>px;    /*108 菜单宽度*/
	height:28px;
	padding-top:8px;
	padding-left:3px;
}
</style>
<link href="headlist.css" type=text/css rel=stylesheet>
<center class="cen">
 <table  width="100%" border="0" align="center">
   <tr>
     <td width="100%" height="30" align="center" bgcolor="#F2F2F2"><table  width="1020" border="0" align="center" cellpadding="0" cellspacing="0">
       <tr>
    <td width="215"><a href="http://www.sunsnu.com/" target="_blank">美國德玉堂中草藥 www.Sunsnu.com</a></td>
    <td width="197" align="center"><div id="google_translate_element"></div><script type="text/javascript">
function googleTranslateElementInit() {
  new google.translate.TranslateElement({pageLanguage: 'zh-CN', layout: google.translate.TranslateElement.InlineLayout.SIMPLE, multilanguagePage: true}, 'google_translate_element');
}
</script><script type="text/javascript" src="//translate.google.com/translate_a/element.js?cb=googleTranslateElementInit"></script></td>
    <td width="190" align="center"><a href="/old/" target="_blank">訪問舊版</a></td>
    <td width="418" align="right"><%if session("username")="" or session("s_stat")="" then%>
您好！欢迎光临<%=application("sitename")%>！[ <b><a href="hyzq.asp"><font color="#FF6600">登錄</font></a></b> / <b><a href="regedit.asp"><font color="#FF6600">註冊</font></a></b> ]
<%else
response.write session("username") & " 您好！ "& session("customkindname")&" | <a href='hyzq.asp'>會員專區</a> | <a href='quit.asp'>安全退出</a>"
			end if%><!-- 
&nbsp; <a href="old/" target="_blank">访问旧版</a>--></td>
  </tr>
</table></td>
    </tr>
  </table>
 <table  width="1020" border="0" align="center" cellpadding="0" cellspacing="0" background="images/headbg.jpg">
   <tr>
     <td width="372" height="120" rowspan="2"><a href="index.asp"><img  src="images/logo.jpg" border="0" align="left" /></a></td>
     <td width="638" height="24" colspan="2" align="right"><a href="order_all.asp?lookorder=1"><img  src="images/gwc.JPG" border="0" class='headimg' /></a> <a href="hyzq.asp"><img src="images/zh.JPG" border="0"  class='headimg' /></a> <a href="show_foot.asp?pkid=2257&amp;c_id=295"><img src="images/zxkf.JPG"  border="0" class='headimg' /></a></td>
   </tr>
   <tr>
     <td colspan="2" align="right">&nbsp;</td>
   </tr>
  </table>
 <table width="100%" border="0" cellpadding="0" cellspacing="0" background="images/menu.gif" >
   <tr>
     <td align="center"><table align="center" width="1020" cellspacing="0" cellpadding="0" border="0">
       <tr>
         <td width="90"  height="40" align="center"><a href="index.asp" class="products0">主 頁</a></td>
         <td width="102" align="center"><a href="productlist.asp" class="products0">產品展示</a></td>
         <td width="102" align="center"><a href="show_foot.asp?c_id=315" class="products0">全球招商</a></td>
         <td width="102" align="center"><a href="zj.asp" class="products0">專家介紹</a></td>
         <td width="102" align="center"><a href="vd.asp" class="products0">視頻展示 </a></td>
         <td width="102" align="center"><a href="jg.asp" class="products0">保健品加工</a></td>
         <td width="102" align="center"><a href="jz.asp?c_id=282" class="products0">客戶見證</a></td>
         <td width="102" align="center"><a href="yn.asp" class="products0">疑惑解答</a></td>
         <td width="102" align="center"><a href="book.asp" class="products0">客戶留言</a></td>
         <td width="102" align="center"><a href="news.asp?l_id=30" class="products0">資訊中心</a></td>
         <td width="102" align="center"><a href="show_foot.asp?c_id=282" class="products0">公司介紹</a></td>
         <td width="102" align="center"><a href="show_foot.asp?c_id=289" class="products0">聯繫我們</a></td>
       </tr>
     </table></td>
   </tr>
 </table>
 <table width="1020" border="0" align="center" cellpadding="0" cellspacing="0">
   <tr>
     
       <form action="productlist.asp" method="post" name="form1" id="form1" ><td width="400" height="32" align="left" bgcolor="#F2F2F2"><div style=" margin:4px 0px 0px 0px; width:660px;height:25px; background:url(images/searchnew.jpg) no-repeat; padding-left:30px;padding-top:2px;padding-bottom:5px;">
			<input name='keyword' style="WIDTH: 247px; HEIGHT: 17px;border:0px; font-size:12px; background-color:#FFFFFF; " maxlength=40 value="" ><BUTTON style="width:70px;HEIGHT: 21px;border:0px;background: url(images/searchnew.jpg) no-repeat 0px 50px;CURSOR:pointer;" type=submit > </BUTTON><BUTTON style="width:70px;HEIGHT: 21px;border:0px;background: url(images/searchnew.jpg) no-repeat 0px 50px;CURSOR:pointer;" type=BUTTON onclick="window.location.href='advsearch.asp'"> </BUTTON>
			<%'response.write "&nbsp;"&request.cookies("sitekeyword")%>
		</div></td></form>
       <td align="right" bgcolor="#F2F2F2"><a href='productlist.asp?kind=00020001'>美容養顏</a>　<a href='productlist.asp?kind=00020002'>護膚保養</a>　<!--<a href='productlist.asp?kind=00020006'>美白產品</a>　--><a href='productlist.asp?kind=00010007'>抗腫瘤</a>　<a href='productlist.asp?kind=00010008'>抗衰老</a>　<a href='productlist.asp?kind=00010009'>強身健體</a> &nbsp; </td>
     
   </tr>
 </table>
</center>
