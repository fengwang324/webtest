<style type="text/css">
<!--
.style1 {font-weight: bold}
.style2 {
	color: #993333;
	font-weight: bold;
}
#font {
	font-family: Arial, Helvetica, sans-serif;
}
#fond {
	font-family: Arial, Helvetica, sans-serif;
}
-->
</style>
<TABLE width="1020" border=0 align="center" cellPadding=0 cellSpacing=0 background="images/foot_1.gif" style="margin-top:10px;">
  <!--DWLayoutTable-->
<TBODY>
    <TR>
	<TD width="10" background=images/foot_b_1.gif></TD>
      <TD width="20">&nbsp;</TD>
      <TD valign="top"> 
        <table border="0" width="98%" cellPadding=0 cellSpacing=0><tr><td valign='top' width='150'>
        <%
		sql4="select l_id,menukind,l_title,zhidingurl,newopen from e_left where showflag='是' and bigmenu='否' order by bigmenu desc,l_num"
		set rs4=server.createobject("adodb.recordset")
		rs4.open sql4,conn,1,1
		if rs4.bof or rs4.eof then
		else
			i=1
			colshu=rs4.recordcount
			
			do while not rs4.eof

				l_id=rs4("l_id")
				menukind=rs4("menukind")
				l_title=rs4("l_title")
				zhidingurl=rs4("zhidingurl")
				newopen=rs4("newopen")
				if i>1 then
					response.write "</td><td width=5><img src='images/foot_3.gif'></td><td valign='top'  width='145'>"
				end if
				response.write "<table cellPadding=0 cellSpacing=0><tr><td height=7></td><tr></table><font class=kindtext>"&l_title&"</font><br>"				
				
				sql3="select * from e_contect where c_parent2="&l_id&" order by c_num,c_id"
				set rs3=server.CreateObject("adodb.recordset")
				rs3.open sql3,conn,1,1
				if rs3.bof or rs3.eof then
				else
					do while not rs3.eof
						set c_id=rs3("c_id")
						set c_title=rs3("c_title")
						set c_addtime=rs3("c_addtime")
						set detail=rs3("detail")
						response.Write "<table cellPadding=0 cellSpacing=0><tr><td height=4></td><tr></table><IMG height=15 src='images/foot_2.gif' width=25 border=0  align='absmiddle' style='line-height:130%;'><a style='line-height:130%;' href='show_foot.asp?c_id="&c_id &"'"
						if newopen="是" then response.write " target='_blank' "
						response.Write "  class=large>"&c_title &"</a><br>"
						rs3.movenext
					loop
				end if
				
			rs4.movenext
  			i=i+1
  			loop			
		end if
		rs4.close
		set rs4=nothing		
		%>
		
	    </td></tr></table>
	  </TD>
	  <%if cstr(colshu)<"6" then%>
      <TD width="55"><a href='javascript:window.scroll(0,0)'><IMG src="images/top.gif" align=absMiddle border=0></a></TD>
	  <%end if%>
	  <TD width="5" background=images/foot_b_2.gif></TD>
    </TR>
  </TBODY>
</table>

 


<TABLE width="1020" border=0 align="center" cellPadding=0 cellSpacing=0>
  <TBODY>
    <TR><td colspan=2 height=10></td></tr>
    <TR> 
      
      <TD width="62%" align="center"> 
        <%
'			sql4="select * from siteright"
'			set rs4=server.createobject("adodb.recordset")
'			rs4.open sql4,conn,1,1
'			if rs4.bof or rs4.eof then
'					copyright=""
'			else
'					copyright=rs4("copyright")
'			end if
'			rs4.close
'			set rs4=nothing
'			response.write copyright
			%>        <div class="large">美國德玉堂 版權所有<span id="font"> COPYRIGHT 2008-2014 SUNSNU.COM ALL RIGHTS RESRVED</span></div>
        <div class="large">電話：<span id="fond">001-626-279-6569, 001-626-780-9038, 001-626-988-7108 傳真：001-626-279-6569  郵箱：sunsnu@sunsnu.com</span></div>
              <div class="large">地址： <span id="fond">10042 Garvey Ave El Monte,  CA 91733 USA </span>  &nbsp; 微信號：<span id="fond">lvySunDeyu</span></div><a target=blank href=tencent://message/?uin=553960013&amp;Site=sunsnu.com&amp;Menu=yes><img border=0 src=images/QQonline1.gif?p=1:553960013:7 alt=点这里联系我 /></a> <script type="text/javascript">
var _bdhmProtocol = (("https:" == document.location.protocol) ? " https://" : " http://");
document.write(unescape("%3Cscript src='" + _bdhmProtocol + "hm.baidu.com/h.js%3F929e26850b18b448c27b0c1f723d48ea' type='text/javascript'%3E%3C/script%3E"));
</script>

        <br /></TD>
    </TR>
	<TR><td colspan=2 height=20></td></tr>
  </TBODY>
</TABLE>
<script>
  (function(i,s,o,g,r,a,m){i['GoogleAnalyticsObject']=r;i[r]=i[r]||function(){
  (i[r].q=i[r].q||[]).push(arguments)},i[r].l=1*new Date();a=s.createElement(o),
  m=s.getElementsByTagName(o)[0];a.async=1;a.src=g;m.parentNode.insertBefore(a,m)
  })(window,document,'script','//www.google-analytics.com/analytics.js','ga');

  ga('create', 'UA-50304145-1', 'sunsnu.com');
  ga('send', 'pageview');

</script>