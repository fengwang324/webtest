<%
pkid=request.QueryString("pkid")
allpic=request.QueryString("allpic")
productname=request.QueryString("productname")

if allpic<>"" then
	allpic=replace(allpic,"||","|")
	allpic=replace(allpic,"||","|")
	allpic=replace(allpic,"||","|")
else
	allpic="zoom.jpg"
end if

arr=split(allpic,"|")

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="imagetoolbar" content="no" />
<title><%=application("sitename")%> - <%=productname%></title>
<link rel="stylesheet" href="images/zoomstyle.css" />
<style>
#productname {
	height:32px;
	margin: 10px auto;
	padding-top: 0px;
	padding-bottom:10px; 
	padding-right: 95px;
	background: #FFF;
	text-align: center;
	color:#808080; 
	font-size:14px;
	font-weight:bold;
	background-image:url(images/line.gif);
	background-attachment:fixed;
	background-repeat:no-repeat;
	background-position:bottom;
}

#gallery_nav { float: left; width: 120px; text-align: center; }
#gallery_nav img { border: 1px solid #ddd; padding:1px;}

#gallery_output { float: left; width: 745px; height: 700px; overflow: hidden; }
#gallery_output img { display: block; margin: 1px auto 0 auto; border: 2px solid #CCC; padding:1px;}

</style>
<script type="text/javascript" src="images/zoomjquery.min.js"></script>
<script language="javascript">
	
		$(document).ready(function() {
			$("#gallery_output img").not(":first").hide();

			$("#gallery a").click(function() {
				if ( $("#" + this.rel).is(":hidden") ) {
					$("#gallery_output img").slideUp();
					$("#" + this.rel).slideDown();
				}
			});
		});

	</script>
</head>
<body>
<div id="content">
  <div id="productname"><%=productname%>&nbsp;</div>
  <div id="gallery">
    <div id="gallery_output"> 
		<%
		for j=0 to Ubound(arr)
			if arr(j)<>"" then
				response.write "<img id='img"&j&"' src='product/"&arr(j)&"' /> "&vbcrlf
			end if
		next
		%>
	</div>
    <div id="gallery_nav">
		<a href="show.asp?pkid=<%=pkid%>&num=1&tk=shop7z"><img src="images/buy.gif" style="border:0;margin-bottom:12px;margin-top:1px;"></a> 
		<%
		for i=0 to Ubound(arr)
			if arr(i)<>"" then
				'alpha的：response.write "<a rel='img"&i&"' href='javascript:;'><img src='product/"&Server.URLEncode(arr(i))&"' width='100' style='FILTER: alpha(opacity=70);' Onmouseover='this.filters.alpha.opacity=100' Onmouseout='this.filters.alpha.opacity=70' /></a> "&vbcrlf       
				response.write "<a rel='img"&i&"' href='javascript:;'><img src='product/"&arr(i)&"' width='100' Onmouseover=""this.style.border='1px solid #0000ff'"" Onmouseout=""this.style.border='1px solid #ddd'"" /></a> "&vbcrlf       
			end if
		next
		%>
		<p style="margin-top: 2px;"><a href="javascript:window.close()">[关闭]</a> </p>
	</div>

    <div class="clear"></div>
  </div>
</div>



</body>
</html>