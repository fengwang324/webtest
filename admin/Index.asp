<%if session("m_gold")="" then
response.Redirect "login.asp"
end if
%>
<html>
<head>
<title>美国德玉堂管理中心</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />

<Script language="javascript">
window.onbeforeunload = function()
{
    var n = window.event.screenX - window.screenLeft;
    var b = n > document.documentElement.scrollWidth-20;
	var number=0;

    if (b && window.event.clientY < 0 || window.event.altKey)
    {   
        if(number<15){
            window.event.returnValue = "如果此时关掉窗口，将视为自动放弃！"; 
		}
    }
}
</Script>

</head>
<frameset rows="64,*"  frameborder="NO" border="0" framespacing="0">
	<frame src="admin_top.asp" noresize="noresize" frameborder="NO" name="topFrame" scrolling="no" marginwidth="0" marginheight="0" target="main" />
  <frameset cols="200,*"  rows="599" id="frame">
	<frame src="left.asp" name="leftFrame" noresize="noresize" marginwidth="0" marginheight="0" frameborder="0" scrolling="no" target="main" />
	<frame src="right.asp" name="main" marginwidth="0" marginheight="0" frameborder="0" scrolling="auto" target="_self" />
  <!--frame src="UntitledFrame-1"><frame src="UntitledFrame-2"--></frameset>
<noframes>
<body></body>
    </noframes>
</html>
