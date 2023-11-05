<!--#include file="conn.asp"-->
<%
if s_head="head4.asp" or s_productkind="4" then
	response.write "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">"
else
	response.write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">"
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=application("sitename")%> - 视频展示</title>
<link href="i.css" type=text/css rel=stylesheet>
<style type="text/css">
<!--

.style1 {font-size: 14px}
.style11 {
	font-size: 14px;
	color: #990000;
	font-weight: bold;
}
.style12 {color: #990000}
.style13 {color: #333333}
.style15 {font-size: 14px; color: #669933; font-weight: bold; }
-->
</style></head>
<link href="style.css" rel="stylesheet" type="text/css" />
<body>
<!-- #include file="head.asp" -->
<TABLE width="1020" border=0 align="center" cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
  <tr>
    <td> 

    <table width="1020" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr>
          <td width="210" valign="top">
		  <!-- #include file="inc_left_all.asp" -->
		  </td>
          <td width="750" valign="top">
	
            <TABLE class=text width="98%" align=center cellSpacing=0 cellPadding=0 border=0>
              <TBODY>
                <TR> 
                  <TD vAlign=bottom height=32><IMG height=19 
            src="images/dian04.gif" width=10 align=absMiddle> <STRONG>&nbsp;<FONT color=#ff6600 style='font-size:14px;'>视频展示</FONT></STRONG></TD>
                </TR>
                <TR> 
                  <TD vAlign=top height=25> <DIV align=left><IMG height=25 src="images/line.gif" width=457></DIV></TD>
                </TR>
                <TR> 
                  <TD height="320" vAlign=top>
<br>
<table width="90%" border="0" align="center">

<%if request("act")=1 or request("act")=""   then%>
  <tr>
    <td colspan="2">
  <div><iframe width="640" height="480" src="//www.youtube.com/embed/XU8XfsE_6i4" frameborder="0" allowfullscreen></iframe></div></td>
    </tr>
  <tr>
    <td width="50%" class="nowprice">美国德玉堂灵草百消丹（一）</td>
    <td width="50%"><a href="show.asp?pkid=4850" target="_blank">产品详情&gt;&gt;&gt;</a></td>
  </tr>
  
<%elseif request("act")=2 then%>
  <tr>
    <td colspan="2"><div>
      <iframe width="640" height="480" src="//www.youtube.com/embed/Nf0FanxTNpI" frameborder="0" allowfullscreen></iframe>
    </div></td>
  </tr>
  <tr>
    <td class="nowprice">美国德玉堂灵草百消丹（二）</td>
    <td><a href="show.asp?pkid=4850" target="_blank">产品详情&gt;&gt;&gt;</a></td>
  </tr>
  
<%elseif request("act")=3 then%>
  
  <tr>
    <td colspan="2"><div>
      <iframe width="640" height="480" src="//www.youtube.com/embed/1_Tk42NTdbk" frameborder="0" allowfullscreen></iframe>
    </div></td>
  </tr>
  <tr>
    <td class="nowprice">美国德玉堂<a href="show.asp?pkid=4848" target="_blank">虫草补肾丸</a>、<a href="show.asp?pkid=4826" target="_blank">胶原蛋白粉</a></td>
    <td><a href="show.asp?pkid=4826" target="_blank">产品详情&gt;&gt;&gt;</a></td>
  </tr>
  
<%elseif request("act")=4 then%>
  
  <tr>
    <td colspan="2"><div>
      <iframe width="640" height="480" src="http://www.youtube.com/embed/6mF3RtnnpyA" frameborder="0" allowfullscreen></iframe>
    </div></td>
  </tr>
  <tr>
    <td class="nowprice">美国德玉堂介绍</td>
  </tr>
<%end if%>
  
  
  <tr>
    <td height="10" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td height="30" colspan="2"><strong class="kindtext">视频列表</strong></td>
  </tr>
  <tr>
    <td height="24" colspan="2"> &nbsp; <a href="?act=1">美国德玉堂灵草百消丹（一）</a></td>
    </tr>
  <tr>
    <td height="24" colspan="2"> &nbsp; <a href="?act=2">美国德玉堂灵草百消丹（二）</a></td>
    </tr>
  <tr>
    <td height="24" colspan="2"> &nbsp; <a href="?act=3">美国德玉堂虫草补肾丸、胶原蛋白粉</a></td>
  </tr>
  <tr>
    <td height="24" colspan="2">&nbsp; <a href="?act=4">美国德玉堂介绍</a></td>
  </tr>
</table></TD>
                </TR>
                <TR> 
                  <TD height="25">&nbsp;</TD>
                </TR>
              </TBODY>
            </TABLE>
            
          </td>
  </tr>
</table>


    </td>
  </tr>

<TR><TD>

</TD></TR>
</table>
<!-- #include file="foot.asp" -->
</body>
</html>



