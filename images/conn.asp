<meta http-equiv="Content-Type" content="text/html; charset=utf-8" c/>
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = 65001
%><!-- #include file="../dataname.asp" -->
<%
response.write "<style>td{font-size:9pt}</style>"
m_user=session("m_user")
m_gold=session("m_gold")

if m_user="" or m_gold<"1" then
%>

<boby onload="document.form1.h_name.focus();">
<p>&nbsp;</p>
<table width="221" height="298" border="0" align="center" cellpadding="0" cellspacing="0" background="../images/backlogin.jpg">
  <tr>
    <td height="93" valign="top"> 
<div align="center">
        <table width="450" height="280" border="0" align="center" cellpadding="0" cellspacing="1">
<form name="form1" method="post" action="loginadmin.asp" target="_top">
            <tr> 
              <td height="121" colspan="2">&nbsp;</td>
            </tr>
            <tr> 
              <td height="23">&nbsp;</td>
              <td height="23"><strong><font color="#cc6600">您没有登陆或已自动退出。</font>请登陆：</strong></td>
            </tr>
            <tr> 
              <td width="143" height="34"> 
<div align="right"></div></td>
              <td width="304"> <input name="h_name" type="text" size="20" style="height:23px;width:180px;"> 
              </td>
            </tr>
            <tr> 
              <td height="27" valign="top"><div align="right"></div></td>
              <td> <input name="h_pwd" type="password" size="20" style="height:23px;width:180px;"> 
              </td>
            </tr>
            <tr> 
              <td height="45" colspan="2" valign="bottom"> 
<div align="center"> 
                  <input type="submit" name="submit" value=" 登 陆 " style="height:25px;width:60px;">
                </div></td>
            </tr>
            <tr> 
              <td height="19" colspan="2"> 　　<a href="../index.asp">回到首页</a></td>
            </tr>
          </form>
        </table>
      </div></td>
  </tr>
</table>
</boby>

<%
	response.end
end if

db="../data/"&mydataname
set conn = server.createobject("adodb.connection")
connstr="driver={microsoft access driver (*.mdb)};uid=admin;pwd="&mydatapsw&";dbq=" & server.mappath(""&db&"")
conn.open connstr

sub connclose
	conn.close
	set conn=nothing
end sub
sub history
	response.write " <a href=javascript:history.back()>返回</a> "
end sub

function addspace(strfield)	'列表显示	
	if strfield="" or isnull(strfield) or strfield="未知" then 
		addspace="&nbsp;"
	else
		addspace=strfield
	end if
end function
%>

