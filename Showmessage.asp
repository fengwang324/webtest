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
<title><%=application("sitename")%></title>
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
-->
</style>
</head>

<body leftmargin="0" topmargin="0">
<!--#include file="head.asp"-->
<TABLE width="960" border=0 align="center" cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
  <tr>
    <td bgcolor="#fafafa"> 

<table width="94%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr> 
          <td width="70%" height="23"><B><%=application("sitename")%>留言簿：</B></td>
          <td width="40%" align=right><B><a href='g_post.asp'>我要留言</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href='message.asp'>返回留言簿</a></B></td>
        </tr>
      </table>
      <table width="95%" height="413" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td valign="top"> 
            <%
dim rs
id=request("id")
set rs=server.createobject("adodb.recordset")
sql="select * from message where g_list="& id &""
rs.open sql,conn,1,1
do while not rs.eof
set g_brow=rs("g_brow")
set g_title=rs("g_title")
set g_man=rs("g_man")
set g_time=rs("g_time")
set g_content=rs("g_content")
set g_ip=rs("g_ip")
set g_city=rs("g_city")
set g_mail=rs("g_mail")
set g_qq=rs("g_qq")
set g_list=rs("g_list")
set id=rs("id")
set check=rs("check")
g_content= replace(g_content,vbcrlf,"<br>")
g_content= replace(g_content," ","&nbsp;")
%>
<table width="99%" border="0" align="center" cellpadding="5" cellspacing="0">
              <tr bgcolor="#9DCEFF"> 
          <td width="551"><strong><font color="#FFFFFF"><%=g_title%></font></strong></td>
          <td width="189">&nbsp;</td>
        </tr>
      </table>
            <table width="99%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolor="#eeeeee">
              <tr> 
          <td width="172" valign="top" bgcolor="#E1F1FF"> 
            <table width="179" height="92" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><img src="pic/icon<%=g_brow%>.gif" width="15" height="15">　<%=g_man%></td>
              </tr>
              <tr> 
                <td><img src="pic/homepage.gif" width="19" height="18"> <%=g_city%></td>
              </tr>
              <tr> 
                <td><img src="pic/message.gif" width="18" height="18"> <%=g_time%> 
                </td>
              </tr>
            </table></td>
          <td width="620" colspan="2" bgcolor="#F4FAFF">
<table width="100%" height="92" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td height="46" style="WIDTH: 600; WORD-WRAP: break-word"><%=g_content%></td>
              </tr>
              <tr> 
                <td height="25"><img src="pic/email.gif" width="18" height="18"> 
                  <%=g_mail%> 　<img src="pic/oicq.gif" width="18" height="18"> 
                  <%=g_qq%> 　 <img src="pic/ip.gif" width="13" height="15"> <%=g_ip%>
				  <%
				   if session("m_gold")="2" or session("m_gold")="4" or session("m_gold")="8" then
					  response.Write"　　　<a href=g_post.asp?re=1&g_title="&g_title&"&id="& id &"><font color=#ff0000>回 复</font></a>"
					  response.Write"　　　<a href=s_delete.asp?all=1&id="& id &" onClick=""return confirm('真的删除留言？')""><font color=#ff0000>删除留言</font></a>"
					  if check="0" then
						response.Write"　　　未审核 <a href=s_check.asp?all=1&check=1&id="& id &" onClick=""return confirm('真的审核通过吗？')""><font color=#ff0000>点击审核</font></a>"
					  else
						response.Write"　　　已审核 <a href=s_check.asp?all=1&check=0&id="& id &" onClick=""return confirm('真的要反审核吗？')""><font color=#ff0000>点击反审</font></a>"
					  end if
				  end if
				  %>
                </td>
              </tr>
            </table></td>
        </tr>
      </table>
	  <%
rs.movenext       
loop
	  %>
	  
	  
<%
id=request("id")
set rs=server.createobject("adodb.recordset")
sql="select * from message where parentid="& id &""
rs.open sql,conn,1,1
do while not rs.eof
set g_brow=rs("g_brow")
set g_title=rs("g_title")
set g_man=rs("g_man")
set g_time=rs("g_time")
set g_content=rs("g_content")
set g_ip=rs("g_ip")
set g_city=rs("g_city")
set g_mail=rs("g_mail")
set g_qq=rs("g_qq")
set g_list=rs("g_list")
set id=rs("id")
g_content= replace(g_content,vbcrlf,"<br>")
g_content= replace(g_content," ","&nbsp;")
%>
<table width="99%" border="0" align="center" cellpadding="5" cellspacing="0">
              <tr bgcolor="#9DCEFF"> 
          <td width="551"><strong><font color="#FF5555">管理员回复： </font><font color="#FFFFFF"><%=g_title%></font></strong></td>
          <td width="189">&nbsp;</td>
        </tr>
      </table>
            <table width="99%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolor="#eeeeee">
              <tr> 
          <td width="172" valign="top" bgcolor="#E1F1FF"> 
            <table width="179" height="92" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><img src="pic/icon<%=g_brow%>.gif" width="15" height="15">　<%=g_man%></td>
              </tr>
              <tr> 
                <td><img src="pic/homepage.gif" width="19" height="18"> <%=g_city%></td>
              </tr>
              <tr> 
                <td><img src="pic/message.gif" width="18" height="18"> <%=g_time%> 
                </td>
              </tr>
            </table></td>
          <td width="620" colspan="2" bgcolor="#F4FAFF">
<table width="100%" height="92" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td height="46"><%=g_content%></td>
              </tr>
              <tr> 
                <td height="25"><img src="pic/email.gif" width="18" height="18"> 
                  <%=g_mail%> 　<img src="pic/oicq.gif" width="18" height="18"> 
                  <%=g_qq%> 　 <img src="pic/ip.gif" width="13" height="15"> <%=g_ip%>
				  <%
				   if session("m_gold")="2" or session("m_gold")="4" or session("m_gold")="8" then
				  response.Write"　　　<a href=s_delete.asp?parentid="&request("id")&"&id="& id &" onClick=""return confirm('真的删除回复？')""><font color=#ff0000>删除回复</font></a>"
				  end if
				  %>
                </td>
              </tr>
            </table></td>
        </tr>
      </table>
	  <%
rs.movenext       
loop
	  %>
	  
	  
    </td>
  </tr>
</table>

</TD></TR></TABLE>
<!-- #include file="foot.asp" -->
<%
conn.close
set conn=nothing
%>
</body>
</html>

