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
.style15 {font-size: 14px; color: #669933; font-weight: bold; }
.helptitle {font-size: 14px; color: #669933; font-weight: bold; background-color:#f7f7f7; }
-->
</style></head>

<body>
<!-- #include file="head.asp" -->
<table width="1020" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="10"></td>
  </tr>
</table>
<TABLE width="1020" border=0 align="center" cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
  <tr>
    <td> 



<table width="1020" height="352" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
	      <td width="230" height="184" valign="top">

			<table width="210" height=32 border="0" cellspacing="0" cellpadding="0" background="images/helpCenter-title.gif"><tr><td style="font-size: 14px; line-height: 130%; font-weight:bold; color:#666666;">&nbsp;帮助中心</td></tr></table>
			
			<table width="210" border="0" cellspacing="0" cellpadding="0" class="weitab">
		  <%
			sql4="select * from view_foot "
			set rs4=server.createobject("adodb.recordset")
			rs4.open sql4,conn,1,1
			if rs4.bof or rs4.eof then
			else
				i=1
				l_titleold=""
				do while not rs4.eof
					l_id=rs4("l_id")
					l_title=rs4("l_title")
					c_id=rs4("c_id")
					c_title=rs4("c_title")
					if i=1 or l_title<>l_titleold then
					response.write "<tr><td height=10></td></tr><tr><td height='27' class='helptitle'>&nbsp;<b>"&l_title&"</b></td></tr>"
					i=1
					end if
		  %>
					<tr>
					<td height="22"><%response.write "&nbsp;&nbsp;<IMG width=4 height=6 src='images/arrow03.gif' align=absMiddle> <a style='line-height:180%;' href=show_foot.asp?pkid="&pkid&"&c_id="&c_id &">"&c_title &"</a>" %></td>
					</tr>
		  <%
		  			l_titleold=l_title			
					rs4.movenext
					i=i+1
				loop			
			end if
			rs4.close
			set rs4=nothing
		  %>
		  <tr><td height=20></td></tr>
		</table>
		
	      </td>
    <td valign="top">
		<table width="99%"  border="0" cellspacing="0" cellpadding="0">
		  <tr>
			<td height=2></td>
		  </tr>
		</table>
<%


sql3="select * from e_contect where c_id="&request("c_id")
set rs3=server.CreateObject("adodb.recordset")
rs3.open sql3,conn,1,1

if rs3.bof or rs3.eof then

else
		set c_id=rs3("c_id")
		set c_title=rs3("c_title")
		set c_contect=rs3("c_contect")
		set c_addtime=rs3("c_addtime")

%>
            <TABLE class=text cellSpacing=0 cellPadding=0 width="100%" align=center 
      border=0>
              <TBODY>
                <TR> 
                  <TD width="91%" height=32 vAlign=bottom><IMG height=19 
            src="images/dian04.gif" width=10 align=absMiddle> <STRONG>&nbsp;<FONT 
            color=#ff6600 style='font-size:14px;'><%=c_title%></FONT></STRONG></TD>
                  <TD width="9%" vAlign=bottom>&nbsp;</TD>
                </TR>
                <TR> 
                  <TD height=25 colspan="2" vAlign=top> <DIV align=left><IMG height=25 src="images/line.gif" 
            width=592></DIV></TD>
                </TR>
                <TR> 
                  <TD height="138" colspan="2" vAlign=top> 
                    <%

	response.write "<table border=0  cellpadding='0' cellspacing='0' width='670' align=center>"
	response.Write "<tr><td style='font-size:12px; line-height:180%; '>"&c_contect &"</td></tr>"
	response.write "</table>"

%>
                  </TD>
                </TR>
                <TR> 
                  <TD height="35" colspan="2" align=right>&nbsp;</TD>
                </TR>
              </TBODY>
            </TABLE>
<%
end if
rs3.close
set rs3=nothing

%>	
			
			
			
			
			
			
            <table width="50%" height="2"  border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td></td>
        </tr>
      </table>
          </td>
  </tr>
</table>


    </td>
  </tr>

<TR><TD>

</TD></TR>
</table>
<!-- #include file="foot.asp" -->
<%
conn.close
set conn=nothing
%>
</body>
</html>



