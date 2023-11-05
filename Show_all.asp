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
-->
</style></head>

<body>
<!-- #include file="head.asp" -->
<TABLE width="1020" border=0 align="center" cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
  <tr>
    <td> 



<table width="1020" height="352" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
    <td width="210" height="184" valign="top">
		<!-- #include file="inc_left_all.asp" -->
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
            <TABLE class=text cellSpacing=0 cellPadding=0 width="99%" align=center 
      border=0>
              <TBODY>
                <TR> 
                  <TD width="91%" height=32 vAlign=bottom><IMG height=19 
            src="images/dian04.gif" width=10 align=absMiddle> <STRONG>&nbsp;<FONT 
            color=#ff6600 style='font-size:14px;'><%=c_title%></FONT></STRONG></TD>
                  <TD width="9%" vAlign=bottom><div align="center">
                      <input type="button" name="Button" value="返回" onClick="javascript:window.history.back()">
                    </div></TD>
                </TR>
                <TR> 
                  <TD height=25 colspan="2" vAlign=top> <DIV align=left><IMG height=25 src="images/line.gif" 
            width=592></DIV></TD>
                </TR>
                <TR> 
                  <TD height="138" colspan="2" vAlign=top> 
                    <%

	response.write "<table border=0  cellpadding='0' cellspacing='0' width='670' align=center>"
	response.Write "<tr><td style='font-size:14px;'>"&c_contect &"</td></tr>"
	response.write "</table>"

%>
                  </TD>
                </TR>
                <TR> 
                  <TD height="35" colspan="2" align=right>发布日期：<%=c_addtime%></TD>
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



