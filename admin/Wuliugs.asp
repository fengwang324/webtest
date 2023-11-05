<!--#include file="conn.asp"-->
<!--#include file="check2.asp"-->
<%
ok=request.form("ok")
if ok="ok" then
	

	
	for i=1 to request.form("maxi")
		id=trim(request.form("id"&i))
		wlcomname=trim(request.form("wlcomname"&i))

		if id="" and wlcomname<>"" then
		sql="insert into sh_wuliugs(wlcomname) values('"&wlcomname&"')"
		conn.execute(sql)
		elseif id<>"" and wlcomname="" then 
		sql="delete * from sh_wuliugs where id="&id
		conn.execute(sql)
		elseif id<>"" and wlcomname<>"" then 
		sql="update sh_wuliugs set wlcomname='"&wlcomname&"' where id="&id
		conn.execute(sql)
		else
		end if
	next
	response.write "<script language=JavaScript>" & "alert('成功保存设置！');" & "window.location.href='wuliugs.asp';" & "</script>"
	response.end
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<style>
<!--
td{font-size:9pt}
body {
	margin-top: 5px;margin-left: 2px;
}
.style3 {color: #A84E22; font-weight: bold; }
#admin{
	BORDER-COLLAPSE: collapse;
	border-top-width: 1px;
	border-left-width: 1px;
	border-top-style: solid;
	border-left-style: solid;
	border-top-color: #dddddd;
	border-left-color: #dddddd;
}#admin td {
	line-height: 24px;
	padding-left: 10px;
	padding-top: 2px;
	padding-bottom: 2px;
	color: #333333;
	border-right-width: 1px;
	border-bottom-width: 1px;
	border-right-style: solid;
	border-bottom-style: solid;
	border-right-color: #dddddd;
	border-bottom-color: #dddddd;
}
-->
</style>
</head>
<BODY bgcolor="#fcfcfc">

<table height="5"><tr><td></td></tr></table>


<table height="5"><tr><td></td></tr></table>
<TABLE width=93% border=0 align="center" cellPadding=0 cellSpacing=0 style="BORDER-RIGHT: #b1bfee 1px solid; BORDER-TOP: #b1bfee 1px solid; BORDER-LEFT: #b1bfee 1px solid; BORDER-BOTTOM: #b1bfee 1px solid" >
  <TBODY>
    <TR>
      <TD width="90%" bgColor=#f4f5fd><font color="#000000">　<B><font color="#FF0000">&nbsp;物流公司</font></B></font></TD>
      <TD bgColor=#f4f5fd height=30>&nbsp;</TD>
    </TR>
  </TBODY>
</TABLE>
		  


<table width="93%"  border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#b1bfee">
  <form action=wuliugs.asp method=post name=setup>
    <%
	i=1
	sql="select * from sh_wuliugs order by id"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.bof or rs.eof then
	else

		do while not rs.eof
		id=rs("id")
		wlcomname=rs("wlcomname")

%>
    <tr bgcolor="#FFFFFF" onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td width="10%" valign="middle" align="center"><%=i%>&nbsp;</td>
      <td  valign="middle">
        <input type="hidden" name="id<%=i%>" value="<%=id%>">
        <input type="text" name="wlcomname<%=i%>" value="<%=wlcomname%>" size="30" maxlength="30" style="BACKGROUND-COLOR: #fcfcfc;height:24px;"></td>
    </tr>
    <%
	    rs.movenext
		i=i+1
		loop
	end if
	rs.close
	set rs=nothing
%>
    <tr bgcolor="#FFFFFF" onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td width="10%" valign="middle" align="center"><%=i%>&nbsp;</td>
      <td  valign="middle">
        <input type="hidden" name="id<%=i%>" value="">
        <input type="text" name="wlcomname<%=i%>" value="" size="30" maxlength="30"  style="BACKGROUND-COLOR: #fcfcfc;height:24px;">(添加此行后，保存，会自动增加一空白行。)</td>
    </tr>
    <tr>
      <td height="45" bgcolor="#FFFFFF" >&nbsp;</td>
      <td bgcolor="#FFFFFF" ><input name="maxi" type="hidden" value="<%=i%>"> 
        <input name="ok" type="hidden" value="ok"> &nbsp;<input name=action type="submit" value="保存设置">
        </td>
    </tr>
  </form>
</table>
</td></tr>
</table> </td></tr> </table> 
         
<table width="95%" border="0">
  <tr> 
    <td height="60">&nbsp;</td>
  </tr>
</table>


</body></html>
