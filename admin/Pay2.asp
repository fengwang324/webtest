<!--#include file="conn.asp"-->
<!--#include file="check2.asp"-->
<%
ok=request.form("ok")
if ok="ok" then
	for i=1 to request.form("maxid")

		bankno=trim(request.form("bankno"&i))
		bankren=trim(request.form("bankren"&i))
		bankaddr=trim(request.form("bankaddr"&i))

		sql="update sh_pay2 set bankno='"&bankno&"',bankren='"&bankren&"',bankaddr='"&bankaddr&"' where id="&i
		conn.execute(sql)
		
	next
	response.write "<script language=JavaScript>" & "alert('成功保存设置！');" & "window.location.href='pay2.asp';" & "</script>"
	conn.close
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
#tabcss{
	BORDER-COLLAPSE: collapse;
	border-top-width: 1px;
	border-left-width: 1px;
	border-top-style: solid;
	border-left-style: solid;
	border-top-color: #dddddd;
	border-left-color: #dddddd;
}#tabcss td {
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
<TABLE width=93% border=0 align="center" cellPadding=0 cellSpacing=0 style="BORDER-RIGHT: #b1bfee 1px solid; BORDER-TOP: #b1bfee 1px solid; BORDER-LEFT: #b1bfee 1px solid; BORDER-BOTTOM: #b1bfee 1px solid" 
>
  <TBODY>
    <TR>
      <TD width="90%" bgColor=#f4f5fd><font color="#000000">　<B><font color="#FF0000">&nbsp;银行帐号设置</font></B></font></TD>
      <TD bgColor=#f4f5fd height=30>&nbsp;</TD>
    </TR>
  </TBODY>
</TABLE>
		  


<table width="93%"  border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#b1bfee">
  <form action=pay2.asp method=post name=setup>
    <%
	sql="select * from sh_pay2 order by id"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.bof or rs.eof then
	else
		do while not rs.eof
		id=rs("id")
		bankpic=rs("bankpic")
		bankname=rs("bankname")
		bankno=rs("bankno")
		bankren=rs("bankren")
		bankaddr=rs("bankaddr")
		bankmemo=rs("bankmemo")
%>
    <tr bgcolor="#FFFFFF"> 
      <td width=20% valign="middle" align="center"><img border="0" src="../IMAGES/bank/bank<%=id%>.gif">&nbsp;</td>
      <td  valign="middle"><input name="id<%=id%>" type="hidden" value="<%=id%>">
        帐&nbsp;&nbsp;号：
        <input type="text" name="bankno<%=id%>" value="<%=bankno%>" size=40 maxlength=250>
        <br>
        持卡人：
        <input type="text" name="bankren<%=id%>" value="<%=bankren%>" size=40 maxlength=250>
        <br>
        开户行：
        <input type="text" name="bankaddr<%=id%>" value="<%=bankaddr%>" size=40 maxlength=250>
        <br> </td>
    </tr>
    <%
	    rs.movenext
		loop
	end if
	rs.close
	set rs=nothing
%>
    <tr>
      <td height="30" bgcolor="#FFFFFF" >&nbsp;</td>
      <td height="30" bgcolor="#FFFFFF" ><input name="maxid" type="hidden" value="<%=id%>"> 
        <input name="ok" type="hidden" value="ok"> &nbsp;
<input name=action type="submit" value="保存设置">
        （在这里保存后，还在要“栏目管理”-&gt;“支付方式”中进行相应的修改。）</td>
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

<%conn.close%>
</body></html>
