
<!--#include file="conn.asp"-->
<!--#include file="check2.asp"-->


<%
okkind=request.form("okkind")
if okkind="ok" then

		wuliuflag=trim(request.form("wuliuflag"&i))
		memo=trim(request.form("memo"&i))
		sql="update sh_wuliuflag set wuliuflag='"&wuliuflag&"' "
		conn.execute(sql)

	response.write "<script language=JavaScript>" & "alert('成功保存设置！');" & "window.location.href='sendkind.asp';" & "</script>"
	conn.close
	response.end
end if

ok=request.form("ok")
if ok="ok" then
	conn.execute("delete * from sh_sendkind")
	
	freetotal=trim(request.form("freetotal"))
	if freetotal="" then freetotal="0"
	freezaiyao=trim(request.form("freezaiyao"))
	
	for i=1 to request.form("maxi")

		sendkind=trim(request.form("sendkind"&i))
		sendmoney=trim(request.form("sendmoney"&i))
		zaiyao=trim(request.form("zaiyao"&i))
		if not isnumeric(sendmoney) then sendmoney="0"
		if sendkind<>"" then
		sql="insert into sh_sendkind(sendkind,sendmoney,zaiyao,freetotal,freezaiyao) values('"&sendkind&"','"&sendmoney&"','"&zaiyao&"','"&freetotal&"','"&freezaiyao&"')"
		conn.execute(sql)
		end if
	next
	response.write "<script language=JavaScript>" & "alert('成功保存设置！');" & "window.location.href='sendkind.asp';" & "</script>"
	response.end
end if
%>
<html>
<head>
<meta http-equiv=Content-Type content=text/html; charset=utf-8>
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
<TABLE width=93% border=0 align="center" cellPadding=0 cellSpacing=0 style="BORDER-RIGHT: #b1bfee 1px solid; BORDER-TOP: #b1bfee 1px solid; BORDER-LEFT: #b1bfee 1px solid; BORDER-BOTTOM: #b1bfee 1px solid" >
  <TBODY>
    <TR>
      <TD width="90%" bgColor=#f4f5fd><font color="#000000">　<B><font color="#FF0000">选择配送计费方式</font></B></font></TD>
      <TD bgColor=#f4f5fd height=30>&nbsp;</TD>
    </TR>
  </TBODY>
</TABLE>

<%
dim sql4,rs4,id,num,pkid
	sql4="select wuliuflag from sh_wuliuflag"
	set rs4=server.createobject("adodb.recordset")
	rs4.open sql4,conn,1,1
	if rs4.bof or rs4.eof then
	else
		wuliuflag=rs4("wuliuflag")
	end if
	rs4.close
	set rs4=nothing
%>

<table width="93%"  border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#b1bfee">
  <form action=sendkind.asp method=post name=setup>
    <tr bgcolor="#FFFFFF"> 
      <td width="10%" height="30" align="center" valign="middle"><div align="right">
        <input type="radio" name="wuliuflag" value="1" <%if wuliuflag="1" then response.write "checked"%>>
      </div></td>
      <td  valign="middle">普通计件方式 （通常零售，商品量少，重量轻，运费差额不大，计费简单的网店可选择此方式）</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="30" align="center" valign="middle"><div align="right">
        <input type="radio" name="wuliuflag" value="2" disabled >
      </div></td>
      <td  valign="middle">按商品重量计算运费方式 （通常做商品批发，每单的商品量大，重量大导致配送运费大，按不用地区、配送公司计费，可选择此方式）</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="30">&nbsp;</td>
      <td>说明：<br>
	  1、可以在决定用哪种方式前，模拟做几张订单，测试一下这两种计算运费方式再决定。<br>
	  2、选择不同的计算运费方面后，在顾客下订单结算时会按所选的方式计算运费。请按具体情况选择。</td>
    </tr>
    <tr> 
      <td height="30" bgcolor="#FFFFFF" >&nbsp;</td>
      <td height="30" bgcolor="#FFFFFF" ><input name="okkind" type="hidden" value="ok"> 
        <input name=action type="submit" value="保存设置">
        &nbsp;</td>
    </tr>
  </form>
</table>
</td></tr>
</table> </td></tr> </table> 




<table height="20"><tr><td></td></tr></table>
<TABLE width=93% border=0 align="center" cellPadding=0 cellSpacing=0 style="BORDER-RIGHT: #b1bfee 1px solid; BORDER-TOP: #b1bfee 1px solid; BORDER-LEFT: #b1bfee 1px solid; BORDER-BOTTOM: #b1bfee 1px solid" >
  <TBODY>
    <TR>
      <TD width="90%" bgColor=#f4f5fd><font color="#000000">　已选择了“普通计件方式”，<B><font color="#FF0000">&nbsp;配送设置(送货方式)</font></B></font></TD>
      <TD bgColor=#f4f5fd height=30>&nbsp;</TD>
    </TR>
  </TBODY>
</TABLE>
		  


<table width="93%"  border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#b1bfee">
  <form action=sendkind.asp method=post name="setup2">
    <%
	i=1
	sql="select * from sh_sendkind order by id"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.bof or rs.eof then
	else
		freetotal=rs("freetotal")
		if freetotal=0 then freetotal=""
		freezaiyao=rs("freezaiyao")

		do while not rs.eof
		id=rs("id")
		sendkind=rs("sendkind")
		sendmoney=rs("sendmoney")
		zaiyao=rs("zaiyao")

%>
    <tr bgcolor="#FFFFFF" onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td width="10%" valign="middle" align="center"><%=i%>&nbsp;</td>
      <td  valign="middle">
        方式：
        <input type="text" name="sendkind<%=i%>" value="<%=sendkind%>" size="30" maxlength="30" style="BACKGROUND-COLOR: #fcfcfc;height:24px;">
        <br>
        运费：
        <input type="text" name="sendmoney<%=i%>" value="<%=sendmoney%>" size="30" maxlength="30">
        <br>
        说明：
        <input type="text" name="zaiyao<%=i%>" value="<%=zaiyao%>" size="30" maxlength="250">
        <br> </td>
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
        方式：
        <input type="text" name="sendkind<%=i%>" value="" size="30" maxlength="30"  style="BACKGROUND-COLOR: #fcfcfc;height:24px;">
        <br>
        运费：
        <input type="text" name="sendmoney<%=i%>" value="" size="30" maxlength="30">
        <br>
        说明：
        <input type="text" name="zaiyao<%=i%>" value="" size="30" maxlength="250">(添加此行，保存后，会自动增加一空白行。)
        <br> </td>
    </tr>
    <tr bgcolor="#F1F1F1"> 
      <td width="10%" valign="middle" align="center">&nbsp;</td>
      <td  valign="middle">
        满<input type="text" name="freetotal" value="<%=freetotal%>" size="6" maxlength="9"  style="BACKGROUND-COLOR: #fcfcfc;height:24px;">元免运费
        <br>
        说明：
        <input type="text" name="freezaiyao" value="<%=freezaiyao%>" size="30" maxlength="250">
        <br> </td>
    </tr>
    <tr>
      <td height="45" bgcolor="#FFFFFF" >&nbsp;</td>
      <td bgcolor="#FFFFFF" ><input name="maxi" type="hidden" value="<%=i%>"> 
        <input name="ok" type="hidden" value="ok"> &nbsp;<input name=action type="submit" value="保存设置"></td>
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
