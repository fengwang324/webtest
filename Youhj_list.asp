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
.style12 {color: #990000}
-->
</style>

</head>

<body>
<!-- #include file="head.asp" -->
<TABLE width="1020" border=0 align="center" cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
  <tr>
    <td> 

<TABLE height=278 cellSpacing=0 cellPadding=0 width=1020 align=center bgColor=#ffffff border=0>
        <TBODY>
          <TR> 
            <TD vAlign=top width="180"><!-- #include file="zq_left.asp" --></TD>
            <TD width=780 valign="top" bgColor=#ffffff>
			
<%
if session("username")="" or session("s_stat")="" then
	response.Redirect("QuicklyCheck.asp")
%>	

              <%else%>
					<br>
					<table width="98%" border="0" align="center" cellpadding="1" cellspacing="0"><tr>
                  <td><b><font color="#330099">优惠券对帐：</font></b></td>
</tr></table>
					<table width="98%" border="1" align="center" cellpadding="1" cellspacing="0" bordercolorlight=#cccccc bordercolordark=#ffffff>
					<form name="theForm" method="post" action="chongzhidel.asp?menuid=<%=menuid%>">
					  <tr bgcolor="#cccccc"> 
						<td><div align="center">获得日期</div></td>
						<td><div align="center">相关订单号</div></td>
						<td><div align="center">用户名</div></td>
						<td><div align="center">优惠券名称</div></td>
						<td><div align="center">优惠券金额</div></td>
						<td><div align="center">使用开始日期</div></td>
						<td><div align="center">使用结束日期</div></td>
						<td><div align="center">已使用</div></td>
						<td><div align="center">使用的日期</div></td>
					  </tr>
					  <%
					function mydot(shu)	'显示0
						if isnumeric(shu) then
							shu=cdbl(shu)
							if shu<0 and abs(shu)<1 then
								mydot="-0"&formatnumber(abs(shu),2)
							elseif shu>0 and shu<1 then
								mydot="0"&formatnumber(abs(shu),2)
							elseif shu=0 then
								mydot="0"&formatnumber(shu,2)
							else
								mydot=formatnumber(shu,2)
							end if
						else
							mydot="-"
						end if
					end function
					

					sql="select * from view_bill_youhj where username='"&session("username")&"' order by youhj_getdate desc "
					
					sum_sql="select sum(yhj_dk_jine) as sumyhj_dk_jine from view_bill_youhj where username='"&session("username")&"'"
					
					set rs=server.createobject("adodb.recordset")
					rs.open sql,conn,1,1
					if rs.bof or rs.eof then
						response.write "<tr><td colspan='9'><div align=center>当前没有符合条件记录。"
						response.write "</div></td></tr>"
					else
						session("righturl")="1" 
					   rs.pagesize=20
					   rs.absolutepage=1
					   if request("page")<>"" then rs.absolutepage=request("page")
					   rowcount =rs.pagesize
					   if not rs.eof then
							do while not rs.eof and rowcount>0 
								
							youhj_getdate=rs("youhj_getdate")
							id=rs("id")   '获得的订单号
							billno=rs("billno")
							username=rs("username")
							yhj_name=rs("yhj_name")
							yhj_memo=rs("yhj_memo")
							yhj_mz_jine=rs("yhj_mz_jine")
							yhj_dk_jine=rs("yhj_dk_jine")
							yhj_sy_time=rs("yhj_sy_time")
							yhj_zs_time=rs("yhj_zs_time")
							youhj_used=rs("youhj_used")
							youhj_useddate=rs("youhj_useddate")
							youhj_used_billid=rs("youhj_used_billid")
				
				
							response.write "<tr Onmouseover='return SetBgColor(this);' Onmouseout='return RestoreBgColor(this);'>"
							response.write "<td height=24>"&youhj_getdate&"</td>"
							response.write "<td><a style='CURSOR:hand;' onclick=""window.open('order_print.asp?id="&id&"&billno="&billno&"','','top=55,left=130,Toolbar=yes,scrollbars=yes,width=680,height=500,resizable=yes,menubar=yes,resizable=yes')""><font color=#0033FF>"&billno&"</font></a></td>"
							response.write "<td >"&username&"</td>"
							response.write "<td >"&yhj_name&"</td>"
							response.write "<td align=right>"&mydot(yhj_dk_jine)&"</td>"
							response.write "<td >"&yhj_sy_time&"</td>"
							response.write "<td >"&yhj_zs_time&"</td>"
							response.write "<td >"&youhj_used&"&nbsp;</td>"
							response.write "<td ><a style='CURSOR:hand;' onclick=""window.open('order_print.asp?id="&youhj_used_billid&"&billno=','','top=55,left=130,Toolbar=yes,scrollbars=yes,width=680,height=500,resizable=yes,menubar=yes,resizable=yes')""><font color=#0033FF>"&youhj_useddate&"</font></a>&nbsp;</td>"		
							response.write "</tr>"&vbcrlf
				
			
								rs.movenext
							m = m + 1
							rowcount=rowcount-1
							loop
							end if
					end if
					
					if nowpage="" then nowpage=1
					if cint(nowpage)=cint(rs.pagecount) then
						set sumrs=conn.execute(sum_sql)
								response.write "<tr height=24 bgcolor='#FFCCCC'>"
								response.write "<td align=right colspan=4>合计：</td>"
								response.write "<td align=right>"&mydot(sumrs("sumyhj_dk_jine"))&"</td>"
								response.write "<td align=right colspan=4>&nbsp;</td>"
						sumrs.close
					end if
					
					%>
					</form>
					</table>
					<%
					response.write "<table border=0 width='98%' align=center><tr><td align=right height=30>"
							backnexturl="inputsearch="&inputsearch&"&inputkind="&inputkind&"&none=&"
							if nowpage="" then nowpage="1"
							lastpage=rs.pagecount
							nowpath=request.ServerVariables("SCRIPT_NAME")
							response.write "共"&rs.recordcount&"条 每页20条 现在第"&nowpage&"/"&rs.pagecount&"页&nbsp; "
							response.write "&nbsp;<input type='button' name='Button1' value='首页' onclick=window.location.href='"&nowpath&"?"&backnexturl&"page=1'>&nbsp;"
							if nowpage>1 then
								response.write"<input type='button' name='Button2' value=上页 onclick=window.location.href='"&nowpath&"?"&backnexturl&"page="&nowpage-1&"'>&nbsp;"
							else
								response.write"<input type='button' name='Button3' value=上页>&nbsp;"
							end if
							if clng(nowpage) < clng(lastpage) then
								response.write"<input type='button' name='Button4' value=下页 onclick=window.location.href='"&nowpath&"?"&backnexturl&"page="&nowpage+1&"'>&nbsp;"
							else
								response.write"<input type='button' name='Button5' value=下页>&nbsp;"
							end if
							response.write"<input type='button' name='Button6' value=尾页 onclick=window.location.href='"&nowpath&"?"&backnexturl&"page="&clng(lastpage)&"'>&nbsp;"
							response.write "</td><form name='form1' method='post' action='chzhdetail.asp?"&backnexturl&"'><td>转到第<input type=text name='page'size=3>页 <input type='submit' name='Submit' value='确定'>"
							response.write "&nbsp;&nbsp;<a href='javascript:window.scroll(0,0)'>TOP↑</a>"
					response.write "</td></form></tr></table>"
					
					response.cookies("hypage")=nowpath&"?"&backnexturl&"page="&nowpage
					
					
					rs.close
					set rs=nothing
					
					%>
					<br><br>
              <%end if%>
            </TD>
          </TR>
        </TBODY>
      </TABLE>

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

