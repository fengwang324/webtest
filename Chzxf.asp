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
<SCRIPT language=javascript>
function Juge()
{
	if (document.theForm.password.value=="")
	{
		alert ("请输入用户登陆密码。")
		document.theForm.password.focus();
		return false;
	}
	if (document.theForm.password.value.length<4||document.theForm.password.value.length>20)
	{
		alert ("用户登陆密码字符数太短或太长")
		document.theForm.password.focus();
		return false;
	}
	if (document.theForm.password.value!=document.theForm.password2.value)
	{
		alert ("登陆密码确认错误")
		document.theForm.password2.focus();
		return false;
	}
	if (document.theForm.password_prompt.value=="")
	{
		alert ("请输入密码提示问题")
		document.theForm.password_prompt.focus();
		return false;
	}
	if (document.theForm.password_Answer.value=="")
	{
		alert ("请输入密码提示问题答案")
		document.theForm.password_Answer.focus();
		return false;
	}
	
	if (document.theForm.truename.value=="")
	{
		alert ("请输入用户真实姓名")
		document.theForm.truename.focus();
		return false;
	}
	
	if (document.theForm.telphone1.value=="")
	{
		alert ("请输入联系电话1")
		document.theForm.telphone1.focus();
		return false;
	}
	
	if (document.theForm.address.value=="")
	{
		alert ("请输入联系地址")
		document.theForm.address.focus();
		return false;
	}


   if(confirm("现在真的要修改您的个人资料吗?"))
      return true
   else
      return false;
}
</SCRIPT>
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
					<table width="98%" border="0" align="center" cellpadding="1" cellspacing="0"><tr><td><b>充值-消费对帐：</b></td></tr></table><table width="98%" border="1" align="center" cellpadding="1" cellspacing="0" bordercolorlight=#cccccc bordercolordark=#ffffff>
					<form name="theForm" method="post" action="chongzhidel.asp?menuid=<%=menuid%>">
					  <tr bgcolor="#cccccc" height=23> 
						<td width="8%"><div align="center">日期</div></td><td width="8%"><div align="center">用户名</div></td>
						<td width="8%"><div align="center">类型</div></td><td width="7%"><div align="center">金额</div></td></tr>
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
					
					sql="select * from view_cz_xf_union where username='"&session("username")&"' order by czdate"
					
					sum_sql="select sum(czmoney) as sumczmoney from view_cz_xf_union where username='"&session("username")&"'"
					
					set rs=server.createobject("adodb.recordset")
					rs.open sql,conn,1,1
					if rs.bof or rs.eof then
						response.write "<tr><td colspan='8'><div align=center>当前没有符合条件记录。"
						response.write "</div></td></tr>"
					else
						session("righturl")="1" 
					   rs.pagesize=20
					   rs.absolutepage=1
					   if request("page")<>"" then rs.absolutepage=request("page")
					   rowcount =rs.pagesize
					   if not rs.eof then
							do while not rs.eof and rowcount>0 
								
								kind=rs("kind")
								id=rs("id")
								czdate=rs("czdate")
								customid=rs("customid")
								username=rs("username")
								czmoney=rs("czmoney")
					
					
								response.write "<tr Onmouseover='return SetBgColor(this);' Onmouseout='return RestoreBgColor(this);'>"
								if kind="充值" then
								response.write "<td height=24>"&czdate&"&nbsp;</td>"
								else
								response.write "<td height=24><a style='CURSOR:hand;' onclick=""window.open('order_print.asp?id="&id&"&billno="&billno&"','','top=55,left=130,scrollbars=yes,width=680,height=500,resizable=yes,menubar=yes,resizable=yes')""><font color=#0033FF>"&czdate&"</font></a>&nbsp;</td>"
								end if
								response.write "<td >"&username&"&nbsp;</td>"
								response.write "<td >"&kind&"&nbsp;</td>"
								response.write "<td align=right>"&mydot(czmoney)&"&nbsp;</td>"			
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
								response.write "<td align=right colspan=3>合计：</td>"
								response.write "<td align=right>"&mydot(sumrs("sumczmoney"))&"&nbsp;</td>"
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

