<!--#include file="conn.asp"-->
<!--#include file="check4.asp"-->
<!-- #include file="../product/page.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link REL='stylesheet' HREF='../szcrm.css' TYPE='text/css'>
<title>请选择会员</title>
<SCRIPT LANGUAGE="JavaScript"> 
<!-- 
function pick(symbol1,symbol2,symbol3) { 
if (window.opener && !window.opener.closed) 
window.opener.document.theForm.customname.value = symbol1; 
window.opener.document.theForm.customid.value = symbol2; 
window.close(); 
} 
// --> 
</SCRIPT>
<script language="javascript">
<!--
function SetBgColor(Menu)
{
		Menu.style.background="#E7E7E7";
}
function RestoreBgColor(Menu)
{
		Menu.style.background="#ffffff";
}
-->
</script>
</head>
<body topmargin="10" >
  <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td height=6></td>
    </tr>
  </table>
  <table width="95%"  border="0" align="center" cellpadding="0" cellspacing="0">
<form name="theForm22" method="post" action="openhy.asp?menuid=<%=request("menuid")%>">
    <tr>
	<td height="18">
请选择：
<select language='javascript' class=select name='inputkind' size='1'>
          <option value="username">用户名</option> 
          <option value="id">会员编号</option>  
          <option value="truename">真实姓名</option>
          <option value="province">省份</option>
          <option value="city">城市</option>
          <option value="telphone1">联系电话1</option>
          <option value="telphone2">联系电话2</option>
          <option value="fax">传真</option>
          <option value="mobile">移动电话</option>
          <option value="address">联系地址</option>
          <option value="postno">邮政编码</option>
          <option value="email">电子邮件</option>
</select>&nbsp;关鍵字：
<input name='inputsearch' type='text' class=input size='12' value=''>&nbsp;&nbsp;
<input type=submit name='search' value='查询'>
</td>
<td>&nbsp;</td>
    </tr>
</form>
  </table>
  <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td height=4></td>
    </tr>
  </table>
  <%
dim nowpage,sql,rs,rowcount,myorder,menuid,order_condition
dim id,customname,sex,cardnum,city,carddate,zmd,BA,customname2,cardnum2
dim m,mykind,inputsearch,inputkind,sql_condition
m=1

nowpage=request("page")
myorder=request("myorder")
menuid=request("menuid")

inputsearch=trim(request("inputsearch"))
inputkind=trim(request("inputkind"))

mykind=trim(request("mykind"))

function changechar(info)
	if info<>"" then
		info=replace(info,"	","")
		info=replace(info," ","")
	end if
	changechar=info
end function

sql_condition="where id>2 "

if inputsearch<>"" then
	sql_condition = sql_condition & " and "&inputkind&" like '%"&inputsearch&"%' "
end if

if myorder="" then
	order_condition=" id desc "
else
	order_condition=" "&myorder&" "
end if

sql="select * from x_huiyuan "&sql_condition&" order by "&order_condition

set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof or rs.eof then
	response.write "<div align=center>当前没有会员记录。"
	response.write "</div>"
else
%>

  <table width="98%" border="1" align="center" cellpadding="1" cellspacing="0" bordercolorlight=#5BD77D bordercolordark=#ffffff>
  <form name="theForm2" method="post" action="b_customdel.asp?menuid=<%=menuid%>">
  <tr bgcolor="#E1E1E1"> 
    <td height="24"> 
<div align="center">用户名</div></td>
    <td> 
<div align="center">真实姓名</div></td>
    <td> 
<div align="center">联系电话1</div></td>
    <td > 
<div align="center">联系电话2</div></td>
    <td > 
<div align="center">移动电话</div></td>
    <td > 
<div align="center">联系地址</div></td>
  </tr>
    <%
   rs.pagesize=20
   rs.absolutepage=1
   if request("page")<>"" then rs.absolutepage=request("page")
   rowcount =rs.pagesize
   if not rs.eof then
   		do while not rs.eof and rowcount>0 
			id=rs("id")
			username=rs("username")
			vipno=addspace(rs("vipno"))
			truename=rs("truename") 
			telphone1=rs("telphone1") 
			telphone2=addspace(rs("telphone2"))
			mobile=addspace(rs("mobile")) 
			address=rs("address") 
			addtime=rs("addtime") 

			username2=changechar(username)
    		
  			response.write "<tr Onmouseover='return SetBgColor(this);' Onmouseout='return RestoreBgColor(this);'>"
    		response.write "<td height='20'>&nbsp;<a href=javascript:pick('"&username2&"','"&id&"')>"&username&"</a></td>"
			%>
			<td><%=truename%></td>
			<td><%=telphone1%></td>
			<td><%=telphone2%></td>
			<td><%=mobile%></td>
			<td style="WIDTH: 200; WORD-WRAP: break-word;line-height:140%;"><%=address%></td>
			
			<%
  			response.write "</tr>"
			rs.movenext
		m = m + 1
		rowcount=rowcount-1
		loop
		end if
%>
  </table>
<%
response.write "<table border=0 width='97%' align=center><tr><td align=right height=30>"
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
response.write "</td></tr></table>"

end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</body>
</html>



