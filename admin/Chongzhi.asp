<!--#include file="conn.asp"-->
<!--#include file="check4.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link REL='stylesheet' HREF='../szcrm.css' TYPE='text/css'>
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
<style>
.tablehead
{
	background-color:#B4EDC2; COLOR:#FFFFFF; text-align:center;   FONT-SIZE: 14px;color :#FFFFFF;
    FONT-FAMILY: 宋体，Arial,Times New Roman
}
td{font-size:9pt;line-height:140%;}
body {
	margin-top: 5px;
}
</style>
</head>
<body topmargin="5">
  <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td height=5></td>
    </tr>
  </table>
<%
dim nowpage,sql,rs,rowcount,myorder,menuid
dim id,customid,operator,chongzhidate,constract,actionkind,actionglass,actionsubject
dim m
dim gold_rs,gold_sql
dim chongzhidate1,chongzhidate2,actionkind1,actionsubject1,actionaddr1,alertdate1,alertdate2,memo1,advsearch,advcondition

m=1

nowpage=request("page")
if nowpage="" then nowpage=1
myorder=request("myorder")
menuid=request("menuid")
sublist=request("sublist")
customid=request("customid")
inputsearch=trim(request("inputsearch"))
inputkind=trim(request("inputkind"))
desk=request("desk")
'--------------------


sql_condition="where id>0 "
'-----------begin-------------------
s_selec=request("s_selec")
s_date1=request("s_date1")
s_date2=request("s_date2")
if s_selec<>"" then
	if isdate(s_date1) then
		sql_condition = sql_condition & " and "&s_selec&">#"&cdate(s_date1)-1&"# "
	end if
	if isdate(s_date2) then
		sql_condition = sql_condition & " and "&s_selec&"<#"&cdate(s_date2)+1&"# "
	end if
end if
searchdate="&s_selec="&s_selec&"&s_date1="&s_date1&"&s_date2="&s_date2&""
'------------end---------------------


if chongzhidate1<>"" and chongzhidate2<>"" then
	sql_condition = sql_condition & " and (chongzhidate between #"&chongzhidate1&"# and #"&chongzhidate2&"#) "
end if
if alertdate1<>"" and alertdate2<>"" then
	sql_condition = sql_condition & " and (alertdate between #"&alertdate1&"# and #"&alertdate2&"#) "
end if

if inputsearch<>"" then
	sql_condition = sql_condition & " and "&inputkind&" like '%"&inputsearch&"%' "
end if
if desk="1" then
	sql_condition = sql_condition & " and operator='"&session("username")&"' "
end if

if myorder="" then
	order_condition=" id desc "
else
	order_condition=" "&myorder&" "
end if

sql="select * from VIEW_chongzhi "&sql_condition&" order by "&order_condition

sum_sql="select sum(czmoney) as sumczmoney from VIEW_chongzhi "&sql_condition

'response.write sql

%>
<table width="99%"  border="0" align="center" cellpadding="3" cellspacing="0">
<form name="theForm22" method="post" action="<%=nowpath%>?menuid=<%=request("menuid")%>&sublist=<%=request.Cookies("sublist")%>">
    <tr bgcolor="#BFDFFF">
	<td width="10%" height="18"><b>会员充值：</b></td>
      <td height="29">查询:<select language='javascript' style="width=100px;" name='inputkind' size='1'>
          <option value='username'>用户名
		  <option value='mobile'>手机
          <option value='czby'>充值方式
		  <option value='czmoney'>充值金额
          <option value='czmemo'>备注
		  <option value='confirmflag'>是否确认
		  <option value='operator'>确认人
        </select>&nbsp;关鍵字:<input name='inputsearch' type='text' class=input size='10' value=''> <input type=submit name='search' value='查询'>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
</form>
</table>
  
<table width="99%" border="1" align="center" cellpadding="1" cellspacing="0" bordercolorlight=#cccccc bordercolordark=#ffffff>

  <tr bgcolor="#cccccc"> 
    <td width="8%"><div align="center">充值日期</div></td>
    <td width="8%"><div align="center">用户名</div></td>
    <td width="8%"><div align="center">手机</div></td>
    <td width="7%"><div align="center">充值方式</div></td>
    <td width="7%"><div align="center">充值金额</div></td>
    <td width="6%"><div align="center">确认</div></td>
    <td width="15%"><div align="center">最后确认人/时间</div></td>
	<td width="19%"><div align="center">备注</div></td>
    <td width="4%" height="23">&nbsp;</td>
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

set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof or rs.eof then
	response.write "<tr><td colspan='8'><div align=center>当前没有符合条件记录。"
	response.write "<input name='b1' type='button' value='新增' onclick=""window.location.href='chongzhiadd.asp';""> "
	response.write "</div></td></tr>"
else
	session("righturl")="1" 
   rs.pagesize=20
   rs.absolutepage=1
   if request("page")<>"" then rs.absolutepage=request("page")
   rowcount =rs.pagesize
   if not rs.eof then
   		do while not rs.eof and rowcount>0 
			id=rs("id")
			czdate=addspace(rs("czdate"))
			customid=addspace(rs("customid"))
			username=addspace(rs("username"))
			mobile=addspace(rs("mobile"))
			czby=addspace(rs("czby"))
			czmoney=addspace(rs("czmoney"))
			czmemo=addspace(rs("czmemo"))
			confirmflag=addspace(rs("confirmflag"))
			operator=addspace(rs("operator"))

  			response.write "<tr Onmouseover='return SetBgColor(this);' Onmouseout='return RestoreBgColor(this);'>"
    		response.write "<td height=24><a href='chongzhishow.asp?menuid="&menuid&"&id="&id&"'>"&czdate&"</a></td>"
    		response.write "<td >"&username&"</td>"
			response.write "<td >"&mobile&"</td>"
			response.write "<td align=center>"&czby&"</td>"
			response.write "<td align=right>"&mydot(czmoney)&"</td>"			
			
			if confirmflag="是" then
			response.write "<td bgcolor='#69DA8D' align=center>"&confirmflag&"</td>"
			else
			response.write "<td bgcolor='#FFB997' align=center>"&confirmflag&"</td>"
			end if
			response.write "<td>"&operator&"</td>"
			response.write "<td style='word-wrap: break-word; word-break: break-all;'>"&czmemo&"</td>"
			response.write "<td align='center'><a href='chongzhidel.asp?strid="&id&"' onClick=""return confirm('是否此记录吗?')"">删除</a></td>" 
  			response.write "</tr>"
			rs.movenext
		m = m + 1
		rowcount=rowcount-1
		loop
		end if
end if

if cint(nowpage)=cint(rs.pagecount) then
	set sumrs=conn.execute(sum_sql)
  			response.write "<tr height=24 bgcolor='#FFCCCC'>"
    		response.write "<td align=right colspan=4>合计：</td>"
			response.write "<td align=right>"&mydot(sumrs("sumczmoney"))&"</td>"
			response.write "<td align=right colspan=4>&nbsp;</td>"
	sumrs.close
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
		response.write "<input name='b1' type='button' value='新增' onclick=""window.location.href='chongzhiadd.asp';"">&nbsp; "
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
		response.write "</td><form name='form1' method='post' action="&nowpath&"?"&backnexturl&"><td>转到第<input type=text name='page'size=3>页 <input type='submit' name='Submit' value='确定'>"
		response.write "&nbsp;&nbsp;<a href='javascript:window.scroll(0,0)'>TOP↑</a>"
response.write "</td></form></tr></table>"

response.cookies("hypage")=nowpath&"?"&backnexturl&"page="&nowpage


rs.close
set rs=nothing

conn.close
set conn=nothing

%>
<br><br>
</body>
</html>







