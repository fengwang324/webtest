<!--#include file="conn.asp"-->
<!--#include file="check4.asp"-->
<%
dim sql4,rs4,id,num,pkid
dim sql5,rs5,model,productname,price2,price,money

fromyear=request("selyear")
frommonth=request("selmonth")
fromday=request("selday")
toyear=request("selyear2")
tomonth=request("selmonth2")
today=request("selday2")

if request("fromdate")="" then
	fromdate1=fromyear&"-"&frommonth&"-"&fromday
	todate1=toyear&"-"&tomonth&"-"&today
	fromdate=cdate(fromdate1)
	todate=cdate(todate1)+1
else
	fromdate=request("fromdate")
	todate=request("todate")
end if
	
model=trim(request("model"))
productname=trim(request("productname"))
username=trim(request("username"))
sendornot=trim(request("sendornot"))

nowpage=request("page")
if nowpage="" then nowpage="1"

condition=" where billdate between #"&fromdate&"# and #"&todate&"# "
if model<>"" then
	condition = condition & " and model='"&model&"' "
end if
if productname<>"" then
	condition = condition & " and productname like '%"&productname&"%' "
end if
if username<>"" then
	condition = condition & " and username='"&username&"' "
end if

if sendornot<>"" then
	condition = condition & " and sendornot='"&sendornot&"' "
end if
sql4="select * from x_view_report "&condition&" order by id"

'response.write "fromdate="&fromdate&"&todate="&todate&"&model="&trim(request("model"))&"&productname="&trim(request("productname"))&"&username="&trim(request("username"))&"&sendornot="&trim(request("sendornot"))&"&none=&"
'response.end
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>报表</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<style>
<!--
td{font-size:9pt}
body {
	margin-top: 5px;
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
	padding-left: 3px;
	padding-right: 2px;
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
<body bgcolor="#fcfcfc">

<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0" id="tabcss">
    <tr bgcolor="#BFDFFF"> 
                <td > 
                  <div align="center"><font color="#FFFFFF"><strong>订单号</strong></font></div></td>
                <td> 
                  <div align="center"><font color="#FFFFFF"><strong>订单日期</strong></font></div></td>
                <td> 
                  <div align="center"><font color="#FFFFFF"><strong>商品号</strong></font></div></td>
                <td> 
                  <div align="center"><font color="#FFFFFF"><strong>商品名</strong></font></div></td>
                <td> 
                  <div align="center"><font color="#FFFFFF"><strong>单 价</strong></font></div></td>
                <td> 
                  <div align="center"><font color="#FFFFFF"><strong>数 量</strong></font></div></td>
                <td> 
                  <div align="center"><font color="#FFFFFF"><strong>金 额</strong></font></div></td>
                <td> 
                  <div align="center"><font color="#FFFFFF"><strong>处理状态</strong></font></div></td>
              </tr>
<%
		set rs4=server.createobject("adodb.recordset")
		rs4.open sql4,conn,1,1
		if rs4.bof or rs4.eof then
			response.write "没有符合条件订单记录！"
		else
		i=1
		rs4.pagesize=20
		rs4.absolutepage=1
		if request("page")<>"" then rs4.absolutepage=request("page")
		rowcount=rs4.pagesize
		
			do while not rs4.eof and rowcount>0
				id=rs4("id")
				billno=rs4("billno")
				billdate=rs4("billdate")
				model=rs4("model")
				productname=rs4("productname")
				price=rs4("price")
				num=rs4("num")
				money=rs4("money")
				sendornot=rs4("sendornot")
				billdate=year(billdate)&"-"&month(billdate)&"-"&day(billdate)

				response.write "<tr bgcolor='#Ffffff' onMouseOver=""this.style.background='#FFE4CA';"" onMouseOut=""this.style.background='#ffffff';"">"
				response.write "<td height='20'><a href='dingdan_detail.asp?id="&id&"&billno="&billno&"' target='_blank'><font color=#0033FF>"&billno&"</font></a></td>"
				response.write "<td>"&billdate&"</td>"
				response.write "<td>"&model&"</td>"
				response.write "<td>"&productname&"</td>"
				response.write "<td align=right>"&formatnumber(price,2)&"</td>"
				response.write "<td align=right>"&formatnumber(num,2)&"</td>"
				response.write "<td align=right>"&formatnumber(money,2)&"</td>"
				response.write "<td align=center>"&sendornot&"</td>"
				response.write "</tr>"
				rs4.movenext
				i=i+1
				rowcount=rowcount-1
			loop
			if cint(nowpage)=cint(rs4.pagecount) then
				sql5="select sum(num) as sumnum,sum([money]) as summoney from x_view_report "&condition&" "
				set rs5=conn.execute(sql5)
				sumnum=rs5("sumnum")
				summoney=rs5("summoney")
				response.write "<tr bgcolor='#eeeeee'><td colspan=5 height=23 align=right>合计：</td><td align=right>"&formatnumber(sumnum,2)&"</td><td align=right>"&formatnumber(summoney,2)&"</td><td>&nbsp;</td></tr>"
			end if
		end if
			  %>
            </table>
<%
response.write "<table border=0 width='95%' align=center><tr><td align=right height=30>"
		backnexturl="fromdate="&fromdate&"&todate="&todate&"&model="&trim(request("model"))&"&productname="&trim(request("productname"))&"&username="&trim(request("username"))&"&sendornot="&trim(request("sendornot"))&"&none=&"
		
		if nowpage="" then nowpage="1"
		lastpage=rs4.pagecount
		nowpath=request.ServerVariables("SCRIPT_NAME")
		response.write "共"&rs4.recordcount&"条 每页20条 现在第"&nowpage&"/"&rs4.pagecount&"页&nbsp; "
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
		response.write "</td><form name='form1' method='post' action="&nowpath&"?"&backnexturl&"><td>到第<input type=text name='page'size=5>页 <input type='submit' name='Submit' value='确定'>"
		response.write "&nbsp;&nbsp;<a href='javascript:window.scroll(0,0)'>TOP↑</a>"
response.write "</td></form></tr></table>"

response.cookies("rppage")=nowpath&"?"&backnexturl&"page="&nowpage
%><br><br><br>

</body>
</html>
<%
rs4.close
set rs4=nothing
call connclose
%>

