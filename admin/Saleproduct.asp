<!--#include file="conn.asp"-->
<!--#include file="check4.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>销售排行榜</title>
<link href="i.css" type=text/css rel=stylesheet>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 5px;
	margin-bottom: 0px;
	margin-right: 0px;
}
.style1 {font-size: 14px}
.style12 {color: #990000}
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
<%
		dim sql4,rs4,id,num,pkid
		dim sql5,rs5,model,productname,price2,price,money,a,q,allmoney,allnum
		allmoney=0
		allnum=0
		nowpage=request("page")
	
	
		sql_condition=" where sendornot='已发货' "
		
		'----------begin----------
		fromyear=request("selyear")
		frommonth=request("selmonth")
		fromday=request("selday")
		toyear=request("selyear2")
		tomonth=request("selmonth2")
		today=request("selday2")
		
		if request("fromdate")="" and fromyear<>"" then
			fromdate1=fromyear&"-"&frommonth&"-"&fromday
			todate1=toyear&"-"&tomonth&"-"&today
			fromdate=cdate(fromdate1)
			todate=cdate(todate1)+1
		elseif request("fromdate")<>"" then
			fromdate=request("fromdate")
			todate=request("todate")		
		else
		end if
		if request("checkdate")="1" then
			if fromdate<>"" then
				sql_condition=sql_condition & " and billdate between #"&fromdate&"# and #"&todate&"# "
			end if
		end if
		'----------end----------backnexturl="checkdate="&request("checkdate")&"&
%>

</head>

<body bgcolor="#fcfcfc">

<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0" id="tabcss">
    <tr bgcolor="#BFDFFF"> 
			<form name="form1" method="post" action="saleproduct.asp">
                <td height=30 colspan=6><b>销售商品排行</b>
<script>
function ChangDate(formyear,formmonth,formday,isBlank){
	var i;
	var j;
	var year,month,day;
	year=formyear.value;
	month=formmonth.value;
	if (month=="2"){
		if (((year%400)==0) || (((year%100)!=0) && ((year%4)==0))){
			j=29;
		}else{
			j=28;
		}
	}else{
		if (month=="4" || month=="6" || month=="9" || month=="11"){
			j=30;
		}else{
			j=31;
		}
	}
	var k=formday.length;
	for (i=0;i<k;i++){
		formday.remove(formday.i);
	}
	if (isBlank){
		var oOption = document.createElement("OPTION");
		oOption.text="请选择";
		oOption.value="";
		formday.add(oOption);
	}
	for (i=0;i<j;i++){
		var oOption = document.createElement("OPTION");
		oOption.text=i+1;
		oOption.value=i+1;
		formday.add(oOption);
	}
}
</script>
<%
sub CreateSelect(pSelName,pStart,pEnd,pSelected,pOption)
		dim intStart,intEnd,intSelected,i
		intStart=int(pStart)
		intEnd=int(pEnd)
		response.write "<select name=" & pSelName & " style='width=55;BACKGROUND-COLOR: #ddeeff;'>"
		if pOption="" then
			if not isnumeric(pSelected) then 
				response.write "<option value='' selected>"+ pSelected +"</option> "
				intSelected=""
			else
				intSelected=cint(pSelected)
			end if
		end if
		for i=intStart to intEnd
			if intSelected="" and isnumeric(pOption) then
				if i=cint(pOption) then
					response.write "<option value='' selected>"+ pSelected +"</option> "
				end if
			end if
			if i=intSelected and intSelected<>"" then
				response.write "<option value=" & trim(i) & " selected>" & trim(i) & "</option> "
			else
				response.write "<option value=" & trim(i) & ">" & trim(i) & "</option> "
			end if
		next 
		response.write "</select>"
end sub
%>
<%
response.write "日期: <input type='checkbox' name='checkdate' value='1'"
if request("checkdate")="1" then response.write " checked"
response.write ">从"
CreateSelect "SelYear onChange='ChangDate(SendSample.SelYear,SendSample.SelMonth,SendSample.SelDay,false);'",year(now())-30,year(now())+1,year(now()),""
response.write "年"
CreateSelect "SelMonth onChange='ChangDate(SendSample.SelYear,SendSample.SelMonth,SendSample.SelDay,false);'",1,12,month(now()),""
response.write "月"
if month(now())=2 then
	if cint(year(now())) mod 400 =0 or (cint(year(now())) mod 100 <>0 and cint(year(now())) mod 4 =0) then
		strEndDay=29
	else
		strEndDay=28
	end if
else
	if month(now())=4 or month(now())=6 or month(now())=9 or month(now())=11 then
		strEndDay=30
	else
		strEndDay=31
	end if
end if 
CreateSelect "SelDay",1,strEndDay,day(now()),""
response.write "日"
response.write "&nbsp;&nbsp;至"

CreateSelect "SelYear2 onChange='ChangDate(SendSample.SelYear2,SendSample.SelMonth2,SendSample.SelDay2,false);'",year(now())-30,year(now())+1,year(now()),""
response.write "年"
CreateSelect "SelMonth2 onChange='ChangDate(SendSample.SelYear2,SendSample.SelMonth2,SendSample.SelDay2,false);'",1,12,month(now()),""
response.write "月"
if month(now())=2 then
	if cint(year(now())) mod 400 =0 or (cint(year(now())) mod 100 <>0 and cint(year(now())) mod 4 =0) then
		strEndDay=29
	else
		strEndDay=28
	end if
else
	if month(now())=4 or month(now())=6 or month(now())=9 or month(now())=11 then
		strEndDay=30
	else
		strEndDay=31
	end if
end if 
CreateSelect "SelDay2",1,strEndDay,day(now()),""
response.write "日"
%>
 <input type="submit" name="Submit" value="确 定">&nbsp;<%=fromdate%>&nbsp;&nbsp;<%=todate%>
				
				
				</td></form>
                </tr>
			  <%if fromdate<>"" then%>
              <tr><td colspan="6">从 <%=fromdate%>&nbsp;&nbsp;至 <%=todate%></td></tr>
			  <%end if%>
              <tr bgcolor="#E1E1E1"> 
                  
                <td width="53" height="23" > 
<div align="center"><strong>序号</strong></div></td>
                  
                <td width="126"> 
<div align="center"><strong>商品号</strong></div></td>
                  
                <td width="186"> 
<div align="center"><strong>商品名</strong></div></td>
				  
                <td width="112"> 
<div align="center"><strong>品牌</strong></div></td>
                  
                <td width="79"> 
<div align="center"><strong>数量</strong></div></td>
                  
                <td width="106"> 
<div align="center"><strong>金额</strong></div></td>
                </tr>
                <%
	
		sql4="select productpkid,model,productname,pipai,sum(num) as sumnum,sum([money]) as summoney from x_view_report "&sql_condition&" group by productpkid,model,productname,pipai order by sum(num) desc "
		set rs4=server.CreateObject("adodb.recordset")
		rs4.open sql4,conn,1,1
		if rs4.bof or rs4.eof then
		else
		i=1
		rs4.pagesize=20
		rs4.absolutepage=1
		if request("page")<>"" then rs4.absolutepage=request("page")
		rowcount=rs4.pagesize

			do while not rs4.eof and rowcount>0
			productpkid=rs4("productpkid")
			model=rs4("model")
			productname=rs4("productname")
			pipai=addspace(rs4("pipai"))
			sumnum=rs4("sumnum")
			'price=rs4("price")
			summoney=rs4("summoney")

					response.write "<tr onMouseOver=""this.style.background='#FFE4CA';"" onMouseOut=""this.style.background='#fcfcfc';"">" 
					response.write "<td height='23' align='center'>"&i&"</td>"
					response.write "<td><a href='../show.asp?pkid="&productpkid&"' target='_blank'><font color=#0000ff>"&model&"</font></a></td>"
					response.write "<td>"&productname&"</td>"
					response.write "<td>"&pipai&"</td>"
					response.write "<td align=right>"&formatnumber(sumnum,2)&"</td>"
					response.write "<td align=right>＄"&formatnumber(summoney,2)&"</td>"
					response.write "</tr>"
			  rs4.movenext
			  rowcount=rowcount-1
			  i=i+1
			  loop
		end if

if nowpage="" then nowpage="1"
if cint(rs4.pagecount)=cint(nowpage) then
	sql5="select sum(num) as sumnum,sum([money]) as summoney from x_view_report "&sql_condition
	set rs5=conn.execute(sql5)
	if rs5.bof or rs5.eof then
	else
		sumnum=rs5("sumnum")
		summoney=rs5("summoney")
%>
			  
                <tr bgcolor="#E1E1E1"> 
                  <td height=25 colspan=4 align=right><strong>合计：</strong></td>
                  <td align=right><strong><%=formatnumber(sumnum,2)%></strong></td>
                  <td align=right><strong>＄<%=formatnumber(summoney,2)%></strong></td>
                </tr>
<%
	end if
	rs5.close
	set rs5=nothing
end if
%>
				
              </table>
            <%
response.write "<table border=0 width='98%' align=center><tr><td align=right height=30>"
		backnexturl="checkdate="&request("checkdate")&"&fromdate="&fromdate&"&todate="&todate&"&keyword="&keyword&"&selectkind="&selectkind&"&comefrom="&comefrom&"&none=&"
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
		response.write "</td><form name='form1' method='post' action="&nowpath&"?"&backnexturl&"><td>转到第<input type=text name='page'size=5>页 <input type='submit' name='Submit' value='确定'>"
		response.write "&nbsp;&nbsp;<a href='javascript:window.scroll(0,0)'>TOP↑</a>"
response.write "</td></form></tr></table>"

rs4.close
set rs4=nothing
call connclose
%>





</body>
</html>

