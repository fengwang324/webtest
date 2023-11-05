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

<script language="javascript">
<!--
function SetBgColor(Menu)
{
		Menu.style.background="#F5F5F4";
}
function RestoreBgColor(Menu)
{
		Menu.style.background="#ffffff";
}
-->
</script>

</head>

<body>
<!-- #include file="head.asp" -->
<TABLE width="960" border=0 align="center" cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
  <tr>
    <td> 
<table width="960" height="352" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
    <td width="210" height="184" valign="top">
		<!-- #include file="inc_left_all.asp" -->
      </td>
    <td width="750" valign="top">
      <TABLE width="100%" 
                  border=0 align="center" cellPadding=0 cellSpacing=2 class=a1>
        <TBODY>
          <TR>
            <TD width="100%" height=30 colSpan=1><TABLE width="100%" 
border=0 cellPadding=0 cellSpacing=0 background="images/bg4.gif">
              <TBODY>
                <TR>
                          <TD width=135 
                            height=30 background=images/hotnewpro.gif>　　　<span class="style11">礼品礼券</span></TD>
                          <TD width="554">&nbsp;</TD>
                          <TD width=57>&nbsp;</TD>
                </TR>
              </TBODY>
            </TABLE>
                    
                  </TD>
          </TR>
		  </tbody>
		  </table>

            <table width="99%" border="0" align="center" cellpadding="0" cellspacing="0">
			<tr>
                  <td height=30><strong>一、本期现金优惠券：</strong></td>
			</tr>
			</table>
			
			<table width="99%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolorlight=#cccccc bordercolordark=#ffffff>
			  <tr bgcolor="#cccccc"> 
				<td width="3%"><div align="center">券号</div></td>
				<td width="10%"><div align="center">优惠券名称</div></td>
				<td width="18%"><div align="center">优惠券说明</div></td>
				<td width="6%"><div align="center">购满金额</div></td>
				<td width="6%"><div align="center">可抵扣金额</div></td>
				<td width="7%"><div align="center">使用开始日期</div></td>
				<td width="7%"><div align="center">使用结束日期</div></td>
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
			
			sql="select * from sh_youhj where yhj_qiyong='是' order by id desc"
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,1,1
			if rs.bof or rs.eof then
				response.write "<tr><td colspan='7' height=26><div align=center>本期没有现金优惠券。</div></td></tr>"
			else
				session("righturl")="1" 
			   rs.pagesize=10
			   rs.absolutepage=1
			   if request("page")<>"" then rs.absolutepage=request("page")
			   rowcount =rs.pagesize
			   if not rs.eof then
					do while not rs.eof and rowcount>0 
			
						id=rs("id")
						yhj_name=rs("yhj_name")
						yhj_memo=rs("yhj_memo")
						'yhj_qiyong=rs("yhj_qiyong")
						yhj_mz_jine=rs("yhj_mz_jine")
						yhj_dk_jine=rs("yhj_dk_jine")
						yhj_sy_time=rs("yhj_sy_time")
						yhj_zs_time=rs("yhj_zs_time")
						yhj_addtime=rs("yhj_addtime")
			
						response.write "<tr Onmouseover='return SetBgColor(this);' Onmouseout='return RestoreBgColor(this);'>"
						response.write "<td align=center>"&id&"</td>"
						response.write "<td height=25>"&yhj_name&"</td>"
						response.write "<td style='word-wrap: break-word; word-break: break-all;'>"&yhj_memo&"</td>"
						response.write "<td align=right>"&mydot(yhj_mz_jine)&"</td>"
						response.write "<td align=right>"&mydot(yhj_dk_jine)&"</td>"			
						response.write "<td>"&yhj_sy_time&"</td>"
						response.write "<td >"&yhj_zs_time&"</td>"
						response.write "</tr>"&vbcrlf
						rs.movenext
					m = m + 1
					rowcount=rowcount-1
					loop
					end if
			end if
			
			
			
			%>
			</table>



<FORM name="theForm" action="lipinupdate.asp" method="post" onSubmit="return Juge()">
            <table width="99%" border="0" align="center" cellpadding="0" cellspacing="0">
			<tr>
                  <td height=30><strong>二、本期可兑换礼品：</strong></td>
			</tr>
			</table>
			<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" id="tabcss">
              <tr bgcolor="#cccccc"> 
                <td height="25" width=16%> <div align="center">图片</div></td>
                <td width=10%> <div align="center">编号</div></td>
                <td width=30%>礼品名称</td>
                <td width=10%> <div align="center">所需积分</div></td>
                <td width=11%><div align="center">兑换数量</div></td>
              </tr>
              <%
i=1
sql="select * from x_lipin where showflag='1' order by num"
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof or rs.eof then
	response.write "<tr><td colspan=7><div align=center>没有记录! </div></td></tr>"
else
	do while not rs.eof 
	
	id=rs("id")
	picpath=rs("picpath")
	'picpath2=rs("picpath2")
	lipinno=rs("lipinno")
	lipinname=rs("lipinname")
	zhifen=rs("zhifen")
	'showflag=rs("showflag")
	'num=rs("num")
	'memo=rs("memo")

%>
              <tr onMouseOver="this.style.background='#f1f1f1';" onMouseOut="this.style.background='#ffffff';"> 
                <td align=center><a href='lipinshow.asp?id=<%=id%>' target='_blank'><img width="90" src='<%=managerfolder%>/<%=picpath%>' border='0'></a></td>
                <td><%=lipinno%><input type="hidden" name="id<%=i%>" value="<%=id%>"></td>
                <td><a href='lipinshow.asp?id=<%=id%>' target='_blank'><b><%=lipinname%></b></a></td>
                <td><div align="right"><%=zhifen%>&nbsp;<input type="hidden" name="zhifen<%=i%>" value="<%=zhifen%>"></div></td>
                <td><div align="center"><input type="input" name="shu<%=i%>" value="" size="3" style="ime-mode:disabled" onkeydown='myKeyDown()' autocomplete=off onBlur="sum()" onClick="sum()"></div></td>
              </tr>
              <%
	rs.movenext
	i=i+1
	loop
end if

rs.close
set rs=nothing

%>			<input type="hidden" name="maxi" size="5" value="<%=i-1%>">
            </table>
			<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
			  <td align=right height=40 width="85%">
			  <%
			  if session("username")="" or session("s_stat")="" then
			  else
					sql6="SELECT id, username,(select sum(billzhf) from x_view_bill where x_view_bill.username=a.username and sendornot='已发货') AS sumzhf, "&_
					"(select sum(kjzhf) from x_huanlipin where x_huanlipin.username=a.username) AS sumkjzhf,  "&_
					"(IIf(IsNull(sumzhf),0,sumzhf)-IIf(IsNull(sumkjzhf),0,sumkjzhf)) AS leftzhf  "&_
					"FROM x_huiyuan AS a where username='"&session("username")&"' "
					
					set rs6=server.createobject("adodb.recordset")
					rs6.open sql6,conn,1,1
					if rs6.bof or rs6.eof then
						response.write "<b>"&session("username")&"</b> 剩余积分：<input type='text' name='leftzhf' value='0' readonly style='width:85;border:0px;BACKGROUND-COLOR: #ffffff;FONT-WEIGHT: bold; FONT-SIZE: 17px; COLOR:#f00;FONT-FAMILY: Century Gothic'>"
					else
						response.write "<b>"&session("username")&"</b> 剩余积分：<input type='text' name='leftzhf' value='"&rs6("leftzhf")&"' readonly style='width:85;border:0px;BACKGROUND-COLOR: #ffffff;FONT-WEIGHT: bold; FONT-SIZE: 17px; COLOR:#f00;FONT-FAMILY: Century Gothic'>"
					end if
					rs6.close
					set rs6=nothing
			  end if
			  %>
			  &nbsp;&nbsp;兑换以上礼品所需积分：<input type='text' name='needzhf' value='' readonly style='width:85;border:0px;BACKGROUND-COLOR: #ffffff;FONT-WEIGHT: bold; FONT-SIZE: 17px; COLOR:#f00;FONT-FAMILY: Century Gothic'>
			  </td>
			  <td align=right width="15%">
			  <%if session("username")="" or session("s_stat")="" then%>
			  <INPUT type="button" name=update style="background-image: url(images/lipin.jpg);border:0;width:84px;height:24px;CURSOR:hand;" value="" onClick="alert('会员登陆后才能进行兑换礼品。');window.location.href='hyzq.asp?jieshun=2'">
			  <%else%>
			  <INPUT type="submit" name=update style="background-image: url(images/lipin.jpg);border:0;width:84px;height:24px;CURSOR:hand;" value="">
			  <%end if%>
			  &nbsp;&nbsp;
			  </td></tr>
			  <tr><td>提示：积分规则的一切解释权归<%=application("sitename")%>所有。</td></tr>
			  </table>
</form>

            <br>
            <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
                <td><strong><font color="#FF0000">礼品换取方法： </font></strong> <br>
                  只要累计达到礼品对应的积分即可换取礼品。请进入“礼品兑换”进行选择。礼品仅在购物的同时与商品一同发送，恕不单独发送。请先换取礼品，然后再选择想购买的商品，并在本订单附言中说明“有礼品一同发送”字样，而本单购物积分要在本单发货后才取使用。 
                  <p><strong><font color="#FF0000">积分查询方式：<br>
                    </font></strong>系统会按发货后自动更新所有会员的累计积分，会员专区可以查询到剩余积分及积分对帐情况。</p>
                  <p><strong><font color="#FF0000">温馨提示：</font></strong> <br>
                    积分仅对登录者本人有效。不可与其他会员的积分合并使用。积分不可折合成现金。<br>
                    如退货、拒收商品，您此次购买行为所累计的购物积分将予以扣除。已确认换取的礼品不可变更、取消或退换。<br>
                  积分兑换礼品为您保留30天，并将随您下一次订购的商品一同发送。</p>
                </td>
              </tr>
            </table>

            <br>
          <br></td>
  </tr>
</table>

    </td>
  </tr>

</table>

<script language="javascript">
<!--
function myKeyDown()
{
	var    k=window.event.keyCode;   
	if ((k==46)||(k==8)||(k==189)||(k==109)||(k==190)||(k==110)|| (k>=48 && k<=57)||(k>=96 && k<=105)||(k>=37 && k<=40)) 
	{}
	else if(k==13||k==9){
		 window.event.keyCode = 9;}
	else{
		 window.event.returnValue = false;}
}

function sum()
{
	var alltotal=0
	for (i=1;i<parseFloat(document.all.maxi.value)+1;i++)
	{
		
		if(document.all("shu"+i).value == parseFloat(document.all("shu"+i).value) )
		{alltotal=alltotal + Math.round(document.all("zhifen"+i).value * document.all("shu"+i).value*100)/100 ;}
	}
	document.all("needzhf").value = Math.round(alltotal*100)/100;
}

function Juge()
{
if( theForm.needzhf.value=="" || theForm.needzhf.value=="0") 
{
	alert("请填写要兑换礼品的数量。");
	return (false);
}

if( parseFloat(theForm.leftzhf.value) < parseFloat(theForm.needzhf.value) )
{
	alert("“剩余积分”不足，未能完成礼品兑换。本次兑换所需积分为："+theForm.needzhf.value+"。");
	return (false);
}


if(confirm("确认进行礼品兑换吗？"))
  return true
else
  return false;

}
-->
</script>
<!-- #include file="foot.asp" -->
<%
conn.close
set conn=nothing
%>
</body>
</html>
