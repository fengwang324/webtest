<!--#include file="conn.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>订单详细</title>
<link href="i.css" type=text/css rel=stylesheet>

<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-bottom: 0px;
	margin-right: 0px;
}
.style1 {font-size: 14px}
.style12 {color: #990000}
-->
</style>
<script language="javascript" >
function dialog(url,name,w,h)
{
	var lookups=window.showModalDialog(url, name, "dialogHeight:" + h + "px; dialogWidth:" + w + "px; center: Yes; help: No; resizable: No; status: No;");
	if(lookups!=null)
	{
		return lookups;
	}
	else
	{
		return null;
	}
}
function getDayDay(inputname)
{
inputname.value = dialog('date.html','mydate','195','225');
}
</script>
</head>

<body bgcolor="#fcfcfc">
              <%
dim sql7,rs7,vipno,truename,province,city,telphone1,telphone2,fax,mobile,address,postno,email
	sql7="select * from x_view_bill where id="&request.QueryString("id")&""
	set rs7=conn.execute(sql7)
	if rs7.bof or rs7.eof then
	else
		id=rs7("id")
		vipno=rs7("vipno") 
		truename=rs7("truename") 
		province=rs7("province") 
		city=rs7("city") 
		telphone1=rs7("telphone1") 
		telphone2=rs7("telphone2") 
		fax=rs7("fax") 
		mobile=rs7("mobile") 
		address=rs7("address") 
		postno=rs7("postno") 
		email=rs7("email")
		billdate=rs7("billdate")
		billno=rs7("billno")
		sendtype=rs7("sendtype")
		paytype=rs7("paytype")
		paystat=rs7("paystat")
		memo=rs7("memo")
		sendornot=rs7("sendornot")
		customid=rs7("customid")
		username1=rs7("username")
		
		sendmoney=rs7("sendmoney")
		total=rs7("total")
		
		youhjmoney=rs7("youhjmoney")  		'优惠券抵扣金额
		lastneedpay=rs7("lastneedpay")		'最后所需付金额
		
		senddate=rs7("senddate")
		sendcom=rs7("sendcom")
		sendno=rs7("sendno")
		sendmemo=rs7("sendmemo")
		
	end if
	rs7.close
	set rs7=nothing

%>
<TABLE width="622" border=0 align="center" cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
  <tr>
    <td> 
<table width="622" height="352" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          
          <td width="557" height="184" valign="top"><table width="622" border=0 align="center" cellPadding=3 cellSpacing=1 bgColor=#cccccc class=black-pix12 id=Table3>

              <tr> 
                <td width="110" height="26"><strong><font color="#CC0000">顾客资料：</font></strong></td>
                <td colspan="3">&nbsp;</td>
              </tr>
              <TR bgColor=#ffffff> 
                <TD class=black align=right width=110> <DIV align=right>用户名：</DIV></TD>
                <TD width=221> <input type="hidden" name="customid" value="<%=customid%>"> 
                  <b><%=username1%></b> </TD>
                <TD width=90> <div align="right">收货人姓名：</div></TD>
                <TD width=172> <input id=Text2 style="FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" size=15 name=truename value="<%=truename%>"></TD>
              </TR>
              <TR bgColor=#ffffff> 
                <TD class=black align=right width=110> <DIV align=right>省份：</DIV></TD>
                <TD width=221> <SELECT id=Select1 name=province>
                    <OPTION value="<%=province%>" selected><%=province%></OPTION>
                    <OPTION value="广东">广东</OPTION>
                    <OPTION value="北京">北京</OPTION>
                    <OPTION value="上海">上海</OPTION>
                    <OPTION value="天津">天津</OPTION>
                    <OPTION value="安徽">安徽</OPTION>
                    <OPTION value="重庆">重庆</OPTION>
                    <OPTION value="福建">福建</OPTION>
                    <OPTION value="甘肃">甘肃</OPTION>
                    <OPTION value="广西">广西</OPTION>
                    <OPTION value="贵州">贵州</OPTION>
                    <OPTION value="海南">海南</OPTION>
                    <OPTION value="河北">河北</OPTION>
                    <OPTION value="河南">河南</OPTION>
                    <OPTION value="黑龙江">黑龙江</OPTION>
                    <OPTION value="湖北">湖北</OPTION>
                    <OPTION value="湖南">湖南</OPTION>
                    <OPTION value="吉林">吉林</OPTION>
                    <OPTION value="江苏">江苏</OPTION>
                    <OPTION value="江西">江西</OPTION>
                    <OPTION value="辽宁">辽宁</OPTION>
                    <OPTION value="内蒙古">内蒙古</OPTION>
                    <OPTION value="宁夏">宁夏</OPTION>
                    <OPTION value="青海">青海</OPTION>
                    <OPTION value="山东">山东</OPTION>
                    <OPTION value="山西">山西</OPTION>
                    <OPTION value="陕西">陕西</OPTION>
                    <OPTION value="四川">四川</OPTION>
                    <OPTION value="西藏">西藏</OPTION>
                    <OPTION value="新疆">新疆</OPTION>
                    <OPTION value="云南">云南</OPTION>
                    <OPTION value="浙江">浙江</OPTION>
                    <OPTION value="香港">香港</OPTION>
                    <OPTION value="澳门">澳门</OPTION>
                    <OPTION value="台湾">台湾</OPTION>
                    <OPTION value="其它">其它</OPTION>
                  </SELECT> </TD>
                <TD width=90> <div align="right">城市：</div></TD>
                <TD width=172> <input name=city value="<%=city%>" id=city size="13" ></TD>
              </TR>
              <TR bgColor=#ffffff> 
                <TD class=black align=right width=110> <DIV align=right>联系电话1：</DIV></TD>
                <TD width=221> <INPUT id=Text9 style="FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" size=15 name="telphone1" value="<%=telphone1%>"> 
                </TD>
                <TD width=90> <div align="right">联系电话2：</div></TD>
                <TD width=172> <input id=telphone2 style="FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" size=15 name="telphone2" value="<%=telphone2%>"></TD>
              </TR>
              <TR bgColor=#ffffff> 
                <TD class=black align=right width=110> <DIV align=right>移动电话：</DIV></TD>
                <TD width=221> <input id='mobile' style="FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" 
                  size=15 name='mobile' value="<%=mobile%>"> </TD>
                <TD width=90> <div align="right">传真：</div></TD>
                <TD width=172> <input name=fax value="<%=fax%>" id=fax style="FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" 
                  size=15 maxlength="15"></TD>
              </TR>
              <TR bgColor=#ffffff> 
                <TD class=black align=right width=110> <DIV align=right>送货地址：</DIV></TD>
                <TD colspan="3" class=black> <INPUT id=Text11 style="FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" 
                  size=40 name=address value="<%=address%>">
                  邮编： 
                  <input name=postno value="<%=postno%>" id=postno style="FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" 
                  size=15 maxlength="15"> </TD>
              </TR>
              <TR bgColor=#ffffff> 
                <TD class=black align=right width=110> <DIV align=right>电子邮件：</DIV></TD>
                <TD class=black> <input name="email" value="<%=email%>" id=email2 
                  style="FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" 
                  size=30 maxlength="50"> </TD>
                <TD class=black><div align="right"></div></TD>
                <TD class=black>&nbsp;</TD>
              </TR>
              <tr> 
                <td width="108" height="23"><font color="#CC0000"><strong>订单信息：</strong></font></td>
                <td colspan="3">&nbsp;</td>
              </tr>
              <TR bgColor=#ffffff> 
                <TD height="23" align=middle> <div align="right">订单编号：</div></TD>
                <TD><b><%=billno%></b></TD>
                <TD><div align="right">下单时间：</div></TD>
                <TD><%=billdate%></TD>
              </TR>
              <TR bgColor=#ffffff> 
                <TD align=middle width=108> <div align="right">送货方式：</div></TD>
                <TD width=194><%=sendtype%></TD>
                <TD align="right">付款方式：</TD>
                <TD width=207><%=paytype%></TD>
              </TR>
              <TR bgColor=#ffffff>
                <TD align=middle> <div align="right">送货费用：</div></TD>
                <TD>＄<b><%=formatnumber(sendmoney,2)%></b></TD>
                <TD align="right">&nbsp;</TD>
                <TD>&nbsp;</TD>
              </TR>
              <!--
                <TR bgColor=#ffffff> 
                  <TD align=middle width=115> 
<div align="right">是否需要发票：</div></TD>
                  <TD width=401> 
                    <SELECT name=fapiao>
                      <OPTION value="否" selected>否</OPTION>
                      <OPTION value="是">是</OPTION> ***
                    </SELECT></TD>
                </TR>
-->
              <TR bgColor=#ffffff> 
                <TD align=middle width=108> <div align="right">特别说明：</div></TD>
                <TD colspan="3"> <TEXTAREA name=memo rows=4 cols=45><%=memo%></TEXTAREA></TD>
              </TR>
            </table>
              <table border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td height="3" width=100></td>
                </tr>
              </table>
              <table width="622" border=0 align="center" cellPadding=4 cellSpacing=1 bgColor=#cccccc class=black-pix12 id=Table3>
                <tr > 
                  <td width="35" height="23" > <div align="center"><strong>序号</strong></div></td>
                  <td width="66"> <div align="center"><strong>商品号</strong></div></td>
                  <td width="164"> <div align="center"><strong>商品名</strong></div></td>
				  <td width="66"> <div align="center"><strong>品牌</strong></div></td>
                  <td width="51"> <div align="center"><strong>单价</strong></div></td>
                  <td width="50"> <div align="center"><strong>数量</strong></div></td>
                  <td width="85"> <div align="center"><strong>金额</strong></div></td>
                </tr>
                <%
		dim sql4,rs4,id,num,pkid
		dim sql5,rs5,model,productname,price2,price,money,a,q,allmoney,allnum
		allmoney=0
		allnum=0
		
		sql4="select * from x_view_bill_product where billid="&request.QueryString("id")&""
		set rs4=conn.execute(sql4)
		if rs4.bof or rs4.eof then
		else
		i=1
			do while not rs4.eof
			id=rs4("id")
			billid=rs4("billid")
			productpkid=rs4("productpkid")
			model=rs4("model")
			productname=rs4("productname")
			pipai=rs4("pipai")
			num=rs4("num")
			price=rs4("price")
			money=rs4("money")
			colorsize=rs4("colorsize")
			allnum = allnum + num
			allmoney = allmoney + money
					response.write "<tr bgcolor='#F1F1F1'>" 
					response.write "<td height='23' align='center'>"&i&"</td>"
					response.write "<td><a href='../show.asp?pkid="&productpkid&"' target='_blank'>"&model&"</a></td>"
					response.write "<td>"&productname&" "&colorsize&"</td>"
					response.write "<td>"&pipai&"</td>"
					response.write "<td align=right>"&formatnumber(price,2)&"</td>"
					response.write "<td align=right>"&formatnumber(num,2)&"</td>"
					response.write "<td align=right>＄"&formatnumber(money,2)&"</td>"
					response.write "</tr>"
			  m=m+1
			  rs4.movenext
			  i=i+1
			  loop
		end if
		rs4.close
		set rs4=nothing
			  %>
                <INPUT type=hidden value="<%=i-1%>" name=maxi>
                <tr> 
                  <td height=25 colspan=5 align=right><strong>合计：</strong></td>
                  <td align=right><strong><%=formatnumber(allnum,2)%></strong></td>
                  <td align=right><strong>＄<%=formatnumber(allmoney,2)%></strong></td>
                </tr>
              </table>

		    <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td colspan="2" height="28" width="65%"><b>
				合计金额：＄<%=formatnumber(total,2)%>&nbsp;&nbsp;
				优惠券抵扣金额：＄<%=formatnumber(youhjmoney,2)%>&nbsp;&nbsp;
				需支付金额：＄<%=formatnumber(lastneedpay,2)%>
				</b></td>
              </tr>
            </table>

            <table width="622" border=0 align="center" cellPadding=4 cellSpacing=1 bgColor=#cccccc class=black-pix12 id=Table3>
              <FORM id=theForm name="theForm" action="dingdan_sendnot.asp" method="post" onSubmit="return Juge()">
                <tr bgcolor='#BFDFFF'> 
                  <td height="40"> 付款状态：<select name="paystat" size="1" style="width:200px;">
                      <OPTION value="<%=paystat%>" selected><%=paystat%></OPTION>
                      <option value="已付款">已付款</option>
                    </select> &nbsp; 
					订单状态：<select name="sendornot" size="1" style="width:130px;">
                      <OPTION value="<%=sendornot%>" selected><%=sendornot%></OPTION>
                      <option value="未处理">未处理</option>
                      <option value="会员自行取消">会员自行取消</option>
                      <option value="无效定单已取消">无效定单已取消</option>
                      <option value="正在备货">正在备货</option>
                      <option value="已发货">已发货</option>
                    </select> &nbsp;
                    <input type="submit" name="Submit" value="修改状态" class="input2"> 
					<input type=hidden name='id' value="<%=request.QueryString("id")%>"> 
					<input type=hidden name='username' value="<%=username1%>"> 
					<input type=hidden name='allmoney' value="<%=allmoney%>">                  </td>
                </tr>
                <tr bgcolor='#EeEAEA'> 
                  <td height="27">
			    发货日期：<input name="senddate" type="text" value="<%=senddate%>"> <input type="button"  name="Submit" value="…" onClick="getDayDay(document.all.senddate)"  onBlur="if(document.all.senddate.value=='null')document.all.senddate.value='<%=senddate%>';">                </tr>
                <tr bgcolor='#EeEAEA'> 
                  <td height="30"><br>
					<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					<input type="submit" name="Submitwl" value="保存" class="input2">				  </td>
                </tr>
              </form>
            </table>
          </td>
        </tr>
      </table>


    </td>
  </tr>

<TR><TD height=50>

</TD></TR>
</table>
<br>
<script language="JavaScript">
<!--
function Juge()
{

	if (theForm.sendornot.value == "已发货")
	{
	  
		  if (theForm.senddate.value == "")
		  {
			alert("请您填写发货日期!");
			document.theForm.senddate.focus();
			return false;
		  }
		
		  if (theForm.sendcom.value == "")
		  {
			alert("请您填写物流公司!");
			document.theForm.sendcom.focus();
			return false;
		  }
		  
		  if (theForm.sendno.value == "")
		  {
			alert("请您填写物流发货单号!");
			document.theForm.sendno.focus();
			return false;
		  }

		
		   if(confirm("你现在要保存吗?"))
			  return true
		   else
			  return false;
	}
}
//-->
</script>
</body>
</html>

