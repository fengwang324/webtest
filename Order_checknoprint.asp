<!--#include file="conn.asp"-->
<%
if request.QueryString("checkno")="" then
	conn.close
	set conn=nothing
	response.Redirect("hyzq.asp")
	response.end
end if

function addspace(strfield)	'列表显示	
	if strfield="" or isnull(strfield) or strfield="未知" then 
		addspace="&nbsp;"
	else
		addspace=strfield
	end if
end function
%>
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
<title>订单明细</title>
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
</style></head>

<body>
<TABLE width="622" border=0 align="center" cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
  <tr>
    <td> 
<table width="622" height="352" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          
            <td width="557" height="184" valign="top"> 
			<table width="99%"  border="0" cellspacing="0" cellpadding="0">
              <tr> 
                  <td height=30 align='center'><b><%=application("sitename")%> 网上订单</b></td>
                </tr>
              </table>

              
            <table width="622" border=1 align="center" cellPadding=2 cellSpacing=0 bordercolordark="#CCCCCC" bordercolorlight="#ffffff" bgColor=#cccccc id=Table3>
              <%
dim sql7,rs7,username,truename,province,city,telphone1,telphone2,fax,mobile,address,postno,email
	sql7="select * from x_view_bill where id="&request.QueryString("id")&""
	set rs7=conn.execute(sql7)
	if rs7.bof or rs7.eof then
		response.write "此订单不是属于 "&session("username")&" 的 或 此订单已被删除。"
		response.end
	else
		id=rs7("id")
		username=addspace(rs7("username")) 
		truename=addspace(rs7("truename")) 
		province=addspace(rs7("province")) 
		city=addspace(rs7("city")) 
		telphone1=addspace(rs7("telphone1")) 
		telphone2=addspace(rs7("telphone2")) 
		fax=addspace(rs7("fax")) 
		mobile=addspace(rs7("mobile")) 
		address=addspace(rs7("address")) 
		postno=addspace(rs7("postno")) 
		email=addspace(rs7("email"))
		sendtype=rs7("sendtype")
		paytype=rs7("paytype")
		memo=addspace(rs7("memo"))
		sendornot=rs7("sendornot")
		billdate=rs7("billdate")
		billno=rs7("billno")
		
		sendmoney=rs7("sendmoney")
		total=rs7("total")
		paystat=rs7("paystat")
		
		senddate=rs7("senddate")
		sendcom=rs7("sendcom")
		sendno=rs7("sendno")
		sendmemo=rs7("sendmemo")
		
	end if
	rs7.close
	set rs7=nothing
if memo<>"" then
memo = replace(memo,vbcrlf,"<br>")
memo = replace(memo," ","&nbsp;")
end if
%>
              <tr> 
                <td width="111" height="23"><strong><font color="#CC0000">顾客资料：</font></strong></td>
                  <td colspan="3">&nbsp;</td>
                </tr>
                <TR bgColor=#ffffff> 
                <TD width=111 height="23" align=right class=black> 
<DIV align=right>下单用户名：</DIV></TD>
                  
                <TD width=235><%=username%> </TD>
                  
                <TD width=87> 
                  <div align="right">收货人姓名：</div></TD>
                  
                <TD width=163><%=truename%></TD>
                </TR>
                <TR bgColor=#ffffff> 
                <TD width=111 height="23" align=right class=black> 
<DIV align=right>省份：</DIV></TD>
                  
                <TD width=235><%=province%></TD>
                  
                <TD width=87> 
                  <div align="right">城市：</div></TD>
                <TD width=163><%=city%></TD>
                </TR>
                <TR bgColor=#ffffff> 
                <TD width=111 height="23" align=right class=black> 
<DIV align=right>联系电话1：</DIV></TD>
                <TD width=235><%=telphone1%></TD>
                  
                <TD width=87> 
                  <div align="right">联系电话2：</div></TD>
                <TD width=163> <%=telphone2%></TD>
                </TR>
                <TR bgColor=#ffffff> 
                <TD width=111 height="23" align=right class=black> 
<DIV align=right>移动电话：</DIV></TD>
                  
                <TD width=235><%=mobile%></TD>
                  
                <TD width=87> 
                  <div align="right">传真：</div></TD>
                  
                <TD width=163><%=fax%></TD>
                </TR>
                <TR bgColor=#ffffff> 
                <TD width=111 height="23" align=right class=black> 
<DIV align=right>送货地址：</DIV></TD>
                  <TD colspan="3" class=black><%=address%>&nbsp;&nbsp;&nbsp;邮编：<%=postno%></TD>
                </TR>
                <TR bgColor=#ffffff> 
                <TD width=111 height="23" align=right class=black> 
<DIV align=right>电子邮件：</DIV></TD>
                  <TD colspan="3" class=black><%=email%></TD>
                </TR>
              </table>
              <table border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td height="3" width=100></td>
                </tr>
              </table>
              
            <table width="622" border=1 align="center" cellPadding=2 cellSpacing=0 bordercolordark="#CCCCCC" bordercolorlight="#ffffff" bgColor=#cccccc class=black-pix12 id=Table3>
              <tr> 
                <td width="111" height="23"><font color="#CC0000"><strong>订单信息：</strong></font></td>
                <td colspan="3">&nbsp;</td>
              </tr>
              <TR bgColor=#ffffff> 
                <TD height="23" align=middle> <div align="right">订单编号：</div></TD>
                <TD><b><%=billno%></b></TD>
                <TD><div align="right">下单时间：</div></TD>
                <TD><%=billdate%></TD>
              </TR>
              <TR bgColor=#ffffff> 
                <TD align=middle width=111> 
<div align="right">送货方式：</div></TD>
                <TD width=236><%=sendtype%></TD>
                <TD width=86> 
                  <div align="right">付款方式：</div></TD>
                <TD width=163><%=paytype%></TD>
              </TR>
              <TR bgColor=#ffffff>
                <TD align=middle> <div align="right">送货费用：</div></TD>
                <TD>＄<b><%=formatnumber(sendmoney,2)%></b></TD>
                <TD align="right">付款状态：</TD>
                <TD><%=paystat%>&nbsp;</TD>
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
                <TD align=middle width=111> 
<div align="right">特别说明：</div></TD>
                <TD colspan="3"><%=memo%></TD>
              </TR>
            </table>
              <table border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td height="3" width=100></td>
                </tr>
              </table>
              
            <table width="622" border="1" align="center" cellpadding="2" cellspacing="0" bordercolordark="#CCCCCC" bordercolorlight="#ffffff">
              <tr bgcolor="#CCCCCC"> 
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
			pipai=addspace(rs4("pipai"))
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
              <tr bgcolor="#CCCCCC"> 
                <td height=25 colspan=5 align=right><strong>合计：</strong></td>
                <td align=right><strong><%=formatnumber(allnum,2)%></strong></td>
                <td align=right><strong>＄<%=formatnumber(allmoney,2)%></strong></td>
              </tr>
            </table>
		    <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td height="35" width="65%">&nbsp;<b>合计金额：＄<%=formatnumber(total,2)%></b></td>
                <td height="35" width="35%">&nbsp;</td>
              </tr>
            </table>
		    <b>发货信息：</b> <br>
           发货日期：<%=senddate%>
            　　物流公司：<%=sendcom%>
            　　发货单号：<%=sendno%> <br>
            备注说明：<%=sendmemo%></td>
        </tr>
      </table>


    </td>
  </tr>

<TR><TD>

</TD></TR>
</table>

</body>
</html>

