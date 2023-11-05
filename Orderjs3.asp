<!--#include file="conn.asp"-->
<%
Response.Buffer = True 
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
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
.style13 {color: #333333}
.style15 {font-size: 14px; color: #669933; font-weight: bold; }
-->
</style></head>

<body>
<!-- #include file="head.asp" -->
<TABLE width="1020" border=0 align="center" cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
  <tr>
    <td> 



<table width="1020" height="352" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
    <td width="210" height="184" valign="top">
		<!-- #include file="inc_left_all.asp" -->
      </td>
    <td width="750" valign="top">
	<table width="99%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height=2></td>
      </tr>
    </table>
            <TABLE width="700" border=0 align="center" cellPadding=0 cellSpacing=2 class=a1>
              <TBODY>
                <TR> 
                  <TD height=30 colSpan=3><TABLE width="100%" border=0 cellPadding=0 cellSpacing=0>
                      <TBODY>
                        <TR> 
                          <TD width="539" 
                            height=30 background="images/buystep.jpg"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                              <tr> 
                                <td width="35%" height="22">&nbsp;</td>
                                <td width="23%"> 
                                  <div align="center">填写交易信息</div></td>
                                <td width="19%"> 
                                  <div align="center"><strong><font color="#FFFFFF">生成订单</font></strong></div></td>
                                <td width="23%"> <div align="center">网上支付或汇款</div></td>
                              </tr>
                            </table></TD>
                          <TD width=157 height="50">&nbsp;</TD>
                        </TR>
                      </TBODY>
                    </TABLE></TD>
                </TR>
              </tbody>
            </table>
            <table width="680" border="0" align="center" cellpadding="0" cellspacing="0">

              <tr> 
                <td height="40" style="font-size:14px;font-weight:bold;"> <%
				billno=trim(request.QueryString("billno"))
				sql="select * from x_bill where billno='"&billno&"'"
				set rs=server.createobject("adodb.recordset")
				rs.open sql,conn,1,1
				if rs.bof or rs.eof then
				else
					billno=rs("billno")
					billdate=rs("billdate")
					sendtype=rs("sendtype")
					paytype=rs("paytype")
					allmoney=rs("allmoney")
					sendmoney=rs("sendmoney")
					total=rs("total")
					
					youhjmoney=rs("youhjmoney")
					lastneedpay=rs("lastneedpay")
					'-----------------------
					truename=rs("truename")
					address=rs("address")
					telphone1=rs("telphone1")
					telphone2=rs("telphone2")
					postno=rs("postno")
					email=rs("email")
					mobile=rs("mobile")
					memo=rs("memo")
				end if
				rs.close
				set rs=nothing
				%> 
                  <table width="100%" height="263" border="0" cellpadding="0" cellspacing="0" background="images/success.gif">
<tr> 				<tr>
                      <td height="76"><font style="font-size:14px;font-weight:bold;">&nbsp;</font></td>
                      <td height="76" colspan="2"><font style="font-size:14px;font-weight:bold;">
					  <%if request.QueryString("nowpay")<>"1" then%>恭喜，您的订单已成功提交。<%end if%>订单号为：<font face="Arial, Helvetica, sans-serif" color=red style="font-size:14pt;font-weight:bold;"><%=request.QueryString("billno")%></font></font></td>
                    </tr>
                    <tr> 
                      <td width="10%" height="25"><div align="left"></div></td>
                      <td width="45%"> 
                        <div align="left">下单日期时间：<%=billdate%></div></td>
                      <td width="45%" rowspan="6">&nbsp;</td>
                    </tr>
                    <tr> 
                      <td height="25"><div align="left"></div></td>
                      <td>您选定的送货方式：<%=sendtype%></td>
                    </tr>
                    <tr> 
                      <td height="25"><div align="left"></div></td>
                      <td>您选定的支付方式：<%=paytype%></td>
                    </tr>
                    <tr> 
                      <td height="25"><div align="left"></div></td>
                      <td>货款金额：＄<b><%=formatnumber(allmoney,2)%></b></td>
                    </tr>
                    <tr> 
                      <td height="25"><div align="left"></div></td>
                      <td>送货费用：＄<b><%=formatnumber(sendmoney,2)%></b></td>
                    </tr>
                    <tr> 
                      <td height="25"><div align="left"></div></td>
                      <td>优惠券抵扣金额：＄<b><%=formatnumber(youhjmoney,2)%></b></td>
                    </tr>
                    <tr> 
                      <td height="35"><div align="left"></div></td>
                      <td><div align="left"><font color="#FF0000" style="font-size:14px;font-weight:bold;">该订单，您需支付金额：</font><font face="Arial, Helvetica, sans-serif" color=red style="font-size:14pt;font-weight:bold;">＄<%=formatnumber(lastneedpay,2)%></font></div></td>
                    </tr>
                  </table></td>
              </tr>
              <%
			  if paytype="支付宝" then
                  if request.QueryString("nowpay")<>"1" then 
					sql4="select * from sh_pay1 where paykind='支付宝'"
					set rs4=server.createobject("adodb.recordset")
					rs4.open sql4,conn,1,1
					if rs4.bof or rs4.eof then				
					else
						myaccount=rs4("myaccount")
						mykey=rs4("mykey")
						myparentid=rs4("myparentid")
						response.Cookies("paybillno")=billno  		'用于alipay_notify.asp中写入数据库 '2010-11-19已修改.alipay_notify.asp将request.Cookies("paybillno")改为out_trade_no
						response.Cookies("subject")=billno					'subject		=	"测试订单号"		'商品名称，如果客户走购物车流程可以设为  "订单号："&request("客户网站订单")
						response.Cookies("body")=application("sitename")&"订单"&billno 				'body			=	"商品描述"		'商品描述
																			'out_trade_no   =   dingdan         '按时间获取的订单号，可以修改成自己网站的订单号，保证每次提交的都是唯一即可
						response.Cookies("price")=lastneedpay '因为有优惠券,所以支付抵扣后的金额  total						'price		    =	"0.01"			'商品单价			0.01～100000000.00  ，注：不要出现3,000.00，价格不支持","号
																			'quantity       =   "1"             '商品数量,如果走购物车默认为1
																			'discount       =   "0"             '商品折扣
																			
																			
						response.Cookies("seller_email")=myaccount		    'seller_email    =    seller_email   '卖家的支付宝帐号，c2c客户，可以更改此参数。
						
						
						response.Cookies("show_url")="http://"&request.ServerVariables("HTTP_HOST") 	'show_url       = ""     '商户网站的网址。 
																										'seller_email	= ""	 '请填写签约支付宝账号，
						response.Cookies("partner")=myparentid											'partner		= ""	 '填写签约支付宝账号对应的partnerID，
						response.Cookies("key")=mykey													'key			= ""	 '填写签约账号对应的安全校验码	
						
						thisurl=request.ServerVariables("url")
						thisurl=replace(thisurl,"/orderjs3.asp","")
						response.Cookies("notify_url")="http://"&request.ServerVariables("HTTP_HOST")&thisurl&"/pay/Alipay_Notify.asp"    		  'notify_url		= "./Alipay_Notify.asp"	        '交易过程中服务器通知的页面 要用 http://格式的完整路径，例如http://www.alipay.com/alipay/Alipay_Notify.asp  注意文件位置请填写正确。
						response.Cookies("return_url")="http://"&request.ServerVariables("HTTP_HOST")&thisurl&"/pay/return_Alipay_Notify.asp"     'return_url		= "./return_Alipay_Notify.asp"	'付完款后跳转的页面			

			  %>
              <form name="payform" method="post" action="pay/index.asp">
                <tr> 
                  <td height="40" style="font-size:14px;line-height:180%;">&nbsp;&nbsp; 
                    <input type="button" name="Submit" value="点击通过“支付宝”完成网上支付" class="input2" style="width:240px;height:24px;" onClick="pay()"> 
                  </td>
                </tr>
              </form>
              <SCRIPT language=JavaScript>
						<!--
						function pay()
						{
						payform.Submit.disabled=true;
						document.payform.target = '_blank';
						document.payform.submit();   
						}
						-->
						</script>
              <%
					end if 
					else	
			%>
<form name="payform" method="post" action="pay/index.asp">
                <tr> 
                  <td height="40" style="font-size:14px;line-height:180%;">&nbsp;&nbsp; 
                    <input type="button" name="Submit" value="点击通过“支付宝”完成网上支付" class="input2" style="width:240px;height:24px;"  onclick="javascript:alert('免费版程序不支持订单后期支付功能(仅支持下完订单后立即支付)，正式版本无限制，官网：Http://www.Sunsnu.com')" > 
                  </td>
                </tr>
              </form>			<%
					end if 			
			  elseif paytype="网银在线" then
					sql4="select * from sh_pay1 where paykind='网银在线'"
					set rs4=server.createobject("adodb.recordset")
					rs4.open sql4,conn,1,1
					if rs4.bof or rs4.eof then
					
					else
						myaccount=rs4("myaccount")
						mykey=rs4("mykey")
						response.Cookies("paybillno")=billno
						
						response.Cookies("v_mid")=myaccount
						response.Cookies("key")=mykey
						thisurl=request.ServerVariables("url")
						thisurl=replace(thisurl,"/orderjs3.asp","")
						response.Cookies("v_url")="http://"&request.ServerVariables("HTTP_HOST")&thisurl&"/chinabank/Receive.asp"
			  %>
              <FORM name="payform2" action="chinabank/Send.asp" method="post">
                <TR>
                  <TD height="40" style="font-size:14px;line-height:180%;"> <input name="v_oid" type="hidden" maxlength="64" value="<%=billno%>"> 
                    <input name="v_rcvname" type="hidden" value="<%=truename%>"> 
                    <input name="v_rcvaddr" type="hidden" id="v_rcvaddr"  value="<%=address%>"> 
                    <input name="v_rcvtel" type="hidden" id="v_rcvtel"  value="<%=telphone1&" "&telphone2%>"> 
                    <input name="v_rcvpost" type="hidden" id="v_rcvpost"  value="<%=postno%>"> 
                    <input type="hidden" name="v_rcvemail" value="<%=email%>"> 
                    <input type="hidden" name="v_rcvmobile" value="<%=mobile%>"> 
                    <input name="remark1" type=hidden id="remark1" value="<%=memo%>"> 
                    <input name="v_ordername" type="hidden" id="v_ordername" value="<%=truename%>"> 
                    <input name="v_orderaddr" type="hidden" id="v_orderaddr"  value="<%=address%>"> 
                    <input name="v_ordertel" type="hidden" id="v_ordertel"  value="<%=telphone1&" "&telphone2%>"> 
                    <input name="v_orderpost" type="hidden" id="v_orderpost"  value="<%=postno%>"> 
                    <input name="v_orderemail" type="hidden" id="v_orderemail" value="<%=email%>"> 
                    <input name="v_ordermobile" type="hidden" id="v_ordermobile" value="<%=mobile%>"> 
                    <input name="remark2" type=hidden id="remark2" value="<%=memo%>"> 
                    <input name="v_amount" type=hidden value="<%=round(lastneedpay,2)%>"> 
                    &nbsp;&nbsp;
                    <input type="button" name="Submit2" value="点击通过“网线在线”完成网上支付"  class="input2" style="width:240px;height:24px;" onClick="pay2()"> 
                  </TD>
                </TR>
              </form>
              <SCRIPT language=JavaScript>
					<!--
					function pay2()
					{
					payform2.Submit2.disabled=true;
					document.payform2.target = '_blank';
					document.payform2.submit();   
					}
					-->
					</script>
              <%
			  		end if
					rs4.close 
					set rs4=nothing
					
			   
			  elseif paytype="货到付款" then
			  %>
              <tr> 
                <td height=50><b><font color='#E10000'>订购已成功，我们会尽快与您联系，并将商品送给收货人，并收取货款及附加金额。</font></b></td>
              </tr>
              <tr> 
                <td height=35><a href='order_list.asp'><font color='#0000ff'>点击这里查看您的订单&gt;&gt;</font></a></td>
              </tr>
			  <%elseif paytype="账户余额" then%>
              <tr> 
                <td height=50><b><font color='#E10000'>订单成功保存，并成功支付。我们会尽快与您联系，并将商品送给收货人。欢迎下次光临。</font></b></td>
              </tr>
              <tr> 
                <td height=35><a href='order_list.asp'><font color='#0000ff'>点击这里查看您的订单&gt;&gt;</font></a></td>
              </tr>
              <%end if%>
              <tr> 
                <td height=85 style="line-height:180%;"> <b><font color='#E10000'>如果选择了银行汇款、邮局汇款，或在线支付不成功，可以到下面相应的银行或邮局汇款。<br>
				汇款后，请电话告知我们您的“订单号”、“汇款金额”及“汇款银行”。<br>
				</font></b><font color='#333333'>请在您的电汇单“事由”一栏处注明您的姓名、订单号。 我们收到您的汇款之后，将立即安排发货。</font></td>
              </tr>
              <tr> 
                <td height=61>
				<table width="99%" align="center" style='BORDER-LEFT: #cccccc 1px solid;BORDER-RIGHT: #cccccc 1px solid;BORDER-TOP: #cccccc 1px solid;  BORDER-BOTTOM: #cccccc 1px solid;'>
					<%
					sql4="select * from sh_pay2 where len(bankno)>5 "
					set rs4=server.createobject("adodb.recordset")
					rs4.open sql4,conn,1,1
					if rs4.bof or rs4.eof then
					else
					%>
					<TR bgColor=#EEEEEE> 
					   <TD height=28><strong><font color="#FF9900">可以汇款的银行帐号：</font></strong></TD>
					</TR>				
					<%
						do while not rs4.eof
						id=rs4("id")
						bankname=rs4("bankname")
						bankno=rs4("bankno")
						bankren=rs4("bankren")
						bankaddr=rs4("bankaddr")
						bankmemo=rs4("bankmemo")
					%>
					<TR bgColor=#ffffff onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
					  <TD height=25><b><%=bankname%></b>&nbsp;&nbsp;帐号：<%=bankno%>&nbsp;&nbsp;户名：<%=bankren%>&nbsp;&nbsp;开户银行：<%=bankaddr%></TD>
					</TR>
					<%
						rs4.movenext
						loop
					end if
					rs4.close
					set rs4=nothing
					%>


				<%
				sql4="select * from sh_pay4 where youstat='1' "
				set rs4=server.createobject("adodb.recordset")
				rs4.open sql4,conn,1,1
				if rs4.bof or rs4.eof then
				else
				%>
                <TR bgColor=#eeeeee>
				   <TD height=28><strong><font color="#FF9900">邮局汇款</font></strong></TD>
                </TR>				
				<%
					youpaykind=rs4("youpaykind")
					memo1=rs4("memo")&"&nbsp;"
					memo1 = replace(memo1,vbcrlf,"<br>")
					memo1 = replace(memo1," ","&nbsp;")
				%>
                <TR bgColor=#ffffff > 
                  <TD height=80 style="line-height:180%;"><%=memo1%></TD>
                </TR>
				<%
				end if
				rs4.close
				set rs4=nothing
				%>
				</table>
				</td>
              </tr>
              <tr> 
                <td height=40><a href='order_list.asp'><font color='#0000ff'> 
                  点击这里查看您的订单&gt;&gt;</font></a></td>
              </tr>
              <%					  		
			  'end if
			  %>
              <tr> 
                <td><br>
                  您可以在 我的订单 中查看或取消订单； 
                  <p>在7日内我们会保留未支付的订单，请及时支付。<br>
                    <br>
                    <br>
                    <br>
                  </p></td>
              </tr>
            </table>



          </td>
  </tr>
</table>


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



