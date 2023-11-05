<!--#include file="conn.asp"-->
<%
Response.Buffer = True 
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
%>
<%
if session("username")="" then
	conn.close
	set conn=nothing
	response.Redirect("hyzq.asp")
	response.end
end if
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
.style12 {color: #990000}
-->
</style></head>

<body>
<!-- #include file="head.asp" -->
<TABLE width="1020" border=0 align="center" cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
  <tr>
    <td> 


<table width="1020" height="352" border="0" align="center" cellpadding="0" cellspacing="0">
<tr> 
          <td height="184" valign="top" bgcolor="#F7F7F7"> 
<table width="958"  border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td height=1></td>
              </tr>
            </table>
            <TABLE width="1018" border=0 align="center" cellPadding=0 cellSpacing=2 class=a1>
              <TBODY>
                <TR> 
                  <TD height=30 colSpan=3>
				  <TABLE width="100%" border=0 cellPadding=0 cellSpacing=0 background="images/bg4.gif">
<TBODY>
                        <TR> 
                          <TD width="621" height=30> <strong>我的订单：</strong></TD>
                          <TD width="137"><a href="javascript:window.history.back()">&lt;&lt; 
                            返回</a></TD>
                        </TR>
                      </TBODY>
                    </TABLE></TD>
                </TR>
              </tbody>
            </table>
            <table width="1018" border="0" align="center" cellpadding="2" cellspacing="1">
              <tr bgcolor="#5C9153"> 
                <td width="32" height="25" > 
					<div align="center"><font color="#FFFFFF"><strong>序号</strong></font></div></td>
                <td width="115"> 
                  <div align="center"><font color="#FFFFFF"><strong>订单号</strong></font></div></td>
                <td width="105"> 
                  <div align="center"><font color="#FFFFFF"><strong>订单时间</strong></font></div></td>
                <td width="53"> 
                  <div align="center"><font color="#FFFFFF"><strong>总金额</strong></font></div></td>
                <td width="53"> 
                  <div align="center"><font color="#FFFFFF"><strong>得积分</strong></font></div></td>
                <td width="53"> 
                  <div align="center"><font color="#FFFFFF"><strong>用户名</strong></font></div></td>
                <td width="57"> 
                  <div align="center"><font color="#FFFFFF"><strong>收货人</strong></font></div></td>
                <!--
				<td width="90"> 
                  <div align="center"><font color="#FFFFFF"><strong>电话</strong></font></div></td>
				-->
                <td width="65"> 
                  <div align="center"><font color="#FFFFFF"><strong>送货方式</strong></font></div></td>
                <td width="65"> 
                  <div align="center"><font color="#FFFFFF"><strong>付款方式</strong></font></div></td>
                <td width="70"> 
                  <div align="center"><font color="#FFFFFF"><strong>付款状态</strong></font></div></td>
				  
                <td width="65"> 
                  <div align="center"><font color="#FFFFFF"><strong>处理状态</strong></font></div></td>
                <td width="50">
                  <div align="center"><font color="#FFFFFF"><strong>编辑</strong></font></div></td>
              </tr>
              <%
		dim sql4,rs4,id,num,pkid
		dim sql5,rs5,model,productname,price2,price,money
		
		kind=request("kind")
		nowpage=request("page")

		sql4="select * from x_view_bill where username='"&session("username")&"' order by id desc"
		set rs4=server.createobject("adodb.recordset")
		rs4.open sql4,conn,1,1
		if rs4.bof or rs4.eof then
			response.write "没有商品记录！"
		else
		i=1
		rs4.pagesize=15
		rs4.absolutepage=1
		if request("page")<>"" then rs4.absolutepage=request("page")
		rowcount=rs4.pagesize
		
			do while not rs4.eof and rowcount>0
				id=rs4("id")
				billno=rs4("billno")
				billdate=rs4("billdate")
				username=rs4("username")
				truename=rs4("truename")
				'telphone1=rs4("telphone1")
				sendornot=rs4("sendornot")
				
				lastneedpay=rs4("lastneedpay")
				sendtype=rs4("sendtype")
				paytype=rs4("paytype")
				paystat=rs4("paystat")
				billzhf=rs4("billzhf")
				if sendornot="已发货" then
				else
				billzhf="&nbsp;"
				end if
		
				response.write "<tr bgcolor='#Ffffff' onMouseOver=""this.style.background='#f1f1f1';"" onMouseOut=""this.style.background='#ffffff';"">"
				response.write "<td height='26' align='center'>"&i&"</td>"
				response.write "<td><a style='CURSOR:hand;' onclick=""window.open('order_print.asp?id="&id&"&billno="&billno&"','','top=55,left=130,scrollbars=yes,width=680,height=500,resizable=yes,menubar=yes,resizable=yes')""><font color=#0033FF>"&billno&"</font></a></td>"
				response.write "<td>"&billdate&"</td>"
				response.write "<td align=right>"&formatnumber(lastneedpay,2)&"</td>"
				response.write "<td align=right>"&billzhf&"</td>"
				if session("m_gold")>"1" then
				response.write "<td><font color='#ff0000'><b>"&username&"</b></font></td>"
				else
				response.write "<td>"&username&"</td>"
				end if
				response.write "<td>"&truename&"</td>"
				'response.write "<td>"&telphone1&"</td>"
				response.write "<td>"&sendtype&"</td>"
				response.write "<td>"&paytype&"</td>"
				
				if (paytype="支付宝" or paytype="网银在线" or paytype="财付通" or paytype="PayPal贝宝") and (paystat="" or isnull(paystat)) and (sendornot<>"会员自行取消" and sendornot<>"无效定单已取消") then
				response.write "<td style='width:100px'><a href='orderjs3.asp?billno="&billno&"&nowpay=1' target='_blank'><font color='#0000ff'>现在去付款</font></a></td>"
				else
				response.write "<td style='width:100px'>"&paystat&"</td>"
				end if
				
				if sendornot="已发货" then
				response.write "<td align=center><a title='点击查看发货信息' style='CURSOR:hand;' onclick=""window.open('order_print.asp?id="&id&"&billno="&billno&"','','top=55,left=130,scrollbars=yes,width=680,height=500,resizable=yes,menubar=yes,resizable=yes')""><font color=#0033FF>已发货</font></a></td>"
				else
				response.write "<td align=center>"&sendornot&"</td>"
				end if
				if paystat="已付款" then
				response.write "<td>&nbsp;</td>"
				elseif sendornot="未处理" and (not instr(paystat,"付款成功")>0 or paystat="" or isnull(paystat)) then
				response.write "<td align=center><a onclick=""return confirm('是否取消订单"&billno&"?')"" href=""order_listdel.asp?id="&id&""">取消</a></td>" 
				else
				response.write "<td>&nbsp;</td>"
				end if
				
				response.write "</tr>"&vbcrlf
				rs4.movenext
				i=i+1
				rowcount=rowcount-1
			loop
		end if
			  %>
            </table>
            <table width="1016" border="0" align="center" cellpadding="0" cellspacing="0">
              <TR bgColor=#ffffff> 
                <TD  height="32" align=middle colSpan=6>
<%
if nowpage="" then nowpage=1
nowpath=request.servervariables("script_name")
kind=server.URLEncode(kind)
response.write "<table width='550' align='center'>"
  response.write "<tr> "
   response.write " <td height='22' align=right style='font-size:9pt;'> "
		response.write "共"&rs4.recordcount&"条记录&nbsp;每页15条&nbsp;现在 "&nowpage&"/"&rs4.pagecount &" 页&nbsp;"	
        response.write "<input type='button' name='Button3' value='首页' onclick=window.location.href='"&nowpath&"?kind="&kind&"&none='>&nbsp;"
        if clng(nowpage)>1 then
        response.write "<input type='button' name='Button4' value='上页' onclick=window.location.href='"&nowpath&"?page="&nowpage-1&"&kind="&kind&"&none='>&nbsp;"
        else
        response.write "<input type='button' name='Button4' value='上页'>&nbsp;"
        end if
		if clng(nowpage)<>clng(rs4.pagecount) then
        response.write "<input type='button' name='Button5' value='下页' onclick=window.location.href='"&nowpath&"?page="&nowpage+1&"&kind="&kind&"&none='>&nbsp;"
        else
        response.write "<input type='button' name='Button5' value='下页'>&nbsp;"
		end if
		response.write "<input type='button' name='Button6' value='末页' onclick=window.location.href='"&nowpath&"?page="&rs4.pagecount&"&kind="&kind&"&none='>"
    response.write "</td>"
  response.write "</tr>"
response.write "<tr><td height=60>&nbsp;</td></tr></table>"
%>
				</TD>
              </TR>
			  
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

