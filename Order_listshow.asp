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
</style></head>

<body>
<!-- #include file="head.asp" -->
<TABLE width="770" border=0 align="center" cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
  <tr><td>


<table width="960" height="352" border="0" cellpadding="0" cellspacing="0">
<tr> 
          <td height="184" valign="top"> <table width="99%"  border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td height=2></td>
              </tr>
            </table>
            <TABLE width="95%" border=0 align="center" cellPadding=0 cellSpacing=2 class=a1>
              <TBODY>
                <TR> 
                  <TD width="548" height=25 colSpan=3>
<TABLE width="715" border=0 cellPadding=0 cellSpacing=0>
<TBODY>
                        <TR> 
                          <TD width="575" 
                            height=25><strong>订单 <font color=#ff0000><%=request.QueryString("billno")%></font> 
                            商品列表：</strong></TD>
                          <TD width=140> 
                            <INPUT class="input2" onClick="javascript:window.history.back();" type=button value="返回订单列表" name=delall></TD>
                        </TR>
                      </TBODY>
                    </TABLE>
                  </TD>
                </TR>
              </tbody>
            </table>
            <table width="95%" border="0" align="center" cellpadding="2" cellspacing="1">
              <tr bgcolor="#5C9153"> 
                <td width="53" height="23" > 
<div align="center"><font color="#FFFFFF"><strong>序号</strong></font></div></td>
                <td width="97"> 
                  <div align="center"><font color="#FFFFFF"><strong>商品号</strong></font></div></td>
                <td width="250"> 
                  <div align="center"><font color="#FFFFFF"><strong>商品名</strong></font></div></td>
                <td width="76"> 
                  <div align="center"><font color="#FFFFFF"><strong>单价</strong></font></div></td>
                <td width="76"> 
                  <div align="center"><font color="#FFFFFF"><strong>数量</strong></font></div></td>
                <td width="149"> 
                  <div align="center"><font color="#FFFFFF"><strong>金额</strong></font></div></td>
              </tr>
              <%
		dim sql4,rs4,id,num,pkid
		dim sql5,rs5,model,productname,price2,price,money,a,q,allmoney,allnum
		allmoney=0
		allnum=0
		
		sql4="select * from x_bill_product where billid="&request.QueryString("id")&""
		set rs4=conn.execute(sql4)
		if rs4.bof or rs4.eof then
		else
		i=1
			do while not rs4.eof
			id=rs4("id")
			num=rs4("num")
			pkid=rs4("productpkid")
			price=rs4("price")
			money=rs4("money")
			
					sql5="select model,productname,price2 from e_product where pkid="&pkid
					set rs5=conn.execute(sql5)
					if rs5.bof or rs5.eof then
					else
						model=rs5("model")
						productname=rs5("productname")
						allmoney = allmoney + money
						allnum = allnum + num
	
						response.write "<tr bgcolor='#F1F1F1'>" 
						response.write "<td height='23' align='center'>"&i&"</td>"
						response.write "<td><a href='show.asp?pkid="&pkid&"'><font color=#0033FF>"&model&"</font></a></td>"
						response.write "<td><a href='show.asp?pkid="&pkid&"'>"&productname&"</a></td>"
						response.write "<td align=right>"&formatnumber(price,2)&"</td>"
						response.write "<td align=center><input type=text name='num"&i&"' value='"&num&"' size='4'><input type=hidden name='id"&i&"' value='"&id&"'></td>"
						response.write "<td align=right>＄"&formatnumber(money,2)&"</td>"
						response.write "</tr>"
						m=m+1
					end if
					rs5.close
			  rs4.movenext
			  i=i+1
			  loop
		end if
			  %>
              <tr bgcolor="#649D67"> 
                <td height=25 colspan=4 align=right><font color="#FFFFFF"><strong>合计：</strong></font></td>
                <td align=right><font color="#FFFFFF"><strong><%=formatnumber(allnum,2)%></strong></font></td>
                <td align=right><font color="#FFFFFF"><strong>＄<%=formatnumber(allmoney,2)%></strong></font></td>
              </tr>
            </table>
            <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
              <TR bgColor=#ffffff> 
                <TD  height="32" align=middle colSpan=6> <INPUT class="input2" onClick="javascript:window.history.back();" type=button value="返回订单列表" name=delall>
                </TD>
              </TR>
            </table></td>
        </tr>
      </table>


    </td>
  </tr>

<TR><TD>

</TD></TR>
</table>

<SCRIPT language=JavaScript>
<!--
function Juge()
{
<%for aa=1 to i-1%>

   if (theForm.num<%=aa%>.value == "")
  {
    alert("请填写数量。");
    theForm.num<%=aa%>.focus();
    return (false);
  }
	  if (isNaN(theForm.num<%=aa%>.value))
	  {
		alert("数量要为数字!");
		theForm.num<%=aa%>.focus();
		return (false);
	  }
<%next%>
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

