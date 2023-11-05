<!--#include file="conn.asp"-->
<!--#include file="inc_product_list.asp"-->
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
</style>
<script language="javascript">
<!--
function SetBgColor(Menu)
{
		Menu.style.background="#F3F3F5";
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
<TABLE width="1020" border=0 align="center" cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
  <tr>
    <td> 



<table width="1020" height="352" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
    <td width="210" height="184" valign="top">
		<!-- #include file="inc_left_all.asp" -->
      </td>
	 <FORM id=theForm name="theForm" action="orderupdate.asp" method="post" onSubmit="return Juge()">
    <td width="780" valign="top">
	<table width="99%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height=2></td>
      </tr>
    </table>

              <TABLE width="780" border=0 align="center" cellPadding=0 cellSpacing=0>
                <TBODY>
                  <TR> 
                    <TD width=546 height=50 background=images/buystep2.jpg> 
					<table width="538" border="0" cellspacing="0" cellpadding="0">
                        <tr> 
                          <td width="193" height="22">&nbsp;</td>
                          <td width="113"> <div align="center">浏览商品</div></td>
                          <td width="112"> <div align="center"><strong><font color="#FFFFFF">购物车</font></strong></div></td>
                          <td width="120"> <div align="center">结算</div></td>
                        </tr>
                      </table></TD>
                    <TD width="144">&nbsp;</TD>
                    <TD width=90>&nbsp;</TD>
                  </TR>
                </TBODY>
              </TABLE>
              <TABLE width=780 border=0 align="center" cellPadding=0 cellSpacing=0 style="margin-top:8px;">
                <TBODY>
                  <TR> 
                    <TD align=left> <TABLE height=36 cellSpacing=0 cellPadding=0 width=200 border=0>
                        <TBODY>
                          <TR> 
                            <TD width=10><IMG height=36 src="images/gwc_6.gif" 
                width=10></TD>
                            <TD width=30 background=images/gwc_10.gif><IMG height=36 
                  src="images/gwc_8.gif" width=25></TD>
                            <TD class=zi14 vAlign=center align=left 
                background=images/gwc_10.gif><A class=zi href=javascript:window.history.back(-2)><b>返回上页 继续购物</b></A></TD>
                            <TD width=10><IMG height=36 src="images/gwc_12.gif" 
                  width=10></TD>
                          </TR>
                        </TBODY>
                      </TABLE></TD>
                    <TD><IMG height=32 src="images/gwc_14.gif" width=155></TD>
                    <TD align=right> <TABLE height=36 cellSpacing=0 cellPadding=0 width=200 border=0>
                        <TBODY>
                          <TR> 
                            <TD width=10><IMG height=36 src="images/gwc_6.gif" width=10></TD>
                            <TD width=30 background=images/gwc_10.gif><IMG height=36 src="images/gwc_16.gif" width=28></TD>
                            <TD class=zi14 vAlign=center align=left background=images/gwc_10.gif>&nbsp;<A class="lv" href="orderjs.asp"><b>就买这些 进结算中心</b></A></TD>
                            <TD width=10><IMG height=36 src="images/gwc_12.gif" width=10></TD>
                          </TR>
                        </TBODY>
                      </TABLE></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <TABLE width=770 align="center" cellPadding=3 cellSpacing=0 style="BORDER-RIGHT: #dddddd 1px solid; BORDER-TOP: #dddddd 1px solid; BORDER-LEFT: #dddddd 1px solid; BORDER-BOTTOM: #dddddd 1px solid">
                <TBODY>
                  <TR> 
                    <TD height="173" align=middle vAlign=top> 
					<DIV align=left> </DIV>
                      <table width="760" border="0" align="center" cellpadding="2" cellspacing="1">
                        <tr bgcolor="#CCCCCC"> 
                          <td width="40" height="23" > 
							<div align="center"><font color="#666666"><strong>序号</strong></font></div></td>
                          <td width="80"> 
                            <div align="center"><font color="#666666"><strong>&nbsp;</strong></font></div></td>
                          <td width="90"> 
                            <div align="center"><font color="#666666"><strong>商品号</strong></font></div></td>
                          <td width="240"> 
                            <div align="center"><font color="#666666"><strong>商品名</strong></font></div></td>
                          <td width="70"> 
                            <div align="center"><font color="#666666"><strong>单价</strong></font></div></td>
                          <td width="60"> 
                            <div align="center"><font color="#666666"><strong>数量</strong></font></div></td>
                          <td width="80"> 
                            <div align="center"><font color="#666666"><strong>金额</strong></font></div></td>
                          <td width="60"> 
                            <div align="center"><font color="#666666"><strong>编辑</strong></font></div></td>
                        </tr>
						<%
						dim sql4,rs4,id,num,pkid
						dim sql5,rs5,model,productname,price2,price,money,a,q,allmoney,allnum
						allmoney=0
						allnum=0
						i=1
						sql4="select id,num,pkid,colorsize from x_order where notemp='"&request.Cookies("notemp")&"'"
						set rs4=conn.execute(sql4)
						if rs4.bof or rs4.eof then
						else
							do while not rs4.eof
							id=rs4("id")
							num=rs4("num")
							pkid=rs4("pkid")
							colorsize=rs4("colorsize")
									sql5="select kind,model,productname,smallpicpath,price"&session("customkind")&" from e_product where pkid="&pkid
									set rs5=conn.execute(sql5)
									if rs5.bof or rs5.eof then
									else
										model=rs5("model")
										productname=rs5("productname")
										smallpicpath=rs5("smallpicpath")
										kind=rs5("kind")
										price2=rs5("price"&session("customkind"))
										price2=trim(price2)
										if price2<>"" then
											q=1
											for  q=1  to  len(price2)
												  a=mid(price2,q,1)
												  if a<"0"  or  a>"9" then 
													if a<>"." then exit for
												  end if
											next
											price=left(price2,q-1)
											money = num * price
											allmoney = allmoney + money
											allnum = allnum + num
										end if
					
										response.write "<tr bgcolor='#FaFaFa'>" 
										response.write "<td height='23' align='center'>"&i&"</td>"
										response.write "<td align='center'><a href='show.asp?pkid="&pkid&"' target='_blank'><IMG src='product/"&smallpicpath&"' height=60 border='0'></A></td>"
										response.write "<td><a href='show.asp?pkid="&pkid&"' target='_blank'><font color=#0033FF>"&model&"</font></a></td>"
										response.write "<td><a href='show.asp?pkid="&pkid&"' target='_blank'>"&productname&" "&colorsize&"</a></td>"
										response.write "<td align=right>"&formatnumber(price,2)&"</td>"
										response.write "<td align=center><input type=text name='num"&i&"' value='"&num&"' size='3'><input type=hidden name='id"&i&"' value='"&id&"'><input type=hidden name='colorsize"&i&"' value='"&colorsize&"'></td>"
										response.write "<td align=right>＄"&formatnumber(money,2)&"</td>"
										response.write "<td align=center><input class='input3' name='delbu"&i&"' type='button' value='删除' onclick=""if(!confirm('是否从购物车中删除“"&productname&"”？'))return false;location.href='orderpro_del.asp?id="&id&"'; ""></td>"
										response.write "</tr>"&vbcrlf
										m=m+1
									end if
									rs5.close
							  rs4.movenext
							  i=i+1
							  loop
						end if
						%>
                        <tr bgcolor="#dddddd"> 
							<td height=25 colspan=5 align=right><strong>合计：</strong></td>
                          <td align=right><strong><%=formatnumber(allnum,2)%></strong></td>
                          <td align=right><strong>＄<%=formatnumber(allmoney,2)%></strong></td>
                          <td>&nbsp;</td>
                        </tr>
                      </table>
                      <table width="770" border="0" align="center" cellpadding="0" cellspacing="0">
                        <TR bgColor=#ffffff> 
						  <TD  height="46" align=middle colSpan=6> 
							<INPUT type="button" name=delall value=" " style="BACKGROUND: url(images/gwc_28.gif) no-repeat; border:0;width:101px;height:29px;CURSOR:hand;" onClick="if(!confirm('确认清空购物车吗?')) return false;location.href='orderpro_del.asp?kind=qingkong';">
							&nbsp;&nbsp;&nbsp; 
							<INPUT type="submit" name=update value=" " style="background-image: url(images/gwc_26.gif);border:0;width:101px;height:29px;CURSOR:hand;">
							&nbsp;&nbsp;&nbsp;  
							<%if i>1 then%>
							<INPUT type="button" name=df value=" " style="background-image: url(images/gwc_jiesuan.gif);border:0;width:101px;height:29px;CURSOR:hand;" onClick="javascript:window.location.href='orderjs.asp';">
							&nbsp;&nbsp;&nbsp; 
							<%else%>
							<INPUT type="button" name=df value=" " style="background-image: url(images/gwc_jiesuan.gif);border:0;width:101px;height:29px;CURSOR:hand;" onClick="alert('您没有选购任何商品，请选购了商品再结算。');">
							&nbsp;&nbsp;&nbsp; 
							<%end if%>
							&nbsp;<INPUT type=hidden value="<%=i-1%>" name=maxi> 
						  </TD>
                        </TR>
                      </table>
                      
					  
                    </TD>
                  </TR>
                </TBODY>
              </TABLE>

<%if kind<>"" then
s_col = 4

	'相关商品
	sql_new="select top "&s_col * 2&" pkid,model,productname,smallpicpath,price1,price"&session("customkind")&",pipai,addtime from view_product where kind like '"&left(kind,4)&"%' and pkid<>"&pkid&" order by pkid desc "  
	set rs_new=server.createobject("adodb.recordset")
	rs_new.open sql_new,conn,1,1
	if rs_new.bof or rs_new.eof then
		response.write "<td>此栏暂时没有商品记录！</td>"
	else
%>
	<br><br>
	<TABLE width=780 border=0 align="center" cellPadding=1 cellSpacing=0  class="weitab">
		<tr>
		  <td height="32" colspan=6 background="images/index_png24_a1.png">
		  <table border=0 width="100%"><tr><td><font color="#999999"><strong>购买上面的商品，通常还会购买下面商品：</strong></font></td></tr></table>
		  </td>
		</tr>
		<TR> 
<%
	i=1
	do while not rs_new.eof
	pkid=rs_new("pkid")
	model=rs_new("model")
	productname=rs_new("productname")
	smallpicpath=rs_new("smallpicpath")
	price1=rs_new("price1")
	price2=rs_new("price"&session("customkind"))
	pipai=rs_new("pipai")
	addtime=rs_new("addtime")

	call product_list(1)
	
	if i mod s_col =0 then
		response.write "</tr><tr><td height=8>&nbsp;</td></tr><tr>"
		i=1
	else
		i=i+1
	end if
	rs_new.movenext
	loop
%>
	  </TR>
   </TABLE>
<%
	end if
	rs_new.close
	set rs_new=nothing

end if
%>


              <br></td>
		</form>
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
  if (parseFloat(theForm.num<%=aa%>.value)<0)
  {
	alert("请正确填写数量!");
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

