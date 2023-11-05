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
.style11 {
	font-size: 14px;
	color: #666666;
	font-weight: bold;
}
.style12 {color: #990000}
.STYLE13 {
	color: #999999;
	font-weight: bold;
}
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
<%
kind=request("kind")
nowpage=request("page")
cx=request("cx")
hot=request("hot")

searchkind=trim(request("searchkind"))
keyword=trim(request("keyword"))

 
dim display_mode
if request.querystring("display_mode")<>"" then 
	response.cookies("display_mode_cook")=request.querystring("display_mode")
	response.cookies("display_mode_cook").Expires="2025-2-28"
	display_mode=request.querystring("display_mode")
elseif request.cookies("display_mode_cook")<>"" then
	display_mode=request.cookies("display_mode_cook")
else
	display_mode="grid"
end if
%>
</head>

<body>
<!-- #include file="head.asp" -->
<br>
		<TABLE width="538" border=0 align="center" cellPadding=0 cellSpacing=2 class=a1>
		  <TBODY>
			<TR> 
			  <TD height=50 background=images/buystep0.jpg align="center">
				<table width="538" border="0" cellspacing="0" cellpadding="0" align="center">
				  <tr> 
					<td width="193" height="22">&nbsp;</td>
					<td width="113"> <div align="center"><strong><font color="#FFFFFF">瀏覽商品</font></strong></div></td>
					<td width="112"> <div align="center">購物車</div></td>
					<td width="120"> <div align="center">結算</div></td>
				  </tr>
				</table>
			  </TD>
			</TR>
		  </TBODY>
		</TABLE>
<TABLE width="1020" border=0 align="center" cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
  <tr>
    <td> 


	<TABLE width="1020" align="center" border=0 cellPadding=0 cellSpacing=0 background="images/bg4.gif" style="margin-top:8px;">
	  <TBODY>
		<TR>
				  <TD width=100 height=30 background=images/hotnewpro.gif> &nbsp; &nbsp; &nbsp;&nbsp;<STRONG><FONT color=#ff6600 style='font-size:14px;'><%
					if cx="1" then
						response.write "促銷商品列表"
					elseif hot="1" then
						response.write "熱賣商品列表"
					else
						response.write "商品列表"
					end if
					%></FONT></STRONG></TD>
				  <TD width="406">&nbsp;</TD>
				  <TD width="73">&nbsp;</td>
				   <td width="96">&nbsp;</td>
					<td valign="top">
					
					</td>
					<td valign="top">
					<!--
					<input type="image" name="imageField" src="images/bnt_go.gif" style="border:0;height:18;width:36" alt="go"/>
					-->	
					</TD>
		</TR>
	  	</TBODY>
      </TABLE>
<%

sql_condition=" where updown='1' "
if kind<>"" then
	sql_condition = sql_condition & " and ( kind like '"&kind&"%' or kind2 like '"&kind&"%' ) "
end if
if cx<>"" then
	sql_condition = sql_condition & " and good='1' "
end if
if hot<>"" then
	sql_condition = sql_condition & " and hot='1' "
end if
if searchkind<>"" then
	sql_condition = sql_condition & " and kind like '"&searchkind&"%' "
end if
if keyword<>"" then
	sql_condition = sql_condition & " and (model like '%"&keyword&"%' or productname like '%"&keyword&"%' or pipai like '%"&keyword&"%') "
end if


order_name=" pkid desc "

sql="select pkid,model,productname,smallpicpath,price1,price"&session("customkind")&",kindname,pipai,addtime from view_product "&sql_condition&" order by "&order_name

set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof or rs.eof then
	response.write "暫時沒有符合條件商品記錄！<br><br><br><br>"
else


i=1
%>
                          <TABLE align=center border=0 cellPadding=2 cellSpacing=0 width="100%">
                            <TBODY>
                              <TR> 
                                
<%
	rs.pagesize=s_col*s_row
	rs.absolutepage=1
	if request("page")<>"" then rs.absolutepage=request("page")
	rowcount=rs.pagesize
	if not rs.eof then
		do while not rs.eof and rowcount>0
		pkid=rs("pkid")
		model=rs("model")
		productname=rs("productname")
		smallpicpath=rs("smallpicpath")
		price1=rs("price1")
		price2=rs("price"&session("customkind"))
		pipai=rs("pipai")
		kindname=rs("kindname")
		addtime=rs("addtime")

if display_mode="grid" then	
	
	call product_list(2)	
%>

<%elseif display_mode="list" then %>
			<TD vAlign=top align=middle>
			<div align="center" style="padding-top:3px;padding-left:2px;padding-right:3px;"> 
				<A href="show.asp?pkid=<%=pkid%>" <%=newwindow%>> <IMG src="product/<%=smallpicpath%>" width='55' border="0" class=img1></A>
			</div>
			</td>
			<td>
			<DIV align="left" style="MARGIN-TOP:1px;width:330px;">
				<A class=font_link href="show.asp?pkid=<%=pkid%>" <%=newwindow%> style="font-weight:bold;color:#666666;"><%=productname%></A>
			</DIV>
			<DIV align="left" style="MARGIN-TOP: 2px;">
			<%if s_model="是" then%>
			商品號：<font style="width:98px;"><%=model%></font>
			<%end if%>
			<%if s_pipai="是" then%>
			品　牌：<%=pipai%>
			<%end if%>
			</div>				
			<DIV align="left" style="MARGIN-TOP: 2px;">
			<%=session("customkindname")%>價：<font class='nowprice' style="width:80px;">＄<%=price2%></font>市場價：<FONT class='oldprice' style="width:70px;"><del>＄<%=price1%></del></FONT>
			上市時間：<%=addtime%>
			</DIV>
			</td>
			<td>
			<a href="javascript:window.external.AddFavorite('http://#<%=request.ServerVariables("SCRIPT_NAME")%>?<%=request.serverVariables("Query_string")%>', '<%=application("sitename")%>-<%=productname%>')"><img src="images/button-shoucang.gif" border="0"></a>
			</td>
			<td>
			<DIV align="center" style="MARGIN-TOP: 5px; WIDTH: <%=s_picwidth%>px">
			<%if s_colorsize="是" then%>
			<a href="show.asp?pkid=<%=pkid%>" <%=newwindow%>><img src="images/icon_car.gif" border=0></a>
			<%else%>
			<a href="show.asp?pkid=<%=pkid%>&num=1&tk=shop7z" <%=newwindow%>><img src="images/buy.gif" border=0></a>
			<%end if%>
			</div>
			</TD>
 
			</div>				
			<DIV align="left" style="MARGIN-TOP: 2px;">
			<%=session("customkindname")%>價：<font class='nowprice' style="width:80px;">＄<%=price2%></font>市場價：<FONT class='oldprice' style="width:70px;"><del>＄<%=price1%></del></FONT>
			上市時間：<%=addtime%>
			</DIV>
			</td>
			<td>
			<a href="javascript:window.external.AddFavorite('http://#<%=request.ServerVariables("SCRIPT_NAME")%>?<%=request.serverVariables("Query_string")%>', '<%=application("sitename")%>-<%=productname%>')"><img src="images/button-shoucang.gif" border="0"></a>
			</td>
			<td>
			<DIV align="center" style="MARGIN-TOP: 5px; WIDTH: <%=s_picwidth%>px">
			<%if s_colorsize="是" then%>
			<a href="show.asp?pkid=<%=pkid%>" <%=newwindow%>><img src="images/icon_car.gif" border=0></a>
			<%else%>
			<a href="show.asp?pkid=<%=pkid%>&num=1&tk=shop7z" <%=newwindow%>><img src="images/buy.gif" border=0></a>
			<%end if%>
			</div>
			</TD>
<%
end if

			if display_mode="grid" then
				if i=s_col then
					response.write "</tr><tr><td height=8></td></tr><tr>"
					i=1
				else
					i=i+1
				end if
			else
				i=i+1
				if i mod 2 = 0 then
				response.write "</tr><tr><td height=1></td></tr><tr bgcolor='#f5f5f5'>"
				else
				response.write "</tr><tr><td height=1></td></tr><tr>"
				end if
				
			end if

			
	
		rs.movenext
		rowcount=rowcount-1
		loop
	end if

%>
							   
                              </TR>
                            </TBODY>
                          </TABLE>
						   <table><TR> 
							<TD height="5"></TD>
						  </TR></table>

<%
if nowpage="" then nowpage=1
nowpath=request.servervariables("script_name")
'kind=server.URLEncode(kind)
'searchkind=trim(request("searchkind"))
'keyword=trim(request("keyword"))
myurl="kind="&kind&"&cx="&cx&"&hot="&hot&"&searchkind="&searchkind&"&keyword="&keyword
response.write "<table width='650' align='center'>"
  response.write "<tr> "
   response.write " <td width='500' height='22' align=right style='font-size:9pt;'> "
		response.write "共"&rs.recordcount&"條記錄&nbsp;每頁"&s_col*s_row&"条&nbsp;現在 "&nowpage&"/"&rs.pagecount &" 页&nbsp;"	
        response.write "<input type='button' name='Button3' value='' style='background-image:url(images/z1.jpg);width:37px;height:19px;border:0px' onclick=window.location.href='"&nowpath&"?"&myurl&"&none='>&nbsp;"
        if clng(nowpage)>1 then
        response.write "<input type='button' name='Button4' value='' style='background-image:url(images/z2.jpg);width:45px;height:19px;border:0px' onclick=window.location.href='"&nowpath&"?page="&nowpage-1&"&"&myurl&"&none='>&nbsp;"
        else
        response.write "<input type='button' name='Button4' value='' style='background-image:url(images/z2.jpg);width:45px;height:19px;border:0px'>&nbsp;"
        end if
		if clng(nowpage)<>clng(rs.pagecount) then
        response.write "<input type='button' name='Button5' value='' style='background-image:url(images/z3.jpg);width:45px;height:19px;border:0px' onclick=window.location.href='"&nowpath&"?page="&nowpage+1&"&"&myurl&"&none='>&nbsp;"
        else
        response.write "<input type='button' name='Button5' value='' style='background-image:url(images/z3.jpg);width:45px;height:19px;border:0px'>&nbsp;"
		end if
		response.write "<input type='button' name='Button6' value='' style='background-image:url(images/z4.jpg);width:37px;height:19px;border:0px' onclick=window.location.href='"&nowpath&"?page="&rs.pagecount&"&"&myurl&"&none='>"
    response.write "</td><form name='form29' method='post' action='"&nowpath&"?"&myurl&"&none='><td width='150'>&nbsp;轉到<input type='text' name='page' size=3 style='width:30px; height:18px;font-size:12pt;font-weight:bold;font-family: Arial;'>頁 <input type='submit' name='Submitpage' value='GO' style='width:30px; height:22px; color:#ffffff; background-color:#999999;font-size:10pt;font-weight:bold;font-family: Arial; '></td></form>"
  response.write "</tr>"
response.write "</table>"

end if
rs.close
set rs=nothing

%> 



    </td>
  </tr>

</table>
<%
myurl2=myurl&"&page="&nowpage
'display_mode()放在程序的下面,主要是为了得到本页的地址myurl2
%>
<script language="javascript">
<!--
function display_mode(Menu)
{
	if(Menu=="list")
	{
		window.location.href="<%=nowpath&"?display_mode=list&"&myurl2%>";
	}
	if(Menu=="grid")
	{
		window.location.href="<%=nowpath&"?display_mode=grid&"&myurl2%>";
	}
	if(Menu=="text")
	{
		window.location.href="<%=nowpath&"?display_mode=text&"&myurl2%>";
	}
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



