<!-- #include file="conn.asp" -->
<!--#include file="check2.asp"-->
<!-- #include file="ActionButton.asp" -->
<!-- #include file="page.asp" -->
<%


dim nowpage,sql,rs,rowcount,myorder
dim id,productkind
dim m
m=1
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link REL='stylesheet' HREF='szcrm.css' TYPE='text/css'>
<script language="javascript">
<!--
function SetBgColor(Menu)
{
		Menu.style.background="#E7E7E7";
}
function RestoreBgColor(Menu)
{
		Menu.style.background="#fafafa";
}
-->
</script>
<style type="text/css">
<!--
body {
	margin-top: 5px;
}
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
A {
	FONT-SIZE: 12px; TEXT-DECORATION: none
}
A:visited {
	COLOR: #3333ff; FONT-SIZE: 12px; TEXT-DECORATION: none
}

A:hover {
	COLOR: #ff3366; BACKGROUND-COLOR: #f7f7f7; TEXT-DECORATION: underline
}
-->
</style>

<style type="text/css">
	.a2 { border:#CCCCCC solid 1px; }
	.a2 div {CLEAR: both; height:26px; padding-top:6px; padding-left:6px;}
	.g1 { background-color:#dddddd;  }
	.g2 { background-color:#f7f7f7;  }
	.a2 UL { LIST-STYLE-TYPE: none;FLOAT: left;width:185px; height:23px; PADDING-RIGHT: 0px; PADDING-LEFT:6px; PADDING-BOTTOM: 0px; PADDING-TOP: 0px; MARGIN: 0px; line-height:25px;  }
</style>
</head>

<body bgcolor="#fcfcfc">

  <%
nowpage=request("page")
menuid=request("menuid")

sql="select *,( select iif(IsNull(count(id)),0,count(id)) from e_product where e_product.kind=a.kindnum ) as sumpro from sh_kind a order by kindnum"

set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof or rs.eof then
	response.write "<div align=center>当前没有记录！"
	call actionbutton(menuid,"3","b_kindadd.asp?menuid="&request("menuid"))
	response.write "</div>"
else



%>
<table width="750" border="0" align="center" cellpadding="0" cellspacing="0" id="tabcss">
<tr bgcolor="#BFDFFF">
<td><b>商品类别一览表</b></td>
<td><b><a href='b_kind2.asp'>点这去类别列表</a>&nbsp;&nbsp;&nbsp; <a href='b_kindadd.asp'>点这新增类别</a></b></td>
</tr></table>
<table width="750" border="0" align="center" cellpadding="8" cellspacing="0" class="a2">
  <tr >
    <td>
    <%
		line=1
		t=1
		if not rs.eof then
		do while not rs.eof
			id=rs("id")
			kindnum=rs("kindnum")
			kindname=rs("kindname")
			sumpro = rs("sumpro")
			indexshow = rs("indexshow")
			
			grade=len(kindnum)/4

			if grade=1 then
				response.write "<div class='g1'><img src='../images/t06.gif' border=0 align='absmiddle'> <a href='b_kindshow.asp?id="&id&"'><font color=#FF3366><b>"&kindnum&" "&kindname&"</b></font></a>&nbsp;(<a href='productlist.asp?kind="&kindnum&"' target='_blank'>"&sumpro&"</a>) &nbsp;&nbsp;&nbsp;<a href='b_kindadd.asp?kindnum_up="&kindnum&"'><font color=#ff0000><b>新增下级</b></font></a>"
				response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;该类首页显示："
				if indexshow="1" then
				response.write "<font color='#ff0000'><b>是</b></font>"
				else
				response.write "否"
				end if
				response.write "&nbsp;&nbsp;&nbsp;&nbsp;设置：<a href='b_kindset.asp?kindnum="&kindnum&"&v=1'>显示</a> &nbsp; <a href='b_kindset.asp?kindnum="&kindnum&"&v=0'>不显示</a> </div>"&vbcrlf
			elseif grade=2 then
				response.write "<div class='g2'><img src='../images/li.gif' border=0 align='absmiddle'>  <a href='b_kindshow.asp?id="&id&"'><font color=#444666><b>"&kindnum&" "&kindname&"</b></font></a>&nbsp;(<a href='productlist.asp?kind="&kindnum&"' target='_blank'>"&sumpro&"</a>) &nbsp;</div>"&vbcrlf
			else
				
			end if
			
		rs.movenext
		loop
		end if
	%>
	</td></tr>
	<tr><td>共 <%response.write rs.recordcount%> 条记录 <a href='b_kind2.asp'>点这去类别列表</a>&nbsp;&nbsp;&nbsp; <a href='b_kindadd.asp'>点这新增类别</a></td></tr>
	</table>
<%
	
end if
rs.close
set rs=nothing
conn.close

%>
<br><br><br>
</body>
</html>

