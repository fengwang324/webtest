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
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
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
-->
</style></head>

<body bgcolor="#fcfcfc">

<form name="theForm" method="post" action="b_kinddel.asp">

  <%
nowpage=request("page")
menuid=request("menuid")

sql="select *,( select iif(IsNull(count(id)),0,count(id)) from e_product where e_product.kind=a.kindnum ) as sumpro from sh_kind a order by kindnum"

set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof or rs.eof then
	response.write "<div align=center>当前没有记录。"
	call actionbutton(menuid,"3","b_kindadd.asp?menuid="&request("menuid"))
	response.write "</div>"
else
%>
<table width="650" border="0" align="center" cellpadding="0" cellspacing="0" id="tabcss">
<tr bgcolor="#BFDFFF">
<td><b>商品类别列表</b></td>
<td><b> <a href='b_kind.asp'>点这查看类别一览表</a> &nbsp;&nbsp;&nbsp; <a href='b_kindadd.asp'>点这新增类别</a></b></td>
</tr></table>

<table width="650" border="0" align="center" cellpadding="0" cellspacing="0" id="tabcss">
  <tr bgcolor="#f1f1f1">
    <td width="30" height="22">&nbsp;</td>
    <td width="150"> <div align="left">分类编号</div></td>
    <td width="240"> <div align="left">分类名称</div></td>
	<td width="50"> <div align="left">&nbsp;</div></td>
	<td width="70"> <div align="left">相关商品数</div></td>
    </tr>
    <%
   rs.pagesize=20
   rs.absolutepage=1
   if request("page")<>"" then rs.absolutepage=request("page")
   rowcount =rs.pagesize
   if not rs.eof then
   		do while not rs.eof and rowcount>0 
			id=rs("id")
    		kindnum=rs("kindnum")
			kindname=rs("kindname")
			sumpro=rs("sumpro")
			
			grade=len(kindnum)/4
			'p=1
			gradestr=""
			'for p=1 to grade-1
				gradestr=gradestr & String(grade-1, "—")
			'next
			
  			response.write "<tr Onmouseover='return SetBgColor(this);' Onmouseout='return RestoreBgColor(this);'>"
    		response.write "<td align=center><input type='checkbox' name='checkbox"&m&"' value='"&kindnum&"'></td>"
			response.write "<td><a href='b_kindshow.asp?id="&id&"'>"&kindnum&"</a></td>"
			
			if grade=1 then
    		response.write "<td>|—"&gradestr&"<font color=#ff0000><b>"&kindname&"</b></font></td>"
			response.write "<td><a href='b_kindadd.asp?kindnum_up="&kindnum&"'><font color=#ff0000><b>新增下级</b></font></a></td>"
			elseif grade=2 then
			response.write "<td>|—"&gradestr&"<font color=#ff0000>"&kindname&"</font></td>"
			else
			response.write "<td>|—"&gradestr&kindname&"</td>"
			response.write "<td>&nbsp;</td>"
			end if
			if cstr(sumpro)="0" then
			response.write "<td>&nbsp;</td>"
			else
			response.write "<td>&nbsp;<a href='productlist.asp?kind="&kindnum&"' target='_blank'>"&sumpro&"</a></td>"
			end if
  			response.write "</tr>"&vbcrlf
			rs.movenext
		m = m + 1
		rowcount=rowcount-1
		loop
		end if
%>
</table>
<%
	dim pagerecord,allrecord,allpage,addbuttonurl,backandnexturl
	pagerecord=20
	allrecord=rs.recordcount
	allpage=rs.pagecount
	addbuttonurl="b_kindadd.asp?menuid="&menuid
	backandnexturl="b_kind2.asp?menuid="&menuid
	
	call page(pagerecord,allrecord,allpage,backandnexturl,addbuttonurl)

end if
rs.close
set rs=nothing
conn.close

response.cookies("nowpage_cook")=nowpage
%>
<br><br><br>
</body>
</html>

