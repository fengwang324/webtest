<!-- #include file="conn.asp" -->
<!--#include file="check2.asp"-->
<%


dim nowpage,sql,rs,rowcount,myorder
dim id,productkind
dim m
m=1
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link REL='stylesheet' HREF='szcrm.css' TYPE='text/css'>
<script language="javascript">
<!--
function SetBgColor(Menu)
{
		Menu.style.background="#E7E7E7";
}
function RestoreBgColor(Menu)
{
		Menu.style.background="#ffffff";
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
	line-height: 22px;
	color: #333333;
	border-right-width: 1px;
	border-bottom-width: 1px;
	border-right-style: solid;
	border-bottom-style: solid;
	border-right-color: #dddddd;
	border-bottom-color: #dddddd;
}
-->
</style><title>商品分类</title></head>

<body bgcolor="#fcfcfc">

<form name="theForm" method="post" action="b_kinddel.asp">

  <%
nowpage=request("page")
mykind=request("mykind")

sql="select * from sh_kind order by kindnum"

set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof or rs.eof then
	response.write "<div align=center>当前没有记录。"
	response.write "</div>"
else
%>

<table width="450" border="0" align="center" cellpadding="0" cellspacing="0" id="tabcss">
  <tr><td colspan=2 height="28">请点击下面的分类编号选择商品类别： <a href=javascript:window.close()>关闭窗口</a></td></tr>
  <tr bgcolor="#BFDFFF">
	  
    <td width="150"> <div align="left">分类编号</div></td>
      
    <td width="240"> <div align="left">分类名称</div></td>
    </tr>
    <%

   		do while not rs.eof 
			id=rs("id")
    		kindnum=rs("kindnum")
			kindname=rs("kindname")
			whokindadd=rs("whoaddkind")
			kindname2=replace(kindname," ","")
			
			grade=len(kindnum)/4
			'p=1
			gradestr=""
			'for p=1 to grade-1
				gradestr=gradestr & String(grade-1, "—")
			'next
			kindname2="—"&gradestr&kindname2
			
  			response.write "<tr Onmouseover='return SetBgColor(this);' Onmouseout='return RestoreBgColor(this);'>"
			
			if kindnum=mykind then
			response.write "<td><a href=javascript:pick('"&kindnum&"','"&kindname2&"')><font color='#ff0000'>"&kindnum&"</font></a> **</td>"
			else
			response.write "<td><a href=javascript:pick('"&kindnum&"','"&kindname2&"')>"&kindnum&"</a></td>"
			end if
			if grade=1 then
    		response.write "<td>|—"&gradestr&"<font color=#ff0000><b>"&kindname&"</b></font></td>"
			elseif grade=2 then
			response.write "<td>|—"&gradestr&"<font color=#ff0000>"&kindname&"</font></td>"
			else
			response.write "<td>|—"&gradestr&kindname&"</td>"
			end if
  			response.write "</tr>"
			rs.movenext
		m = m + 1
		rowcount=rowcount-1
		loop
%>
<tr><td colspan=2 align=center><a href=javascript:window.close()>关闭窗口</a></td></tr>
</table>
<%

end if
rs.close
set rs=nothing
conn.close
%>
<br><br>
<SCRIPT LANGUAGE="JavaScript"> 
<!-- 
function pick(symbol1,symbol2) { 
if (window.opener && !window.opener.closed) 
window.opener.document.theForm.kind<%=request.querystring("i")%>.value = symbol1; 
window.opener.document.theForm.kindname<%=request.querystring("i")%>.value = symbol2; 
window.close(); 
} 
// --> 
</SCRIPT>
</body>
</html>

