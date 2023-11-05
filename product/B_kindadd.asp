<!-- #include file="conn.asp" -->
<!--#include file="check2.asp"-->
<%
kindnum_up=request.QueryString("kindnum_up")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link REL='stylesheet' HREF='szcrm.css' TYPE='text/css'>
<SCRIPT language=JavaScript>
<!--

function Juge(theForm)
{
  if (theForm.kindname.value == "")
  {
    alert("请输入类别名称!");
    theForm.kindname.focus();
    return (false);
  }
}
-->
</script>
<style type="text/css">
<!--
.style1 {
	color: #FF0000;
	font-weight: bold;
}
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
</style>
</head>
<body bgcolor="#fcfcfc">


<form name="theForm" method="post" action="b_kindaddadmin.asp" onSubmit="return Juge(this)">

<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0" id="tabcss">
  <tr bgcolor="#BFDFFF">
      <td height=28 colspan=2 class=titlehead><b>增加商品分类</b></td>
    </tr>

    <tr > 
	<td height="38"> <div align="right">隶属的上级分类&nbsp;</div></td>
      <td> 
        <select name="kindnum" size="20" style="width:200px; font-size:14px; ">
<%
	  dim rs,sql,maxkind,kindnum,kindname,grade,i,gradestr

	response.write "<option value=''>##没有上级##</option>"

		sql="select * from sh_kind order by kindnum"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
		if rs.bof or rs.eof then
		else
		do while not rs.eof
			kindnum=rs("kindnum")
			kindname=rs("kindname")
			grade=len(kindnum)/4
			'p=1
			gradestr=""
			'for p=1 to grade-1
				gradestr=gradestr & String(grade-1, "—")
			'next
			if kindnum_up=kindnum then
			response.write "<option value='"&kindnum&"' selected>|—"&gradestr&kindname&"</option>"
			else
			response.write "<option value='"&kindnum&"'>|—"&gradestr&kindname&"</option>"
			end if
		rs.movenext
		loop
		end if
		rs.close
		set rs=nothing
%>
	  
      </select>
        (请在左边中选择将要新增的分类的上级。如果是第一级选择“没有上级”)</td>
    </tr>
    <tr> 
<td width="24%" height="39"> 
<div align="right">类别名称&nbsp;</div></td>
      <td width="76%"> 
        <input name="kindname" type="text" id="kindname" style="width:200px; "></td>
    </tr>

    <tr> 
	<td height="34"> </td>
      <td height="34"> 

  <input type="submit" name="Submit" value=" 保存 ">
          <%'call Actionbutton(request("menuid"),"3","1")%>
          <%
conn.close

response.write "&nbsp;&nbsp;&nbsp;&nbsp;<a href='b_kind.asp?page="&request.cookies("nowpage_cook")&"'>返回一览表</a>"
response.write "&nbsp;&nbsp;&nbsp;&nbsp;<a href='b_kind2.asp?page="&request.cookies("nowpage_cook")&"'>返回列表</a>"
%>

          </td>
    </tr>
  </table>
</form>
</body>
</html>

