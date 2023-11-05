<!-- #include file="conn.asp" -->
<!--#include file="check2.asp"-->
<%
menuid=request("menuid")
id=request("id")



sql="select * from sh_kind where id="&request("id")
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof or rs.eof then
	response.write "<br><div align=center>请点击 <a href='b_kind.asp?menuid="&menuid&"'>返回</a> 重试。</div>"
else
	kindnum=rs("kindnum")
	kindname=rs("kindname")
	
end if
rs.close
set rs=nothing

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
  if (theForm.kindnum.value == "")
  {
    alert("请输入编号!");
    theForm.kindnum.focus();
    return (false);
  }


  if (isNaN(theForm.kindnum.value))
  {
    alert("编号要为数字!");
    theForm.kindnum.focus();
    return (false);
  }

	if (theForm.kindnum.value.length != theForm.kindnumold.value.length)
	{
    alert("本次填写的编号与原来的“位数”不同。");
    theForm.kindnum.focus();
    return (false);	
	}


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


<form name="theForm" method="post" action="b_kindmodify.asp" onSubmit="return Juge(this)">  
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0" id="tabcss">
  <tr bgcolor="#BFDFFF">
      <td colspan=2><b>修改商品类别</b></td>
    </tr>

    <tr>
      <td height="61"><div align="right">编号&nbsp;</div></td>
	  
      <td>
	      <input name="kindnumold" type="text" value="<%=kindnum%>" readonly style='BACKGROUND-COLOR: #FDEEDB;'>(原来的编号)<br>
	  <br><input name="kindnum"    type="text" value="<%=kindnum%>">
        <br>
      (编号的位数与要与原来的<b>位数</b>一样，而且位数要为<b>4的倍数</b>，修改后的编号不能和<b>已有</b>的编号一样，例如：能将0001改为0009或者0012等，但不能改为00012，又如,可以将00010005改为00010007,或改为00010018等，这样修改后，原来的下级会跟着改变的。通过这里可修改分类的<strong>顺序</strong>,也可<strong>移动分类</strong>。)</td>
    </tr>
    <tr> 
      <td width="28%" height="39"> <div align="right">类别名称&nbsp;</div></td> 
      <td width="72%"> <input name="kindname" type="text" id="kindname" value="<%=kindname%>"></td>
    </tr>

    <tr> 
	 <td height="38"> </td>
      <td height="38"> 
<input type="submit" name="Submit" value=" 修改 " >
<input name="id" type="hidden" id="id" value="<%=request("id")%>">
<%
call connclose()
%>
</td>
    </tr>
  </table>
</form>
	</body>
</html>










