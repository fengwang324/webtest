<!-- #include file="conn.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv=Content-Type content=text/html; charset=utf-8>
<link REL='stylesheet' HREF='szcrm.css' TYPE='text/css'>
<SCRIPT language=JavaScript>
<!--

function Juge(theForm)
{
  if (theForm.oldpassword.value == "")
  {
    alert("请输入旧密码!");
    theForm.oldpassword.focus();
    return (false);
  }
    if (theForm.newpassword.value == "")
  {
    alert("请输入新密码!");
    theForm.newpassword.focus();
    return (false);
  }
    if (theForm.newpassword.value.length>12 || theForm.newpassword.value.length<6)
  {
    alert("密码要为6-12位!");
    theForm.newpassword.focus();
    return (false);
  } 
  
  if (theForm.newpassword.value != theForm.newpassword2.value)
  {
    alert("确认密码和新密码不一致!");
    theForm.newpassword.focus();
    return (false);
  }
  
}
-->
</script>
<style type="text/css">
<!--
input {
      	height:23px;BORDER-RIGHT: #999999 1px solid; BORDER-TOP: #999999 1px solid; FONT-SIZE: 10pt; BORDER-LEFT: #999999 1px solid; BORDER-BOTTOM: #999999 1px solid; FONT-FAMILY: "宋体"; BACKGROUND-COLOR: #ffffff
}
div{FONT-SIZE: 9pt;}

body {
	margin-left: 0px;
	margin-top: 5px;
	margin-bottom: 0px;
	margin-right: 0px;
	background-color:#fcfcfc
}
.style1 {font-size: 14px}
.style12 {color: #990000}
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
	padding-left: 3px;
	padding-right: 2px;
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
<br>
<form name="theForm" method="post" action="passwordmodifyadmin.asp" onSubmit="return Juge(this)">  
<table width="600" border="0" align="center" cellpadding="0" cellspacing="0" id="tabcss">
    <tr bgColor=#f2f2fd> 
      <td height=27 colspan="2"><font color="#FF0000"><strong>修改网站后台管理密码：</strong></font></td>
    </tr>

    <tr> 
      <td height="26"> <div align="right">帐号(用户名)&nbsp;</div></td>
      <td>&nbsp;<%=session("m_user")%> 
        <input name="username" type="hidden" id="username" value="<%=session("m_user")%>"></td>
    </tr>
    <tr> 
      <td width="33%" height="26"> <div align="right">旧密码&nbsp;</div></td>
      <td width="67%"> <input name="oldpassword" type="password" id="oldpassword"> 
      </td>
    </tr>
    <tr> 
      <td height="26"> <div align="right">新密码&nbsp;</div></td>
      <td> <input name="newpassword" type="password" id="newpassword">
        (6-12位英文、数字)</td>
    </tr>
    <tr> 
      <td height="26"> <div align="right">确认密码&nbsp;</div></td>
      <td> <input name="newpassword2" type="password" id="newpassword2">
        (6-12位英文、数字)</td>
    </tr>

    <tr> 
      <td height="40">&nbsp;</td>
	  <td>
	  <input type="submit" name="Submit" value=" 修 改 " style="background-color:#dddddd">&nbsp;&nbsp; <input type="reset" name="Submit2" value=" 清 除 "  style="background-color:#dddddd"></td>
    </tr>
  </table>
</form>
	</body>
</html>











