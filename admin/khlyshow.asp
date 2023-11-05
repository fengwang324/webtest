<!--#include file="conn.asp"-->
<SCRIPT language=JavaScript>
<!--
function CheckForm(theForm){
if (theForm.title.value.length<1)
	{alert("请填写题目！");
	 theForm.title.focus();
	 return false;
	}

 return true;
 }
//-->
</SCRIPT>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">


<html xmlns="http://www.w3.org/1999/xhtml">
<link href="main.css" rel="stylesheet" type="text/css" />
<link href="../css/css.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
body {
	background-color: #EFF7FF;
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.style3 {font-size: 16px;
	font-weight: bold;
}
.style4 {font-size: 12px}
-->
</style><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>后台系统</title>
</head>

<body>
<div align="center">
  <table width="650" border="0" cellpadding="0" cellspacing="0" class="style4">
    <tr>
    <td width="502" height="20"><font color="#0000CC">留言管理---查看详情<!--#include file="char.inc"--><%Set rs=Server.CreateObject("ADODB.Recordset")
Sqln="select * from khly where id="& request("id") &""
rs.open Sqln,conn,1,1%></font></td>
    <td width="148" align="center"><a href="khly.asp">返回</a></td>
  </tr>
</table>
<table width="80%" border="1" align="center" cellpadding="2" cellspacing="1" bordercolor="#CCCCCC" class="style4">
<tr><td width="13%" height="30" align="right">姓名：</td>
  <td width="87%" align="left"><%=rs("xingming")%>&nbsp;</td>
</tr>
<tr>
  <td height="30" align="right">电话：</td>
  <td align="left"><%=rs("dianhua")%>&nbsp;</td>
</tr>
<tr>
  <td height="30" align="right">邮箱：</td>
  <td align="left"><%=rs("email")%>&nbsp;</td>
</tr>
<tr>
  <td height="30" align="right">地址：</td>
  <td align="left"><%=rs("dizhi")%>&nbsp;</td>
</tr>
<tr>
  <td height="30" align="right">留言主题：</td>
  <td align="left"><%=rs("title")%>&nbsp;</td>
</tr>
<tr>
  <td height="30" align="right" valign="top">留言内容：</td>
  <td align="left" valign="top" style="line-height:20px; "><%=htmlencode2(rs("content"))%>&nbsp;</td>
</tr>
<tr>
  <td height="30" align="right">留言时间：</td>
  <td align="left"><%=rs("ntime")%>&nbsp;</td>
</tr>

<%rs.close
set rs=nothing
conn.close
set conn=nothing%>
</table>
</div>
</body>
</html>
