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
.tableborder{
	BORDER-BOTTOM: #5BD77D 1px solid; BORDER-LEFT: #5BD77D 1px solid; BORDER-RIGHT: #5BD77D 1px solid; 
	BORDER-TOP: #5BD77D 1px solid;FONT-FAMILY: "宋体"; FONT-SIZE: 9pt;
}
.tablehead
{
	background-color:#B4EDC2; COLOR:#FFFFFF; text-align:center;   FONT-SIZE: 10.5pt;color :#FFFFFF;
    FONT-FAMILY: Arial,宋体，Times New Roman
}
.titlehead
{
	background-color:#B4EDC2; COLOR:#FFFFFF; text-align:left;   FONT-SIZE: 9pt;color :#333333;
    FONT-FAMILY: Arial,宋体，Times New Roman
}
-->
</style>
<SCRIPT language=javascript>
function Echeck()
{
	if (document.formlogin.username.value=="")
	{
		alert ("请输入用户登陆名称")
		document.formlogin.username.focus();
		return false;
	}
	
	if (document.formlogin.password.value=="")
	{
		alert ("请输入用户登陆密码")
		document.formlogin.password.focus();
		return false;
	}
}
</SCRIPT>

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
    if (theForm.newpassword.value.length>20 || theForm.newpassword.value.length<4)
  {
    alert("密码要为4-12位!");
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
</head>

<body>
<!-- #include file="head.asp" -->
<TABLE width="1020" border=0 align="center" cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
  <tr><td>

<TABLE height=278 cellSpacing=0 cellPadding=0 width=1020 align=center 
bgColor=#ffffff border=0>
        <TBODY>
          <TR> 
            <TD vAlign=top width="180"> 
<TABLE cellSpacing=0 cellPadding=0 width=174 border=0>
<TBODY>
                  <TR vAlign=top align=middle> 
                    <TD colSpan=3><IMG height=13 src="images/zuo_bg_1.gif" 
            width=172></TD>
                  </TR>
                  <TR vAlign=top align=middle> 
                    <TD align=right width=12 background=images/zuo_bg_5.gif>　</TD>
                    <TD align=middle width=151><TABLE height=200 cellSpacing=0 cellPadding=0 width="85%" 
            align=center>
                        <TBODY>
                          <TR align=middle> 
                            <TD align=middle><div align="center"><img src="images/member_stor_logo.gif" width="149" height="61"></div></TD>
                          </TR>
                          <TR align=middle>
                            <TD height="8"></TD>
                          </TR>
                          <TR align=middle> 
                            <TD><IMG height=16 src="images/favadd.gif" 
                  width=16>&nbsp;<A href="hyzq.asp">&nbsp;<SPAN 
                  class=style4>个人资料</SPAN></A></TD>
                          </TR>
                          <TR align=middle> 
                            <TD vAlign=center align=middle background=images/hui.gif 
                height=15></TD>
                          </TR>
                          <TR align=middle> 
                            <TD><IMG height=16 src="images/TOP2.gif" width=16>&nbsp;
							<A href="order_all.asp?lookorder=1"><SPAN class=style4>我的购物车</SPAN></A></TD>
                          </TR>
                          <TR align=middle> 
                            <TD vAlign=center align=middle background=images/hui.gif 
                height=15></TD>
                          </TR>
                          <TR align=middle> 
                            <TD><IMG height=16 src="images/TOP2.gif" 
                  width=16>&nbsp;<A class=style4 
                  href="order_list.asp">&nbsp;<SPAN 
                  class=style4>我的订单</SPAN></A></TD>
                          </TR>
                          <TR align=middle> 
                            <TD vAlign=center align=middle background=images/hui.gif 
                height=15></TD>
                          </TR>
                          <TR align=middle> 
                            <TD vAlign=center><IMG height=16 src="images/SAVEAS.gif" 
                  width=16>&nbsp;<A 
                  href="changepassword.asp">&nbsp;<SPAN 
                  class=style4>修改密码</SPAN></A></TD>
                          </TR>
                          <TR align=middle> 
                            <TD vAlign=center align=middle background=images/hui.gif 
                height=15></TD>
                          </TR>
                          <TR align=middle> 
                            <TD><IMG height=16 src="images/VOTE.gif" 
                  width=16>&nbsp;<A 
                  href="getpassword.asp">&nbsp;<SPAN 
                  class=style4>取回密码</SPAN></A></TD>
                          </TR>
                          <TR align=middle> 
                            <TD vAlign=center align=middle background=images/hui.gif 
                height=15></TD>
                          </TR>
                          <TR align=middle> 
                            <TD><A href="quit.asp"><FONT color=#ff0000>退出登录</FONT></A></TD>
                          </TR>
                          <TR align=middle> 
                            <TD vAlign=center align=middle background=images/hui.gif 
                height=15></TD>
                          </TR>
                        </TBODY>
                      </TABLE> </TD>
                    <TD width=11 background=images/zuo_bg_2.gif></TD>
                  </TR>
                  <TR vAlign=top align=middle> 
                    <TD colSpan=3><IMG height=11 src="images/zuo_bg_4.gif" 
            width=172></TD>
                  </TR>
                </TBODY>
              </TABLE></TD>
            <TD width=591 valign="top" bgColor=#ffffff>
			
<%
if session("username")="" then
	response.Redirect("QuicklyCheck.asp")
%>	
<%
else
%>
<br>
<form name="theForm" method="post" action="changepasswordadmin.asp" onSubmit="return Juge(this)">  
  <table width="80%" align="center" border=1 cellspacing=0 cellpadding=1 bordercolorlight=#5BD77D bordercolordark=#ffffff>
    <tr>
                    <td class=titlehead height=24>修改密码</td>
    </tr>
  </table>
                <table width="80%" height="108" border="0" align="center" cellpadding="0" cellspacing="0" class="tableborder">
                  <tr>
                    <td height="30"> 
<div align="right">用户名：&nbsp;</div></td>
      <td>&nbsp;<b><%=session("username")%></b>
        <input name="username" type="hidden" id="username" value="<%=session("username")%>"></td>
    </tr>
    <tr> 
                    <td width="32%" height="30"> 
<div align="right">旧密码：&nbsp;</div></td>
                    <td width="68%"> 
                      <input name="oldpassword" type="password" id="oldpassword" size="18">
      </td>
    </tr>
    <tr> 
                    <td height="30"> 
<div align="right">新密码：&nbsp;</div></td>
      <td> <input name="newpassword" type="password" id="newpassword" size="18">
                      4-20个字符（0-9，a-z，A-Z）</td>
    </tr>
    <tr> 
                    <td height="30"> 
<div align="right">确认密码：&nbsp;</div></td>
      <td> <input name="newpassword2" type="password" id="newpassword2" size="18">
                      4-20个字符（0-9，a-z，A-Z）</td>
    </tr>
    <tr> 
      <td height="12">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td height="34"> 
<div align="center"> 
<input type="submit" name="Submit" value=" 修 改 ">
          &nbsp;&nbsp; 
          <input type="reset" name="Submit2" value=" 清 除 ">
        </div></td>
    </tr>
  </table>
</form>
<%end if%>
			  </TD>
          </TR>
        </TBODY>
      </TABLE>

    </td>
  </tr>
</table>
<!-- #include file="foot.asp" -->
<%
conn.close
set conn=nothing
%>
</body>
</html>

