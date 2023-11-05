<!--#include file="conn.asp"-->
<%
if session("verifycode")<>request.Form("code") then
response.write "<br><br><div align=center><a href=javascript:window.history.back()>请输入正确的验证码，点击这里返回重试</a></div>"
response.end
end if
username=trim(request.Form("username"))
password_Answer=trim(request.Form("password_Answer"))
mail=trim(request.Form("mail"))
if InStr(password_Answer,"'")>0 or InStr(password_Answer,"--")>0 or InStr(password_Answer,"(")>0 or InStr(password_Answer,";")>0 then
	response.write "密码提示答案不合法。"
	response.end
END IF
if InStr(mail,"'")>0 or InStr(mail,"(")>0 or InStr(mail,";")>0 then
	response.write "mail不合法。"
	response.end
END IF

sql="select password from x_huiyuan where username='"&username&"' and password_Answer='"&password_Answer&"' and email='"&mail&"' "
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof or rs.eof then
	response.write "<br><br><div align=center style='BACKGROUND-COLOR: #fafafa;height:90px;width:500px;BORDER-RIGHT: #666666 1px solid;BORDER-TOP: #666666 1px solid; BORDER-LEFT: #666666 1px solid; BORDER-BOTTOM: #666666 1px solid;text-valign:center;'><br><br><a href='getpassword.asp'>对不起，您所输入的答案不对，请点击这里返回重试</a><br>如果尝试多次未能取回密码，请与本站客服人员联系。</div>"
	response.end
else
	password=rs("password")
end if

%>
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
  if (theForm.password_Answer.value == "")
  {
    alert("请输入密码提示答案!");
    theForm.password_Answer.focus();
    return (false);
  }
  if (theForm.code.value == "")
  {
    alert("请输入验证码!");
    theForm.code.focus();
    return (false);
  } 
}
-->
</script>
</head>

<body>
<!-- #include file="head.asp" -->
<TABLE width="960" border=0 align="center" cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
  <tr>
    <td> 

<TABLE height=278 cellSpacing=0 cellPadding=0 width=960 align=center bgColor=#ffffff border=0>
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
            <TD width=591 valign="top" bgColor=#ffffff> <br>
<form name="theForm" method="post" action="getpassword3.asp" onSubmit="return Juge(this)">
                
                <table width="85%" align="center" border=1 cellspacing=0 cellpadding=1 bordercolorlight=#5BD77D bordercolordark=#ffffff>
                  <tr>
                    <td class=titlehead height=24>取回密码：</td>
    </tr>
  </table>
                <table width="85%" height="86" border="0" align="center" cellpadding="0" cellspacing="0" class="tableborder">
                  <tr> 
                    <td height="20"> <div align="right">&nbsp;</div></td>
                    <td>&nbsp;</td>
                  </tr>
                  <TR bgColor=#ffffff> 
                    <TD width=21% height="25" align=right class=black> <DIV align=right>您的密码是：</DIV></TD>
                    <TD width=79%><b><%=password%></b> 。请牢记您的密码。</TD>
                  </TR>
                  <tr> 
                    <td height="41" colspan="2">&nbsp; </td>
                  </tr>
                </table>
  </form>
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

