<!--#include file="conn.asp"-->
<%
on error resume next
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
	color: #990000;
	font-weight: bold;
}
.style12 {color: #990000}
.style13 {color: #333333}
.style15 {font-size: 14px; color: #669933; font-weight: bold; }
-->
</style></head>
<link href="style.css" rel="stylesheet" type="text/css" />
<body>
<!-- #include file="head.asp" -->
<TABLE width="1020" border=0 align="center" cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
  <tr>
    <td> 

    <table width="1020" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr>
          <td width="210" valign="top">
		  <!-- #include file="inc_left_all.asp" -->
		  </td>
          <td width="750" valign="top">
	
            <TABLE class=text width="98%" align=center cellSpacing=0 cellPadding=0 border=0>
              <TBODY>
                <TR> 
                  <TD vAlign=bottom height=32><IMG height=19 src="images/dian04.gif" width=10 align=absMiddle> <STRONG>&nbsp;访客留言

			      </STRONG></TD>
                </TR>
                <TR> 
                  <TD vAlign=top height=25> <DIV align=left><IMG height=25 src="images/line.gif" width=457></DIV></TD>
                </TR>
                <TR> 
                  <TD height="320" vAlign=top>
<%if request("xingming")="" or request("dianhua")="" or request("dizhi")="" or request("content")="" then%>
<script language="javascript">alert("请将各项填写完整！");history.go(-1);</script>
<%else%>
<!--#include file="char.inc"-->
<%set rs10=server.createobject("adodb.recordset")
sqln10="select * from khly"
rs10.open sqln10,conn,3,3
rs10.addnew
rs10("xingming")=request("xingming")
rs10("dianhua")=request("dianhua")
rs10("email")=request("email")
rs10("dizhi")=request("dizhi")
rs10("title")=request("title")
rs10("content")=htmlencode1(request("content"))
rs10("ntime")=now()
rs10.update
rs10.close
set rs10=nothing

Set rsemsz=Server.CreateObject("ADODB.Recordset")
Sqlnemsz="select * from emailsetting where id=1"
rsemsz.open Sqlnemsz,conn,1,1
if rsemsz("zhuangtai")=1 then

		set rs1=server.createobject("adodb.recordset")
		sqln1="select top 1 * from khly where xingming='"& request("xingming") &"' order by id desc"
		rs1.open sqln1,conn,1,1
		if not rs1.eof then
		stranr=""
		stranr=stranr&"<table style=""font-size:14px;line-height:28px;"" align=""left"" width=""760"" border=""0"" cellspacing=""0"" cellpadding=""0"">"&vbCrLf
		stranr=stranr&"  <tr>"&vbCrLf
		stranr=stranr&"    <td width=""160"" height=""30"" align=""right""><b>姓名：</b></td>"&vbCrLf
		stranr=stranr&"    <td width=""600"" align=""left"">"& rs1("xingming") &"</td>"&vbCrLf
		stranr=stranr&"  </tr>"&vbCrLf
		stranr=stranr&"  <tr>"&vbCrLf
		stranr=stranr&"    <td height=""30"" align=""right""><b>电话：</b></td>"&vbCrLf
		stranr=stranr&"    <td align=""left"">"& rs1("dianhua") &"</td>"&vbCrLf
		stranr=stranr&"  </tr>"&vbCrLf
		stranr=stranr&"  <tr>"&vbCrLf
		stranr=stranr&"    <td height=""30"" align=""right""><b>邮箱：</b></td>"&vbCrLf
		stranr=stranr&"    <td align=""left"">"& rs1("email") &"</td>"&vbCrLf
		stranr=stranr&"  </tr>"&vbCrLf
		stranr=stranr&"  <tr>"&vbCrLf
		stranr=stranr&"    <td height=""30"" align=""right""><b>地址：</b></td>"&vbCrLf
		stranr=stranr&"    <td align=""left"">"& rs1("dizhi") &"</td>"&vbCrLf
		stranr=stranr&"  </tr>"&vbCrLf
		stranr=stranr&"  <tr>"&vbCrLf
		stranr=stranr&"    <td height=""30"" align=""right""><b>留言主题：</b></td>"&vbCrLf
		stranr=stranr&"    <td align=""left"">"& rs1("title") &"</td>"&vbCrLf
		stranr=stranr&"  </tr>"&vbCrLf
		stranr=stranr&"  <tr valign=""top"">"&vbCrLf
		stranr=stranr&"    <td height=""30"" align=""right""><b>留言内容：</b></td>"&vbCrLf
		stranr=stranr&"    <td align=""left"">"& htmlencode2(rs1("content")) &"</td>"&vbCrLf
		stranr=stranr&"  </tr>"&vbCrLf
		stranr=stranr&"  <tr>"&vbCrLf
		stranr=stranr&"    <td height=""30"" align=""right""><b>留言时间：</b></td>"&vbCrLf
		stranr=stranr&"    <td align=""left"">"& rs1("ntime") &"</td>"&vbCrLf
		stranr=stranr&"  </tr>"&vbCrLf
		stranr=stranr&"</table>"&vbCrLf

sendUrl="http://schemas.microsoft.com/cdo/configuration/sendusing"
smtpUrl="http://schemas.microsoft.com/cdo/configuration/smtpserver"
Set objConfig=CreateObject("CDO.Configuration")
objConfig.Fields.Item(sendUrl)=2
objConfig.Fields.Item(smtpUrl)="relay-hosting.secureserver.net"
objConfig.Fields.Update
Set objMail=CreateObject("CDO.Message")
Set objMail.Configuration=objConfig
objMail.From = ""& rsemsz("fajianxiang") &"" '这里一定要是GoDaddy提供的邮局,否则发送不了
objMail.To = ""& rsemsz("shoujianxiang") &""
objMail.BodyPart.Charset = "gb2312"
objMail.Subject="网站留言  "& rs1("xingming") &"  "& rs1("ntime") &""
'objMail.HtmlBody = "dddd"
objMail.HtmlBody = "<html><head><META content=zh-cn http-equiv=Content-Language><meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312""></head><body style=""font-size:14px;line-height:28px; "">"& stranr &"</body></html>"
objMail.Send
Set objMail = Nothing
		
		end if
		rs1.close
		set rs1=nothing


end if
rsemsz.close
set rsemsz=nothing
conn.close
set conn=nothing%>
<script language="javascript">alert("留言提交成功！我们会尽快与您联系。");window.location = "index.asp";</script>
<%end if%>
                  
                  
                  
                  
                  <table align="center" style="font-size:14px; color:#000000; " width="700" border="0" cellspacing="0" cellpadding="8">
<SCRIPT language=JavaScript> 
<!--
function CheckFormmessage(theForm){
if (theForm.xingming.value.length<1)
	{alert("请填写姓名！");
	 theForm.xingming.focus();
	 return false;
	}
if (theForm.dianhua.value.length<1)
	{alert("请填写电话！");
	 theForm.dianhua.focus();
	 return false;
	}
//if (theForm.email.value.length<1)
//	{alert("请输入您的电子邮箱！");
//	 theForm.email.focus();
//	 return false;
//	}
//if (theForm.email.value.indexOf("@") == -1)	
//	{alert("请输入正确的电子邮箱！");
//	 theForm.email.focus();
//	 return false;
//	}
//if (theForm.email.value.indexOf(".") == -1)	
//	{alert("请输入正确的电子邮箱！");
//	 theForm.email.focus();
//	 return false;
//	}
if (theForm.dizhi.value.length<1)
	{alert("请填写地址！");
	 theForm.dizhi.focus();
	 return false;
	}
if (theForm.content.value.length<1)
	{alert("请填写留言内容！");
	 theForm.content.focus();
	 return false;
	}
return true;
 }
//-->
</SCRIPT>
 <form method="post" action="booksubmit.asp" name="form1" onSubmit="return CheckFormmessage(this)"><tr>
    <td align="right" width="120"><font color="#FF0000">*&nbsp;</font>姓名：</td>
    <td align="left" width="480"><input type="text" name="xingming" size="50" maxlength="30"/></td>
  </tr>
  <tr>
    <td align="right"><font color="#FF0000">*&nbsp;</font>电话：</td>
    <td align="left"><input type="text" name="dianhua" size="50" maxlength="30"/></td>
  </tr>
  <tr>
    <td align="right">邮箱：</td>
    <td align="left"><input type="text" name="email" size="60" maxlength="50"/></td>
  </tr>
  <tr>
    <td align="right"><font color="#FF0000">*&nbsp;</font>地址：</td>
    <td align="left"><input type="text" name="dizhi" size="60" maxlength="200"/></td>
  </tr>
  <tr>
    <td align="right">留言主题：</td>
    <td align="left"><input type="text" name="title" size="60" maxlength="200"/></td>
  </tr>
  <tr>
    <td align="right" valign="top"><font color="#FF0000">*&nbsp;</font>留言内容：</td>
    <td align="left"><textarea name="content" cols="69" rows="8"></textarea></td>
  </tr>
  <tr>
    <td align="right">&nbsp;</td>
    <td align="left" valign="bottom"><INPUT type="image" id="Image6" src="images/tj1.gif" name="Submit" style="width:128px; height:27px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="images/cz.gif" style="CURSOR: hand" onClick="return reset();"></td>
  </tr></form>
</table></TD>
                </TR>
                <TR> 
                  <TD height="25">&nbsp;</TD>
                </TR>
              </TBODY>
            </TABLE>
            
          </td>
  </tr>
</table>


    </td>
  </tr>

<TR><TD>

</TD></TR>
</table>
<!-- #include file="foot.asp" -->
</body>
</html>



