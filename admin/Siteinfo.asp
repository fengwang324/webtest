<!--#include file="conn.asp"-->
<!--#include file="check2.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>网站后台,网站基本信息</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<style>
<!--
td{font-size:9pt}
body {
	margin-top: 5px;margin-left: 2px;
}
.style3 {color: #A84E22; font-weight: bold; }
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
.STYLE4 {	color: #FF0000;
	font-weight: bold;
}
-->
</style>
<%
dim sql4,rs4,id,num,pkid

siteurl=trim(request.form("siteurl"))
sitename=trim(request.form("sitename"))
sitelogo=trim(request.form("sitelogo"))
siteright=trim(request.form("siteright"))
sitetel=trim(request.form("sitetel"))
siteren=trim(request.form("siteren"))
siteemail=trim(request.form("siteemail"))
sitedisc=trim(request.form("sitedisc"))
sitekeyword=trim(request.form("sitekeyword"))
if sitekeyword<>"" then sitekeyword = replace(sitekeyword,vbcrlf,"&nbsp;&nbsp;")
Submit2=trim(request.form("Submit2"))

s_qq=trim(request.form("s_qq"))
if s_qq<>"" then
	s_qq=s_qq&":"
	s_qq=replace(s_qq,"：",":")
	s_qq=replace(s_qq,"::",":")
	s_qq=replace(s_qq,"::",":")
	s_qq=replace(s_qq,"::",":")
end if

if Submit2="保存" then

		sql4="select * from siteinfo"
		set rs4=server.createobject("adodb.recordset")
		rs4.open sql4,conn,1,3
		if rs4.bof or rs4.eof then
				rs4.addnew
				rs4("siteurl")=siteurl
				rs4("sitename")=sitename
				rs4("sitelogo")=sitelogo
				rs4("siteright")=siteright
				rs4("sitetel")=sitetel
				rs4("siteren")=siteren
				rs4("siteemail")=siteemail
				rs4("sitedisc")=sitedisc
				rs4("sitekeyword")=sitekeyword
				rs4.update
				conn.execute("update e_board set title='"&sitename&"'")
		else
				rs4("siteurl")=siteurl
				rs4("sitename")=sitename
				rs4("sitelogo")=sitelogo
				rs4("siteright")=siteright
				rs4("sitetel")=sitetel
				rs4("siteren")=siteren
				rs4("siteemail")=siteemail
				rs4("sitedisc")=sitedisc
				rs4("sitekeyword")=sitekeyword
				rs4.update
				conn.execute("update e_board set title='"&sitename&"'")
		end if
		rs4.close
		set rs4=nothing
		
	 
		
		response.write "<script language=JavaScript>" & "alert('保存成功!');"& "window.location.href='siteinfo.asp';" & "</script>"
end if

%>
</head>

<body bgcolor="#fcfcfc">          
<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0" id="tabcss">
  <tr bgcolor="#E1EFFF"> 
    <td height=30 colspan=2>&nbsp;<b>网站基本信息：</b></td>
  </tr>
<form name="form1" method="post" action="siteinfo.asp">
  <%
		sql4="select * from siteshow"
		set rs4=server.createobject("adodb.recordset")
		rs4.open sql4,conn,1,1
		if rs4.bof or rs4.eof then
			response.write "没有符合条件订单记录！"
		else
			s_zishu=rs4("s_zishu")
			s_qq=rs4("s_qq")
			s_msn=rs4("s_msn")
			s_wangwang=rs4("s_wangwang")
		end if
		rs4.close
		set rs4=nothing


		sql4="select * from siteinfo"
		set rs4=server.createobject("adodb.recordset")
		rs4.open sql4,conn,1,1
		if rs4.bof or rs4.eof then
			response.write "没有符合条件订单记录！"
		else
				id=rs4("id")
				siteurl=rs4("siteurl")
				sitename=rs4("sitename")
				sitelogo=rs4("sitelogo")
				siteright=rs4("siteright")
				sitetel=rs4("sitetel")
				siteren=rs4("siteren")
				siteemail=rs4("siteemail")
				sitedisc=rs4("sitedisc")
				sitekeyword=rs4("sitekeyword")
				if sitekeyword<>"" then sitekeyword = replace(sitekeyword,"&nbsp;&nbsp;",vbcrlf)
%>
  <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
    <td height='23' width="13%" align='right'>网站网址</td>
    <td width="87%">
      <input type=text name='siteurl' value="<%=siteurl%>" size="30"></td>
  </tr>
  <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
    <td height='23' align='right'>网站名称</td>
    <td><input type=text name='sitename' value="<%=sitename%>" size="30"></td>
  </tr>
  <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
    <td height='23' align='right'>网站LOGO图片</td>
    <td><input type=text name='sitelogo' value="<%=sitelogo%>" size="40" readonly style="background:'#FFf4CA';"> 
	<input type="button" name="Button" value="选图片" onClick="window.open('../images/uploadpic.asp?size=sitelogo&mylogo=1','','width=310,height=250,top=100,left=220')">
	图片宽度为<span class="STYLE4">300px以内</span>,高度为<span class="STYLE4">75px以内</span></td>
  </tr>
  <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
    <td height='23' align='right'>网站版权</td>
    <td><input type=text name='siteright' value="<%=siteright%>" size="30"></td>
  </tr>
  <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'>订购热线电话</td>
    <td><input type=text name='sitetel' value="<%=sitetel%>" size="30"></td>
  </tr>
  <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
    <td height='23' align='right'>联系人</td>
    <td><input type=text name='siteren' value="<%=siteren%>" size="30"></td>
  </tr>
  <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
    <td height='23' align='right'>E-mail</td>
    <td><input type=text name='siteemail' value="<%=siteemail%>" size="30"></td>
  </tr>
  <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
    <td height='23' align='right'>网店广告语</td>
    <td><input type=text name='sitedisc' value="<%=sitedisc%>" size="30"></td>
  </tr>
  
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';">
      <td height='35' align='right'>头部搜索的关键字</td>
      <td style="padding:2px;"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="42%"><textarea name="sitekeyword" cols="58" rows="7"><%=sitekeyword%></textarea></td>
            <td width="58%" style="line-height:130%">&nbsp;</td>
          </tr>
        </table></td>
    </tr>
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='35' align='right'>&nbsp;</td>  
<td><input type="submit" name="Submit2" value=" 保存 ">&nbsp;&nbsp;&nbsp;&nbsp;
        <input type="reset" name="Submit3" value=" 重置 "> </td>
  </tr>
  <%				
		end if
rs4.close
set rs4=nothing		
%>
</form>
</table>
<br><br><br>

</body>
</html>
<%
call connclose
%>

