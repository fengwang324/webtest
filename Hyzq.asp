<!--#include file="conn.asp"-->
<%
if session("username")="" or session("s_stat")="" then
		conn.close
		set conn=nothing
		response.Redirect("QuicklyCheck.asp")
		response.end
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
-->
</style>
<SCRIPT language=javascript>
function Juge()
{
	if (document.theForm.password.value=="")
	{
		alert ("请输入用户登陆密码。")
		document.theForm.password.focus();
		return false;
	}
	if (document.theForm.password.value.length<4||document.theForm.password.value.length>20)
	{
		alert ("用户登陆密码字符数太短或太长")
		document.theForm.password.focus();
		return false;
	}
	if (document.theForm.password.value!=document.theForm.password2.value)
	{
		alert ("登陆密码确认错误")
		document.theForm.password2.focus();
		return false;
	}
	if (document.theForm.password_prompt.value=="")
	{
		alert ("请输入密码提示问题")
		document.theForm.password_prompt.focus();
		return false;
	}
	if (document.theForm.password_Answer.value=="")
	{
		alert ("请输入密码提示问题答案")
		document.theForm.password_Answer.focus();
		return false;
	}
	
	if (document.theForm.truename.value=="")
	{
		alert ("请输入用户真实姓名")
		document.theForm.truename.focus();
		return false;
	}
	
	if (document.theForm.telphone1.value=="")
	{
		alert ("请输入联系电话1")
		document.theForm.telphone1.focus();
		return false;
	}
	
	if (document.theForm.address.value=="")
	{
		alert ("请输入联系地址")
		document.theForm.address.focus();
		return false;
	}


   if(confirm("现在真的要修改您的个人资料吗?"))
      return true
   else
      return false;
}
</SCRIPT>
</head>

<body>
<!-- #include file="head.asp" -->
<TABLE width="1020" border=0 align="center" cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
  <tr>
    <td> 

<TABLE height=278 cellSpacing=0 cellPadding=0 width=1020 align=center bgColor=#ffffff border=0>
        <TBODY>
          <TR> 
            <TD vAlign=top width="180"><!-- #include file="zq_left.asp" --></TD>
            <TD width=780 valign="top" bgColor=#ffffff>
			
<%
if session("username")="" or session("s_stat")="" then
		
else
	sql="select * from x_huiyuan where username='"&session("username")&"' "
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.bof or rs.eof then
	else
		'username=rs("username")
		password=rs("password") 
		password_prompt=rs("password_prompt") 
		password_Answer=rs("password_Answer") 
		vipno=rs("vipno") 
		truename=rs("truename") 
		province=rs("province") 
		city=rs("city")
		area=rs("self1")
		telphone1=rs("telphone1") 
		telphone2=rs("telphone2") 
		fax=rs("fax") 
		mobile=rs("mobile") 
		address=rs("address") 
		postno=rs("postno") 
		email=rs("email") 
		babysex=rs("babysex") 
		babybirthday=rs("babybirthday")
		customkind=rs("customkind")
		lasttime=rs("lasttime")
		lastip=rs("lastip")
		
		
	end if
	rs.close
	set rs=nothing	
%>
              <TABLE width="99%" border=0 align="center" cellPadding=3 cellSpacing=0 bgColor=#cccccc class=black-pix12 id=Table3>
                <FORM id=theForm name="theForm" action="hyzqadmin.asp" method="post" onSubmit="return Juge()">
                  <TBODY>
                    <TR bgColor=#ffffff> 
                      <TD colSpan=2 height=28><strong><b><font color=#ff0000><%=session("username")%></font></b> 个人资料：</strong></TD>
                    </TR>
                    <TR bgColor=#ffffff> 
                      <TD width=21% height="28" align=right> <DIV class=black align=right>用户名：</DIV></TD><TD width=79%> <b><font color=#ff0000><%=session("username")%></font></b></TD>
                    </TR>
                    <TR bgColor=#ffffff> 
                      <TD width=21% height="28" align=right> <P class=black align=right>登录密码：</P></TD><TD width=79%> <INPUT name=password value="<%=password%>" type=password id=Password1 
                  style="FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" size=20 maxlength="20"> 
                        <FONT 
                  color=#ff0000>*</FONT> </TD>
                    </TR>
                    <TR bgColor=#ffffff> 
                      <TD align=right width=21%></TD>
                      <TD width=79%><FONT color=#666666>4-20个字符（0-9，a-z，A-Z，下划线） 
                        密码区分大小写；不能与用户名相同。建议您设置一个便于记忆且不易被别人猜到的密码。请牢记密码。</FONT></TD></TR>
                    <TR bgColor=#ffffff> 
                      <TD width=21% height="28" align=right class=black> <DIV align=right>密码确认：</DIV></TD><TD width=79%> <INPUT id=Password2 
                  style="FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" 
                  type=password size=20 name=password2  value="<%=password%>"> 
                        <FONT 
                  color=#ff0000>*</FONT> </TD>
                    </TR>
                    <TR bgColor=#ffffff> 
                      <TD height="28" align=right class=black>电子邮件：</TD>
                      <TD><INPUT name="email" value="<%=email%>" id=email 
                  style="FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" 
                  size=30 maxlength="50"></TD>
                    </TR>
                    <TR bgColor=#ffffff> 
                      <TD width=21% height="28" align=right class=black> <DIV align=right>密码提示问题：</DIV></TD><TD width=79%> <SELECT style="width:160px;FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; HEIGHT: 21px; FONT-VARIANT: normal" size=1 name="password_prompt">
                          <OPTION value="<%=password_prompt%>" selected><%=password_prompt%></OPTION>
                          <OPTION value="" >请选择问题</OPTION>
                          
                          <OPTION value="您的初中班主任是？">您的初中班主任是？</OPTION>
                          <OPTION value="您的高中班主任是？">您的高中班主任是？</OPTION>
                          <OPTION value="您的配偶的名字是？">您的配偶的名字是？</OPTION>
                          <OPTION value="您的配偶的职业是？">您的配偶的职业是？</OPTION>
                          <OPTION value="您的配偶的生日是？">您的配偶的生日是？</OPTION>
                          <OPTION value="您的父亲的名字是？">您的父亲的名字是？</OPTION>
                          <OPTION value="您的父亲的职业是？">您的父亲的职业是？</OPTION>
                          <OPTION value="您的母亲的名字是？">您的母亲的名字是？</OPTION>
                          <OPTION value="您的母亲的生日是？">您的母亲的生日是？</OPTION>
                          <OPTION value="您最喜欢的歌曲名是？">您最喜欢的歌曲名是？</OPTION>
                          <OPTION value="您最喜欢的电影名是？">您最喜欢的电影名是？</OPTION>
                        </SELECT> <FONT 
                  color=#ff0000>*</FONT> <font color="#666666">用于取回密码（如：我的初中班主任是？）</font></TD>
                    </TR>
                    <TR bgColor=#ffffff> 
                      <TD width=21% height="28" align=right class=black> <DIV align=right>密码提示答案：</DIV></TD><TD width=79%> <INPUT id=Text4 
                  style="FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" 
                  size=20 name=password_Answer  value="<%=password_Answer%>"> 
                        <FONT 
                  color=#ff0000>*</FONT> <font color="#666666">当忘记密码，点击“忘记密码”正确回答此答案即可取回密码。</font></TD>
                    </TR>
                    <TR bgColor=#ffffff> 
                      <TD width=21% height="28" align=right>真实姓名：</TD>
                      <TD width=79%> <INPUT id=Text2 
                  style="FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" 
                  size=20 name=truename  value="<%=truename%>"> <FONT color=#ff0000>*</FONT> 
                        <font 
                  color=#666666>请留下您的真实姓名，便于我们与您的联系以及给您发货。</font></TD>
                    </TR>
                    <TR bgColor=#ffffff> 
                      <TD width=21% height="28" align=right class=black> <DIV align=right>省份：</DIV></TD><TD width=79%>
					  <INPUT name=province value="<%=province%>" id=Text8 size="20" style="FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal"></TD>
                    </TR>
                    <TR bgColor=#ffffff> 
                      <TD width=21% height="28" align=right class=black> <DIV align=right>城市：</DIV></TD><TD width=79%>
					  <INPUT name=city value="<%=city%>" id=Text8 size="20" style="FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal"></TD>
                    </TR>
                    <TR bgColor=#ffffff> 
                      <TD width=21% height="28" align=right class=black> <DIV align=right>地区：</DIV></TD><TD width=79%>
					  <INPUT name=area value="<%=area%>" id=Text8 size="20" style="FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal"></TD>
                    </TR>
                    <TR bgColor=#ffffff> 
                      <TD width=21% height="28" align=right class=black> <DIV align=right>联系地址：</DIV></TD><TD class=black width=79%> <INPUT id=Text11 
                  style="FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" 
                  size=50 name=address value="<%=address%>"> <FONT color=#ff0000>*</FONT><font color="#0066CC">本地址用于商品送货上门或邮寄地址</font></TD>
                    </TR>
                    <TR bgColor=#ffffff> 
                      <TD width=21% height="28" align=right class=black> <DIV align=right>邮政编码：</DIV></TD><TD class=black width=79%> <INPUT name=postno value="<%=postno%>" id=Text13 
                  style="FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" 
                  size=20 maxlength="20"> </TD>
                    </TR>
                    <TR bgColor=#ffffff> 
                      <TD width=21% height="28" align=right class=black> <DIV align=right>联系电话：</DIV></TD>
                      <TD width=79%> <INPUT id=Text9 
                  style="FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" 
                  size=20 name=telphone1  value="<%=telphone1%>"> <FONT color=#ff0000>*</FONT> 
                        <FONT 
                  color=#666666>请填写您的固定电话，含区号和电话号码</FONT></TD>
                    </TR>

                    <TR bgColor=#ffffff> 
                      <TD width=21% height="28" align=right class=black> <DIV align=right>移动电话：</DIV></TD><TD class=black width=79%> <INPUT id='Text14' 
                  style="FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" 
                  size=20 name='mobile'  value="<%=mobile%>"> </TD>
                    </TR>



                    <TR bgColor=#ffffff> 
                      <TD height="28" align=right class=black>会员等级：</TD>
                      <TD><b>
					      <%
						  if customkind="2" then 
						 	response.write "普通会员"
                          elseif customkind="3" then 
						 	response.write "铜级会员"
                          elseif customkind="4" then 
						    response.write "银级会员"
                          elseif customkind="5" then 
						    response.write "金级会员"
                          elseif customkind="6" then 
						    response.write "钻级会员"
						  else
						  end if
						  %>
					  </b></TD>
                    </TR>
                    <TR bgColor=#ffffff> 
                      <TD height="28" align=right class=black>账户余额：</TD>
                      <TD>
					  <%
						sql6="SELECT id, username,(select sum(czmoney) from sh_chongzhi where sh_chongzhi.customid=a.id and confirmflag='是') AS sumcz, "&_
						"(select sum(total) from x_bill where x_bill.customid=a.id and paytype='账户余额') AS sumkj,  "&_
						"(IIf(IsNull(sumcz),0,sumcz)-IIf(IsNull(sumkj),0,sumkj)) AS leftye  "&_
						"FROM x_huiyuan AS a where username='"&session("username")&"' "
						
						set rs6=server.createobject("adodb.recordset")
						rs6.open sql6,conn,1,1
						if rs6.bof or rs6.eof then
							response.write "＄<input type='text' name='leftye' value='0' readonly style='width:85;border:0px;BACKGROUND-COLOR: #ffffff;FONT-WEIGHT: bold; FONT-SIZE: 17px; COLOR:#f00;FONT-FAMILY: Century Gothic'>"
						else
							response.write "＄<input type='text' name='leftye' value='"&rs6("leftye")&"' readonly style='width:85;border:0px;BACKGROUND-COLOR: #ffffff;FONT-WEIGHT: bold; FONT-SIZE: 17px; COLOR:#f00;FONT-FAMILY: Century Gothic'>"
						end if
						rs6.close
						set rs6=nothing
					  %>					  </TD>
                    </TR>
                    <TR bgColor=#ffffff> 
                      <TD height="28" align=right class=black>剩余积分：</TD>
                      <TD>
					  <%
						sql6="SELECT id, username,(select sum(billzhf) from x_view_bill where x_view_bill.username=a.username and sendornot='已发货') AS sumzhf, "&_
						"(select sum(kjzhf) from x_huanlipin where x_huanlipin.username=a.username) AS sumkjzhf,  "&_
						"(IIf(IsNull(sumzhf),0,sumzhf)-IIf(IsNull(sumkjzhf),0,sumkjzhf)) AS leftzhf  "&_
						"FROM x_huiyuan AS a where username='"&session("username")&"' "
						
						set rs6=server.createobject("adodb.recordset")
						rs6.open sql6,conn,1,1
						if rs6.bof or rs6.eof then
							response.write "<input type='text' name='leftzhf' value='0' readonly style='width:85;border:0px;BACKGROUND-COLOR: #ffffff;FONT-WEIGHT: bold; FONT-SIZE: 17px; COLOR:#f00;FONT-FAMILY: Century Gothic'>"
						else
							response.write "<input type='text' name='leftzhf' value='"&rs6("leftzhf")&"' readonly style='width:85;border:0px;BACKGROUND-COLOR: #ffffff;FONT-WEIGHT: bold; FONT-SIZE: 17px; COLOR:#f00;FONT-FAMILY: Century Gothic'>"
						end if
						rs6.close
						set rs6=nothing
					  %>					  </TD>
                    </TR>
					
                    <TR bgColor=#ffffff> 
                      <TD height="28" align=right class=black>上次陆登时间：</TD>
                      <TD>
					  <%=lasttime%>					  </TD>
                    </TR>
                    <TR bgColor=#ffffff> 
                      <TD height="28" align=right class=black>上次陆登的IP：</TD>
                      <TD>
					  <%=lastip%>					  </TD>
                    </TR>
                    <TR bgColor=#ffffff>
                      <TD height="28" align=right class=black>&nbsp;</TD>
                      <TD><input type="submit" name="Submit" value="修改注册信息" style="width:100px;color:#ff0000;FONT-SIZE: 10pt; FONT-WEIGHT: bold; background-color:#dddddd">&nbsp; 
					  <input id=Hidden1 type=hidden value=add name=eaction> 
                      <input type="reset" name="Submit2" value="重新填写" style="width:80px;FONT-SIZE: 10pt; FONT-WEIGHT: bold"> </TD>
                    </TR>
					
					
                    <TR vAlign=top bgColor=#ffffff> 
                      <TD colSpan=2>&nbsp;</TD>
                    </TR>
                  </TBODY>
                </form>
              </TABLE>
<%end if%>
		    </TD>
          </TR>
        </TBODY>
      </TABLE>

    </td>
  </tr>

<TR><TD>

</TD></TR>
</table>

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
<!-- #include file="foot.asp" -->
<%
conn.close
set conn=nothing
%>
</body>
</html>

