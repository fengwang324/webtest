<!--#include file="conn.asp"-->
<%session("reg")="1"%>
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

<SCRIPT language=javascript>
function Juge()
{
	if (document.theForm.username.value=="")
	{
		alert ("请输入用户名。")
		document.theForm.username.focus();
		return false;
	}

	if (document.theForm.username.value.length<4||document.theForm.username.value.length>20)
	{
		alert ("“用户名”字符数太短或太长")
		document.theForm.username.focus();
		return false;
	}

	if (document.theForm.password.value=="")
	{
		alert ("请输入用户登陆密码。")
		document.theForm.password.focus();
		return false;
	}
	if (document.theForm.password.value.length<6||document.theForm.password.value.length>20)
	{
		alert ("用户登陆密码字符数太短或太长。请填写6-20位字符的密码。")
		document.theForm.password.focus();
		return false;
	}
	if (document.theForm.password.value!=document.theForm.password2.value)
	{
		alert ("登陆密码确认错误")
		document.theForm.password2.focus();
		return false;
	}
	
	if (document.theForm.email.value=="")
	{
		alert ("请输入电子邮件")
		document.theForm.email.focus();
		return false;
	}
	var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_@.";
	var checkStr = theForm.email.value;
	var allValid = true;
	for (i = 0;  i < checkStr.length;  i++)
	{
		ch = checkStr.charAt(i);
		for (j = 0;  j < checkOK.length;  j++)
		if (ch == checkOK.charAt(j))
			break;
		if (j == checkOK.length)
		{
			allValid = false;
			break;
		}
	}

	if (theForm.email.value.length < 6)
	{
			allValid = false;
	}

	if (!allValid)
	{
		alert("您输入的 \"电子邮件地址\" 无效!");
		theForm.email.focus();
		return (false);
	}

	address=theForm.email.value;
   	if(address.length>0)
	 {
          i=address.indexOf("@");
          if(i==-1)
		{
			window.alert("对不起！您输入的电子邮件地址是错误的！")
			theForm.email.focus();
			return false
                }
       ii=address.indexOf(".")
        if(ii==-1)
		{
			window.alert("对不起！您输入的电子邮件地址是错误的！")
			theForm.email.focus();
			return false
                }

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

  if (theForm.province.value == "")
  {
    alert("请选择省份!");
    document.theForm.province.focus();
    return false;
  }

  if (theForm.city.value == "")
  {
    alert("请选择城市!");
    document.theForm.city.focus();
    return false;
  }
  
  if (theForm.area.value == "市辖区")
  {
    alert("请选择地区!");
    document.theForm.area.focus();
    return false;
  }


	if (document.theForm.address.value=="")
	{
		alert ("请输入详细联系地址")
		document.theForm.address.focus();
		return false;
	}

	if (document.theForm.telphone1.value=="")
	{
		alert ("请输入联系电话")
		document.theForm.telphone1.focus();
		return false;
	}
	




	//document.theForm.Naction.value="Change";
	//document.theForm.submit();
}
</SCRIPT>

<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	
	margin-bottom: 0px;
	margin-right: 0px;
}
.style1 {font-size: 14px}
.style11 {
	font-size: 14px;
	color: #990000;
	font-weight: bold;
}
.style12 {color: #990000}
.style13 {color: #333333}
.style15 {font-size: 14px; color: #669933; font-weight: bold; }

.register_submit {
	BACKGROUND: url(images/regist_sub.jpg) no-repeat; WIDTH: 111px; HEIGHT: 25px; CURSOR: hand; BORDER-BOTTOM-WIDTH: 0px; BORDER-TOP-WIDTH: 0px; BORDER-LEFT-WIDTH: 0px; BORDER-RIGHT-WIDTH: 0px; TEXT-DECORATION: underline
}
-->
</style>
</head>

<body>
<!-- #include file="head.asp" -->
<TABLE width="1020" border=0 align="center" cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
  <tr>
    <td> 



<table width="1020" height="352" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
    <td width="210" height="184" valign="top">
		<!-- #include file="inc_left_all.asp" -->
      </td>
    <td width="750" valign="top">
	<table width="99%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height=2></td>
      </tr>
    </table>
            <table width="50%" height="2"  border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td></td>
        </tr>
      </table>

            <TABLE width="99%" border=0 align="center" cellPadding=0 cellSpacing=0 >
              <FORM name="theForm" action="regeditadmin.asp" method="post" onSubmit="return Juge()">
                <TBODY>
				  <tr><TD height="30" style="padding-top:2px;font-size:11pt;font-weight:bold;color:#ff8800">&nbsp;&nbsp;新用户注册：</TD></tr>
                  <TR> 
                    <TD>
                      <TABLE width="99%" border=0 align="center" cellPadding=3 cellSpacing=1 class=black-pix12 >


<TR bgColor=#ffffff> 
                            
                          <TD width="21%" align=right class=black><strong><font color="#FF0000">注册声明：</font></strong></TD>
                            
                          <TD width="79%">
                            <textarea name="textarea" cols="78" rows="8" readonly>请你先阅读我们的注册声明，如果同意，请进行用户注册。 
1．<%=application("sitename")%>网站服务条款的确认和接纳
<%=application("sitename")%>网站各项服务的所有权和运作权归<%=application("sitename")%>拥有。
2．用户必须：
(1)自行配备上网的所需设备， 包括个人电脑、调制解调器或其他必备上网装置。
　　 (2)自行负担个人上网所支付的与此服务有关的电话费用、 网络费用。
3． 有关个人资料
用户同意：
(1) 提供及时、详尽及准确的个人资料。
(2) 不断更新注册资料，符合及时、详尽准确的要求。所有原始键入的资料将引用为注册资料。
　　 <%=application("sitename")%>网站不公开用户的姓名、地址、电子邮箱和笔名，以下情况除外：
　　 (1)用户授权<%=application("sitename")%>网站透露这些信息。
　　 (2)相应的法律及程序要求<%=application("sitename")%>网站提供用户的个人资料。如果用户提供的资料包含有不正确的信息，<%=application("sitename")%>网站保留结束用户使用<%=application("sitename")%>网站信息服务资格的权利。
4． 服务条款的修改
　　 <%=application("sitename")%>网站有权在必要时修改服务条款，<%=application("sitename")%>网站服务条款一旦发生变动，将会在重要页面上提示修改内容。如果不同意所改动的内容，用户可以主动取消获得的本网站信息服务。 如果用户继续享用<%=application("sitename")%>网站信息服务，则视为接受服务条款的变动。<%=application("sitename")%>网站保留随时修改或中断服务而不需通知用户的权利。<%=application("sitename")%>网站行使修改或中断服务的权利，不需对用户或第三方负责。
5． 用户隐私制度
　　 尊重用户个人隐私是<%=application("sitename")%>网站的一项基本政策。所以，作为对以上第二点个人注册资料分析的补充，<%=application("sitename")%>网站一定不会在未经合法用户授权时公开、编辑或透露其注册资料及保存在<%=application("sitename")%>网站中的非公开内容 ，除非有法律许可要求或<%=application("sitename")%>网站在诚信的基础上认为透露这些信息在以下四种情况是必要的：
(1) 遵守有关法律规定，遵从<%=application("sitename")%>网站合法服务程序。
(2) 保持维护<%=application("sitename")%>网站的商标所有权。
(3) 在紧急情况下竭力维护用户个人和社会大众的隐私安全。
(4)符合其他相关的要求。
<%=application("sitename")%>网站保留发布会员人口分析资询的权利。
6．用户的帐号、密码和安全性
　　 你一旦注册成功成为用户，你将得到一个密码和帐号。如果你不保管好自己的帐号和密码安全，将负全部责任。另外，每个用户都要对其帐户中的所有活动和事件负全责。你可随时根据指示改变你的密码， 也可以结束旧的帐户重开一个新帐户。用户同意若发现任何非法使用用户帐号或安全漏洞的情况，请立即通告<%=application("sitename")%>网站。
7． 拒绝提供担保
　　 用户明确同意信息服务的使用由用户个人承担风险。 <%=application("sitename")%>网站不担保服务不会受中断，对服务的及时性，安全性，出错发生都不作担保，但会在能力范围内，避免出错。
8．有限责任
　　 <%=application("sitename")%>网站对任何直接、间接、偶然、特殊及继起的损害不负责任，这些损害来自：不正当使用<%=application("sitename")%>网站服务，或用户传送的信息不符合规定等。这些行为都有可能导致<%=application("sitename")%>网站形象受损， 所以<%=application("sitename")%>网站事先提出这种损害的可能性，同时会尽量避免这种损害的发生。
9．信息的储存及限制
　　 <%=application("sitename")%>网站有判定用户的行为是否符合<%=application("sitename")%>网站服务条款的要求和精神的权利，如果用户违背<%=application("sitename")%>网站服务条款的规定，<%=application("sitename")%>网站有权中断其服务的帐号。如果用户在连续360天时间内没有登录、查看或使用，<%=application("sitename")%>网站将视为用户自行放弃该帐号的使用权，<%=application("sitename")%>网站将保留中断对其提供服务的权利。
10．用户管理
　　 用户必须遵循：
(1) 从中国境内向外传输技术性资料时必须符合中国有关法规。
(2) 使用信息服务不作非法用途。
(3) 不干扰或混乱网络服务。
(4) 遵守所有使用服务的网络协议、规定、程序和惯例。用户的行为准则是以因特网法规，政策、程序和惯例为根据的。
11．保障
　　 用户同意保障和维护<%=application("sitename")%>网站全体成员的利益，负责支付由用户使用超出服务范围引起的律师费用，违反服务条款的损害补偿费用，其它人使用用户的电脑、帐号和其它知识产权的追索费。
12．结束服务
用户或<%=application("sitename")%>网站可随时根据实际情况中断一项或多项服务。<%=application("sitename")%>网站不需对任何个人或第三方负责而随时中断服务。用户若反对任何服务条款的建议或对后来的条款修改有异议，或对<%=application("sitename")%>网站服务不满，用户可以行使如下权利：
(1) 不再使用<%=application("sitename")%>网站信息服务。
(2) 通知<%=application("sitename")%>网站停止对该用户的服务。
　　结束用户服务后，用户使用<%=application("sitename")%>网站服务的权利马上中止。从那时起，用户没有权利，<%=application("sitename")%>网站也没有义务传送任何未处理的信息或未完成的服务给用户或第三方。 
13．通告
　　所有发给用户的通告都可通过重要页面的公告或电子邮件或常规的信件传送。服务条款的修改、服务变更、或其它重要事件的通告都会以此形式进行。 
14．信息内容的所有权
　　<%=application("sitename")%>网站定义的信息内容包括：文字、软件、声音、相片、录象、图表；在广告中全部内容；<%=application("sitename")%>网站为用户提供的其它信息。所有这些内容受版权、商标、标签和其它财产所有权法律的保护。所以，用户只能在<%=application("sitename")%>网站和广告商授权下才能使用这些内容，而不能擅自复制、再造这些内容、或创造与内容有关的派生商品。 
15．法律
　　<%=application("sitename")%>网站信息服务条款要与中华人民共和国的法律解释一致。用户和<%=application("sitename")%>网站一致同意服从上海悦婴婴幼儿用品有限公司所在地有管辖权的法院管辖。如发生本网站服务条款与中华人民共和国法律相抵触时，则这些条款将完全按法律规定重新解释，而其它条款则依旧保持对用户的约束力。 </textarea></TD>
                          </TR>
					</table>
				
					  <TABLE width="99%" border=0 align="center" cellPadding=3 cellSpacing=1 class=black-pix12 >
                        <TBODY>
                          <TR bgColor=#ffffff> 
                            <TD height=28>&nbsp;</TD>
                            <TD height=28><strong>请正确填写<%=application("sitename")%>会员注册信息：</strong></TD>
                          </TR>
                          <TR bgColor=#ffffff> 
                            <TD width=21% height="35" align=right> <DIV class=black align=right>用户名：</DIV></TD>
                            <TD width=79%>
							<table cellSpacing=0 cellPadding=0 width="510" border=0><tr><td>
							<INPUT type="text" name="username" value="" style="FONT-WEIGHT: normal; FONT-SIZE: 9pt;height:22px; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" size=22 maxlength="20" onBlur="window.checkmem.location.href='check_member.asp?username='+document.theForm.username.value"> 
							<FONT color=#ff0000>* </FONT><FONT color=#666666>4-20个字符(a-z，A-Z，0-9)</FONT>
							</td><td>
							<iframe id=checkmem name=checkmem src="check_member.asp" width=180 height=23 marginheight="2" marginwidth="0" frameSpacing="0" frameborder="0" scrolling=no></iframe>
							</td></tr></table>
							</TD>
                          </TR>
                          <TR bgColor=#ffffff> 
                            <TD width=21% height="35" align=right> <P class=black align=right>登录密码：</P></TD>
                            <TD width=79%> <INPUT type="password" name="password" value=""  size=22 maxlength="20" 
                  style="FONT-WEIGHT: normal; FONT-SIZE: 9pt;height:22px; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" onFocus="window.checkmem.location.href='check_member.asp?username='+document.theForm.username.value"> 
                              <FONT 
                  color=#ff0000>*</FONT> </TD>
                          </TR>
                          <TR bgColor=#ffffff> 
                            <TD width=21% height="35" align=right></TD>
                            <TD width=79%><FONT color=#666666>6-20个字符（0-9，a-z，A-Z，下划线） 
                              密码区分大小写；不能与用户名相同。建议您设置一个便于记忆且不易被别人猜到的密码。请牢记密码。</FONT></TD>
                          </TR>
                          <TR bgColor=#ffffff> 
                            <TD width=21% height="35" align=right class=black> 
                              <DIV align=right>密码确认：</DIV></TD>
                            <TD width=79%> <INPUT type="password" name="password2" value=""  size=22  
                  style="FONT-WEIGHT: normal; FONT-SIZE: 9pt;height:22px; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal"> <FONT 
                  color=#ff0000>*</FONT> </TD>
                          </TR>
                          <TR bgColor=#ffffff> 
                            <TD height="35" align=right class=black>电子邮件：</TD>
                            <TD><INPUT name="email"  
                  style="FONT-WEIGHT: normal; FONT-SIZE: 9pt;height:22px; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" 
                  size=45 maxlength="50"> 
                            *<FONT color=#666666>请正确填写您常用的电子邮件地址。</font></TD>
                          </TR>
                          <TR bgColor=#ffffff> 
                            <TD width=21% height="35" align=right class=black> 
                              <DIV align=right>密码提示问题：</DIV></TD>
                            <TD width=79%> <SELECT style="width:160px;FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; HEIGHT: 21px; FONT-VARIANT: normal" size=1 name="password_prompt">
                                <OPTION value="" selected>请选择问题</OPTION>
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
                  color=#ff0000>*</FONT> <font color="#666666">用于取回密码（如：您的初中班主任是？）</font></TD>
                          </TR>
                          <TR bgColor=#ffffff> 
                            <TD width=21% height="35" align=right class=black> 
                              <DIV align=right>密码提示答案：</DIV></TD>
                            <TD width=79%> <INPUT  
                  style="FONT-WEIGHT: normal; FONT-SIZE: 9pt;height:22px; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" 
                  size=22 name=password_Answer> <FONT 
                  color=#ff0000>*</FONT> <font color="#666666">当忘记密码后，点击“忘记密码”正确回答此答案可取回密码</font></TD>
                          </TR>

                          <TR bgColor=#ffffff> 
                            <TD width=21% height="35" align=right>真实姓名：</TD>
                            <TD width=79%> <INPUT  
                  style="FONT-WEIGHT: normal; FONT-SIZE: 9pt;height:22px; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" 
                  size=22 name=truename> <FONT color=#ff0000>*</FONT> <font 
                  color=#666666>请留下您的真实姓名，便于我们与您的联系。</font></TD>
                          </TR>
                          <TR bgColor=#ffffff> 
                            <TD width=21% height="35" align=right class=black> 
                              <DIV align=right>省份/城市/地区：</DIV></TD>
                            <TD width=79%>
							<script language="javascript" src="pcasunzip.js"></script>
							<select name="province"  style="WIDTH: 130px"></select> <select name="city"  style="WIDTH: 140px"></select> <select name="area"  style="WIDTH: 140px"></select>
							<script language="javascript" defer>
							new PCAS("province","city","area");
							</script>
							</TD>
                          </TR>

                          <TR bgColor=#ffffff> 
                            <TD width=21% height="47" align=right class=black> 
                            <DIV align=right>详细地址：</DIV></TD>
                            <TD class=black width=79%> <INPUT   
                  style="FONT-WEIGHT: normal; FONT-SIZE: 9pt;height:22px; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" 
                  size=45 name=address> <FONT color=#ff0000>*</FONT><font color="#0066CC">本地址可能用于商品送货上门或邮寄地址。</font> 
                              <br>
                            <font color="#666666">例如：东圃镇建设路2号新城花园E幢2203室</font></TD>
                          </TR>
                          <TR bgColor=#ffffff> 
                            <TD width=21% height="35" align=right class=black> 
                              <DIV align=right>邮政编码：</DIV></TD>
                            <TD class=black width=79%> <INPUT name=postno  
                  style="FONT-WEIGHT: normal; FONT-SIZE: 9pt;height:22px; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" 
                  size=22 maxlength="15"> </TD>
                          </TR>
                          <TR bgColor=#ffffff> 
                            <TD width=21% height="35" align=right class=black> 
                              <DIV align=right>固定电话：</DIV></TD>
                            <TD width=79%> <INPUT  
                  style="FONT-WEIGHT: normal; FONT-SIZE: 9pt;height:22px; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" 
                  size=22 name=telphone1> <FONT color=#ff0000>*</FONT> <FONT 
                  color=#666666>含区号和电话号码，或分机号 如:020-83219880、020-85698890-21</FONT></TD>
                          </TR>
                          <TR bgColor=#ffffff> 
                            <TD width=21% height="35" align=right class=black> 
                              <DIV align=right>移动电话：</DIV></TD>
                            <TD class=black width=79%> <INPUT  
                  style="FONT-WEIGHT: normal; FONT-SIZE: 9pt;height:22px; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" 
                  size=22 name='mobile'> </TD>
                          </TR>



                          <TR vAlign=top bgColor=#ffffff> 
                            <TD align=right colSpan=2> <TABLE   cellSpacing=0 cellPadding=0 width="100%" 
                  border=0>
                                <TBODY>
                                  <TR bgColor=#ffffff> 
                                    <TD align=right width="91">&nbsp;</TD>
                                    <TD align=right width=62>&nbsp;</TD>
                                    <TD align=right width="575">&nbsp;</TD>
                                  </TR>
                                  <TR bgColor=#ffffff> 
                                    <TD align=right width="91">&nbsp;                                    </TD>
                                    <TD width=62>&nbsp;</TD>
                                    <TD width="575"><input class=register_submit type=submit value=" " name=Register_Button>
                                      <input   type=hidden value=add name=eaction></TD>
                                  </TR>
                                  <TR bgColor=#ffffff> 
                                    <TD align=right colSpan=3 height=15></TD>
                                  </TR>
                                </TBODY>
                              </TABLE></TD>
                          </TR>
                        </TBODY>
                      </TABLE>
					  <br>
					   </TD>
                  </TR>
                </TBODY>
              </FORM>
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
<%
conn.close
set conn=nothing
%>
</body>
</html>



