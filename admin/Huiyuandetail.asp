<!--#include file="conn.asp"-->
<!--#include file="check4.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>会员资料</title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-bottom: 0px;
	margin-right: 0px;
}
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



   if(confirm("现在真的要修改该网上会员的资料吗?"))
      return true
   else
      return false;
}
</SCRIPT>
</head>

<body>
<br>
<TABLE width="740" border=0 align="center" cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
  <tr><td>
<TABLE width=740 height=278 cellSpacing=0 cellPadding=0 align=center bgColor=#ffffff border=0>
        <TBODY>
          <TR> 
            <TD valign="top" bgColor=#ffffff> 
              <%
	sql="select * from x_huiyuan where id="&request.QueryString("id")&" "
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.bof or rs.eof then
	else
		username=rs("username")
		password=rs("password") 
		password_prompt=rs("password_prompt") 
		password_Answer=rs("password_Answer") 
		vipno=rs("vipno") 
		truename=rs("truename") 
		province=rs("province") 
		city=rs("city") 
		telphone1=rs("telphone1") 
		telphone2=rs("telphone2") 
		fax=rs("fax") 
		mobile=rs("mobile") 
		address=rs("address") 
		postno=rs("postno") 
		email=rs("email") 
		babysex=rs("babysex") 
		babybirthday=rs("babybirthday")
		addtime=rs("addtime")
		lasttime=rs("lasttime")
		s_stat=rs("s_stat")
		customkind=rs("customkind")
		
	end if
	rs.close
	set rs=nothing
	response.Cookies("customid")=request.QueryString("id")
	response.Cookies("username9")=username
%>
              <TABLE width="99%" border=0 align="center" cellPadding=3 cellSpacing=1 bgColor=#dddddd class=black-pix12 id=Table3>
                <FORM id=theForm name="theForm" action="huiyuandetailadmin.asp" method="post" onSubmit="return Juge()">
                  <TBODY>
                    <TR bgColor=#EEEEEE> 
                      <TD colSpan=2 height=28><strong><b><font color=#ff0000><%=username%></font></b> 
                        个人资料：</strong>&nbsp;&nbsp;&nbsp;&nbsp;<b>&nbsp;&nbsp; 
						
						　
                        </b></TD>
                    </TR>
                    <TR bgColor=#ffffff> 
                      <TD align=right width=20%> <DIV class=black align=right>用户名：</DIV></TD>
                      <TD width=80%> <b><font color=#ff0000><%=username%></font></b></TD>
                    </TR>
                    <TR bgColor=#ffffff> 
                      <TD align=right width=20%> <P class=black align=right>登录密码：</P></TD>
                      <TD width=80%> <INPUT name=password value="<%=password%>" type=password id=Password1 
                  style="FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" size=20 maxlength="20"> 
                        <FONT 
                  color=#ff0000>*</FONT> </TD>
                    </TR>

                    <TR bgColor=#ffffff> 
                      <TD class=black align=right width=20%> <DIV align=right>密码确认：</DIV></TD>
                      <TD width=80%> <INPUT id=Password2 
                  style="FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" 
                  type=password size=20 name=password2  value="<%=password%>"> 
                        <FONT 
                  color=#ff0000>*</FONT> </TD>
                    </TR>
                    <TR bgColor=#ffffff> 
                      <TD class=black align=right width=20%> <DIV align=right>密码提示问题：</DIV></TD>
                      <TD width=80%> <SELECT style="width:160px;FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; HEIGHT: 21px; FONT-VARIANT: normal" size=1 name="password_prompt">
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
                      <TD class=black align=right width=20%> <DIV align=right>密码提示答案：</DIV></TD>
                      <TD width=80%> <INPUT id=Text4 
                  style="FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" 
                  size=20 name=password_Answer  value="<%=password_Answer%>"> 
                        <FONT 
                  color=#ff0000>*</FONT> <font color="#666666">当忘记密码，点击“忘记密码”,用此答案即可取回密码。</font></TD>
                    </TR>
                    <TR bgColor=#ffffff> 
                      <TD class=black align=right>电子邮件：</TD>
                      <TD><INPUT name="email" value="<%=email%>" id=email 
                  style="FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" 
                  size=40 maxlength="50"></TD>
                    </TR>
                    <TR bgColor=#ffffff> 
                      <TD align=right width=20%>真实姓名：</TD>
                      <TD width=80%> <INPUT id=Text2 
                  style="FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" 
                  size=20 name=truename  value="<%=truename%>"> <FONT color=#ff0000>*</FONT> 
                        <font 
                  color=#666666>请留下您的真实姓名，便于我们与您的联系以及给您发货。</font></TD>
                    </TR>
                    <TR bgColor=#ffffff> 
                      <TD class=black align=right width=20%> <DIV align=right>省份：</DIV></TD>
                      <TD width=80%> <SELECT id=Select1 name=province style="width:132px;">
                          <OPTION value="">请选择</OPTION>
                          <OPTION value="<%=province%>" selected><%=province%></OPTION>
                          <OPTION value="广东">广东</OPTION>
                          <OPTION value="北京">北京</OPTION>
                          <OPTION value="上海">上海</OPTION>
                          <OPTION value="天津">天津</OPTION>
                          <OPTION value="安徽">安徽</OPTION>
                          <OPTION value="重庆">重庆</OPTION>
                          <OPTION value="福建">福建</OPTION>
                          <OPTION value="甘肃">甘肃</OPTION>
                          <OPTION value="广西">广西</OPTION>
                          <OPTION value="贵州">贵州</OPTION>
                          <OPTION value="海南">海南</OPTION>
                          <OPTION value="河北">河北</OPTION>
                          <OPTION value="河南">河南</OPTION>
                          <OPTION value="黑龙江">黑龙江</OPTION>
                          <OPTION value="湖北">湖北</OPTION>
                          <OPTION value="湖南">湖南</OPTION>
                          <OPTION value="吉林">吉林</OPTION>
                          <OPTION value="江苏">江苏</OPTION>
                          <OPTION value="江西">江西</OPTION>
                          <OPTION value="辽宁">辽宁</OPTION>
                          <OPTION value="内蒙古">内蒙古</OPTION>
                          <OPTION value="宁夏">宁夏</OPTION>
                          <OPTION value="青海">青海</OPTION>
                          <OPTION value="山东">山东</OPTION>
                          <OPTION value="山西">山西</OPTION>
                          <OPTION value="陕西">陕西</OPTION>
                          <OPTION value="四川">四川</OPTION>
                          <OPTION value="西藏">西藏</OPTION>
                          <OPTION value="新疆">新疆</OPTION>
                          <OPTION value="云南">云南</OPTION>
                          <OPTION value="浙江">浙江</OPTION>
                          <OPTION value="香港">香港</OPTION>
                          <OPTION value="澳门">澳门</OPTION>
                          <OPTION value="台湾">台湾</OPTION>
                          <OPTION value="其它">其它</OPTION>
                        </SELECT> </TD>
                    </TR>
                    <TR bgColor=#ffffff> 
                      <TD class=black align=right width=20%> <DIV align=right>城市：</DIV></TD>
                      <TD width=80%> <INPUT name=city value="<%=city%>" id=Text8 size="20" style="FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal"></TD>
                    </TR>
                    <TR bgColor=#ffffff> 
                      <TD class=black align=right width=20%> <DIV align=right>联系地址：</DIV></TD>
                      <TD class=black width=80%> <INPUT id=Text11 
                  style="FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" 
                  size=40 name=address value="<%=address%>"> <FONT color=#ff0000>*</FONT><font color="#0066CC">本地址可能会用于商品送货上门或邮寄地址</font></TD>
                    </TR>
                    <TR bgColor=#ffffff> 
                      <TD class=black align=right width=20%> <DIV align=right>邮政编码：</DIV></TD>
                      <TD class=black width=80%> <INPUT name=postno value="<%=postno%>" id=Text13 
                  style="FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" 
                  size=20 maxlength="15"> </TD>
                    </TR>
					
                    <TR bgColor=#ffffff> 
                      <TD class=black align=right width=20%> <DIV align=right>联系电话：</DIV></TD>
                      <TD width=80%> <INPUT id=Text9 
                  style="FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" 
                  size=20 name=telphone1  value="<%=telphone1%>"> <FONT color=#ff0000>*</FONT> 
                        <FONT 
                  color=#666666>请填写您的固定电话，含区号和电话号码</FONT> </TD>
                    </TR>
                    <TR bgColor=#ffffff> 
                      <TD class=black align=right width=20%> <DIV align=right>移动电话：</DIV></TD>
                      <TD class=black width=80%> <INPUT id='Text14' 
                  style="FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" 
                  size=20 name='mobile'  value="<%=mobile%>"> </TD>
                    </TR>


                    <TR bgColor=#f5f5f5> 
                      <TD class=black align=right height=30>会员级别：</TD>
                      <TD> 
                        <SELECT name="customkind" size=1 style="FONT-WEIGHT: normal; FONT-SIZE: 9pt;width:132px; LINE-HEIGHT: normal; FONT-STYLE: normal; HEIGHT: 21px; FONT-VARIANT: normal">
                          <OPTION value="2" <%if customkind="2" then response.write "selected"%> >普通会员</OPTION>
                          <OPTION value="3" <%if customkind="3" then response.write "selected"%> >铜级会员</OPTION>
                          <OPTION value="4" <%if customkind="4" then response.write "selected"%> >银级会员</OPTION>
                          <OPTION value="5" <%if customkind="5" then response.write "selected"%> >金级会员</OPTION>
                          <OPTION value="6" <%if customkind="6" then response.write "selected"%> >钻级会员</OPTION>
                        </SELECT></TD>
                    </TR>
                    <TR bgColor=#ffffff> 
                      <TD class=black align=right>账户余额：</TD><TD>
					  <%
						sql6="SELECT id, username,(select sum(czmoney) from sh_chongzhi where sh_chongzhi.customid=a.id and confirmflag='是') AS sumcz, "&_
						"(select sum(total) from x_bill where x_bill.customid=a.id and paytype='账户余额') AS sumkj,  "&_
						"(IIf(IsNull(sumcz),0,sumcz)-IIf(IsNull(sumkj),0,sumkj)) AS leftye  "&_
						"FROM x_huiyuan AS a where username='"&username&"' "
						
						set rs6=server.createobject("adodb.recordset")
						rs6.open sql6,conn,1,1
						if rs6.bof or rs6.eof then
							response.write "＄<input type='text' name='leftye' value='0' readonly style='width:85;border:0px;BACKGROUND-COLOR: #ffffff;FONT-WEIGHT: bold; FONT-SIZE: 17px; COLOR:#f00;FONT-FAMILY: Century Gothic'>"
						else
							response.write "＄<input type='text' name='leftye' value='"&rs6("leftye")&"' readonly style='width:85;border:0px;BACKGROUND-COLOR: #ffffff;FONT-WEIGHT: bold; FONT-SIZE: 17px; COLOR:#f00;FONT-FAMILY: Century Gothic'>"
						end if
						rs6.close
						set rs6=nothing
					  %>
					  </TD>
                    </TR>
                    <TR bgColor=#ffffff> 
                      <TD class=black align=right>剩余积分：</TD><TD>
					  <%
						sql6="SELECT id, username,(select sum(billzhf) from x_view_bill where x_view_bill.username=a.username and sendornot='已发货') AS sumzhf, "&_
						"(select sum(kjzhf) from x_huanlipin where x_huanlipin.username=a.username) AS sumkjzhf,  "&_
						"(IIf(IsNull(sumzhf),0,sumzhf)-IIf(IsNull(sumkjzhf),0,sumkjzhf)) AS leftzhf  "&_
						"FROM x_huiyuan AS a where username='"&username&"' "
						
						set rs6=server.createobject("adodb.recordset")
						rs6.open sql6,conn,1,1
						if rs6.bof or rs6.eof then
							response.write "<input type='text' name='leftzhf' value='0' readonly style='width:85;border:0px;BACKGROUND-COLOR: #ffffff;FONT-WEIGHT: bold; FONT-SIZE: 17px; COLOR:#f00;FONT-FAMILY: Century Gothic'>"
						else
							response.write "<input type='text' name='leftzhf' value='"&rs6("leftzhf")&"' readonly style='width:85;border:0px;BACKGROUND-COLOR: #ffffff;FONT-WEIGHT: bold; FONT-SIZE: 17px; COLOR:#f00;FONT-FAMILY: Century Gothic'>"
						end if
						rs6.close
						set rs6=nothing
					  %>
					  </TD>
                    </TR>
                    <TR bgColor=#f5f5f5> 
                      <TD class=black align=right height=30>注册时间：</TD>
                      <TD><%=addtime%>&nbsp;&nbsp;&nbsp;&nbsp;最后登陆时间：<%=lasttime%></TD>
                    </TR>
                    <TR vAlign=top bgColor=#ffffff> 
                      <TD align=right colSpan=2> <TABLE id=Table4 cellSpacing=0 cellPadding=0 width="100%" 
                  border=0>
                          <TBODY>
                            <TR bgColor=#ffffff> 
                              <TD align=right width="125">&nbsp;</TD>
                              <TD align=right width=20>&nbsp;</TD>
                              <TD align=right width="580">&nbsp;</TD>
                            </TR>
                            <TR bgColor=#ffffff> 
                              <TD align=right width="125"> 
<input type=hidden value=add name=eaction> <input name="id" type=hidden value="<%=request.QueryString("id")%>">                              </TD>
                              <TD width=20>&nbsp;</TD>
                              <TD width="580"> 
                                <input type="submit" name="Submit" value="更新会员资料" style="width:130px;height:25px; color:#ff0000;FONT-SIZE: 10pt; FONT-WEIGHT: bold">&nbsp;&nbsp;
                                <%call history()%>                              </TD>
                            </TR>
                            <TR bgColor=#ffffff> 
                              <TD align=right colSpan=3 height=37></TD>
                            </TR>
                          </TBODY>
                        </TABLE></TD>
                    </TR>
                  </TBODY>
                </form>
              </TABLE></TD>
          </TR>
        </TBODY>
      </TABLE>

    </td>
  </tr>

<TR><TD>

</TD></TR>
</table>
<br><br>
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

</body>
</html>

