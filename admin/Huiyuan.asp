<!--#include file="conn.asp"-->
<!--#include file="check4.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>会员列表</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<style>
<!--
td{font-size:9pt;line-height:140%;}
body {
	margin-top: 5px;
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
	line-height: 21px;
	padding-left: 3px;
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

<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" id="tabcss">
  <tr bgcolor="#BFDFFF">
    <form name="form1" method="post" action="huiyuan.asp">
      <td height=35 colspan=11> <font color=#0000ff><strong></strong></font>会员查找： 
        <select name="selectkind">
          <option value="username">用户名</option> 
          <option value="id">会员编号</option>  
          <option value="truename">真实姓名</option>
          <option value="province">省份</option>
          <option value="city">城市</option>
          <option value="telphone1">联系电话1</option>
          <option value="telphone2">联系电话2</option>
          <option value="fax">传真</option>
          <option value="mobile">移动电话</option>
          <option value="address">联系地址</option>
          <option value="postno">邮政编码</option>
          <option value="email">电子邮件</option>
          <option value="addtime">注册时间</option>
        </select> <input name="keyword" type="text" size="15" maxlength="15"> 
        <input type="submit" name="Submit" value="查 找">&nbsp;&nbsp;&nbsp;&nbsp;<a href=huiyuan.asp>全部会员</a></td>
    </form>
  </tr>
  <tr bgcolor="#E1E1E1"> 
    <td> 
<div align="center">编号</div></td>
    <td height="24"> 
<div align="center">用户名</div></td>
    <td> 
<div align="center">真实姓名</div></td>
    <td> 
<div align="center">联系电话</div></td>

    <td > 
<div align="center">移动电话</div></td>
    <td > 
<div align="center">会员等级</div></td>	 
    <td > 
<div align="center">账户余款</div></td>
    <td > 
<div align="center">会员积分</div></td>
    <td > 
      <div align="center">注册时间</div></td>
    <td > 
      <div align="center">最后登陆时间</div></td> 
    <td> </td>
  </tr>
  <%
i=0
nowpage=request("page")
selectkind=trim(request("selectkind"))
keyword=trim(request("keyword"))

sql0="select *,(select sum(czmoney) from sh_chongzhi where sh_chongzhi.customid=a.id and confirmflag='是') AS sumcz, (select sum(total) from x_bill where x_bill.customid=a.id and paytype='账户余额') AS sumkj, (IIf(IsNull(sumcz),0,sumcz)-IIf(IsNull(sumkj),0,sumkj)) AS leftye, "&_
	"(select sum(billzhf) from x_view_bill where x_view_bill.username=a.username and sendornot='已发货') AS sumzhf, (select sum(kjzhf) from x_huanlipin where x_huanlipin.username=a.username) AS sumkjzhf, (IIf(IsNull(sumzhf),0,sumzhf)-IIf(IsNull(sumkjzhf),0,sumkjzhf)) AS leftzhf  from x_huiyuan as a "     

if keyword="" then
	
	sql=sql0 & " where id>2 order by id desc"
else
	if selectkind="id" then
		sql=sql0 & " where id>2 and id = "&keyword&" order by id desc"
	else
		sql=sql0 & " where id>2 and "&selectkind&" like '%"&keyword&"%' order by id desc"
	end if
end if


set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof or rs.eof then
	response.write "<div align=center><br>没有会员记录!</div>"
else
rs.pagesize=20
rs.absolutepage=1
if request("page")<>"" then rs.absolutepage=request("page")
rowcount=rs.pagesize
	do while not rs.eof and rowcount>0
		id=rs("id")
		username=rs("username")
		vipno=addspace(rs("vipno"))
		truename=rs("truename") 
		telphone1=rs("telphone1") 
		telphone2=addspace(rs("telphone2"))
		mobile=addspace(rs("mobile")) 
		address=rs("address") 
		addtime=rs("addtime") 
		customkind=rs("customkind")
	  if customkind="2" then 
		customkind="普通会员"
	  elseif customkind="3" then 
		customkind="铜级会员"
	  elseif customkind="4" then 
		customkind="银级会员"
	  elseif customkind="5" then 
		customkind="金级会员"
	  elseif customkind="6" then 
		customkind="钻级会员"
	  else
	    customkind="&nbsp;"
	  end if
	  lasttime=rs("lasttime")
	  leftye=rs("leftye")
	  leftzhf=rs("leftzhf")
	  
%>
  <tr onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
    <td align="center"><%=id%></td>
    <td height="21" align="center"><a href="huiyuandetail.asp?id=<%=id%>"><%=username%></a></td>
    
    <td><%=truename%></td>
    <td><%=telphone1%></td>
    
    <td><%=mobile%></td>
	<td align=center><%=customkind%></td>
    <td align=right><%=leftye%>&nbsp;</td>
	<td align=right><%=leftzhf%>&nbsp;</td>
    <td style="WIDTH: 125; WORD-WRAP: break-word;line-height:140%;"><%=addtime%></td>
	<td style="WIDTH: 125; WORD-WRAP: break-word;line-height:140%;"><%=lasttime%></td>
	
    <td align="center" width="4%"><a href="huiyuandel.asp?id=<%=id%>" onClick="return confirm('是否删除会员“<%=username%>”?')">删除</a></td>
  </tr>
  <%
	rs.movenext
	rowcount=rowcount-1
	i=i+1
	loop
end if
%>
</table>
<%
response.write "<table border=0 width='95%' align=center><tr><td align=right height=30>"
		backnexturl="keyword="&keyword&"&selectkind="&selectkind&"&none=&"
		if nowpage="" then nowpage="1"
		lastpage=rs.pagecount
		nowpath=request.ServerVariables("SCRIPT_NAME")
		response.write "共"&rs.recordcount&"条 每页20条 现在第"&nowpage&"/"&rs.pagecount&"页&nbsp; "
		response.write "&nbsp;<input type='button' name='Button1' value='首页' onclick=window.location.href='"&nowpath&"?"&backnexturl&"page=1'>&nbsp;"
		if nowpage>1 then
			response.write"<input type='button' name='Button2' value=上页 onclick=window.location.href='"&nowpath&"?"&backnexturl&"page="&nowpage-1&"'>&nbsp;"
		else
			response.write"<input type='button' name='Button3' value=上页>&nbsp;"
		end if
		if clng(nowpage) < clng(lastpage) then
			response.write"<input type='button' name='Button4' value=下页 onclick=window.location.href='"&nowpath&"?"&backnexturl&"page="&nowpage+1&"'>&nbsp;"
		else
			response.write"<input type='button' name='Button5' value=下页>&nbsp;"
		end if
		response.write"<input type='button' name='Button6' value=尾页 onclick=window.location.href='"&nowpath&"?"&backnexturl&"page="&clng(lastpage)&"'>&nbsp;"
		response.write "</td><form name='form1' method='post' action="&nowpath&"?"&backnexturl&"><td>转到第<input type=text name='page'size=5>页 <input type='submit' name='Submit' value='确定'>"
		response.write "&nbsp;&nbsp;<a href='javascript:window.scroll(0,0)'>TOP↑</a>"
response.write "</td></form></tr></table>"

response.cookies("hypage")=nowpath&"?"&backnexturl&"page="&nowpage
%><br><br>

</body>
</html>
<%
rs.close
set rs=nothing
call connclose
%>


