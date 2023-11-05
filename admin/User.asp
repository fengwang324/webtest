<!--#include file="conn.asp"-->
<!--#include file="check8.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>用户表</title>
<meta http-equiv=Content-Type content=text/html; charset=utf-8>
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
-->
</style>
<%
dim sql4,rs4,id,num,pkid


Submit2=trim(request.form("Submit2"))

if Submit2="保存" then

	for j=1 to request.form("maxi")
		id = request.form("id"&j)
		s_user = request.form("s_user"&j)
		s_pwd = request.form("s_pwd"&j)
		self = request.form("self"&j)
		memo = request.form("memo"&j)
		
		if isnumeric(id) and s_user<>"" then
			sql2="update [Admin] set s_user='"&s_user&"',s_pwd='"&s_pwd&"',self1='"&self&"',memo='"&memo&"' where id="&id
			conn.execute(sql2)
		elseif id="" and s_user<>"" then
			sql2="insert into [Admin] (s_user,s_pwd,self1,memo) values('"&s_user&"','"&s_pwd&"',"&self&",'"&memo&"')"
			conn.execute(sql2)
		elseif isnumeric(id) and s_user="" then
			sql2="delete * from [Admin] where id="&id
			conn.execute(sql2)
		else
		end if
	next

	response.write "<script language=JavaScript>" & "alert('保存成功!');"& "window.location.href='user.asp';" & "</script>"
end if

%>
</head>

<body bgcolor="#fcfcfc">          
<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0" id="tabcss">
  <tr bgcolor="#E1EFFF"> 
    <td height=30 colspan=6>&nbsp;<b>后台用户表：</b></td>
  </tr>
  <tr bgcolor='#F5F5F5'> 
    <td width="10%" height='23' align='right'> <div align="center"></div></td>
    <td width="20%" align='right'> <div align="center">用户名</div></td>
    <td width="19%"> <div align="center">密码</div></td>
    <td width="17%"> <div align="center">级别</div></td>
    <td width="21%"> <div align="center">说明</div></td>
  </tr>
  <form name="form1" method="post" action="user.asp">
    <%
		
	sql4="select * from [Admin] order by id "
	set rs4=server.createobject("adodb.recordset")
	rs4.open sql4,conn,1,1
	if rs4.bof or rs4.eof then
		response.write "没有符合条件记录！"
	else
		i=1
		do while not rs4.eof
			id=rs4("id")
			s_user=rs4("s_user")
			s_pwd=rs4("s_pwd")
			self=rs4("self1")
			memo=rs4("memo")
%>
    <%
	  if cstr(id)="1" then
	  %>
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'><div align="center"><%=i%></div></td>
      <td height='23' align='right'><div align="center"> 
          <input name="id<%=i%>" type="hidden" value="<%=id%>">
          <input name="s_user<%=i%>" type="text" value="<%=s_user%>">
        </div></td>
      <td><div align="center"> 
          <input name="s_pwd<%=i%>" type="password" value="<%=s_pwd%>">
        </div></td>
      <td><div align="center"> 
          <input name="backop" type="text" value="超级管理员"  readonly style="BACKGROUND-COLOR: #eeeeee;width:90px">
          <input name="self<%=i%>" type="hidden" value="8">
        </div></td>
      <td><div align="center"> 
          <input name="memo<%=i%>" type="text" value="<%=memo%>">
        </div></td>
    </tr>
    <%else%>
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'><div align="center"><%=i%></div></td>
      <td height='23' align='right'><div align="center"> 
          <input name="id<%=i%>" type="hidden" value="<%=id%>">
          <input name="s_user<%=i%>" type="text" value="<%=s_user%>" >
        </div></td>
      <td><div align="center"> 
          <input name="s_pwd<%=i%>" type="password" value="<%=s_pwd%>">
        </div></td>
      <td><div align="center"> 
          <select name="self<%=i%>" size="1" value="<%=self%>">
            <option value=""></option>
            <option value="2" <%if self="2" then response.write " selected"%> >信息管理员</option>
            <option value="4" <%if self="4" then response.write " selected"%> >订单管理员</option>
            <option value="8" <%if self="8" then response.write " selected"%> >超级管理员</option>
          </select>
        </div></td>
      <td><div align="center"> 
          <input name="memo<%=i%>" type="text" value="<%=memo%>">
        </div></td>
    </tr>
    <%end if%>
    <%
		i=i+1
		rs4.movenext
		loop			
	end if
	rs4.close
	set rs4=nothing		
%>
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'><div align="center"><%=i%></div></td>
      <td height='23' align='right'><div align="center"> 
          <input name="id<%=i%>" type="hidden" value="">
          <input name="s_user<%=i%>" type="text" value="">
        </div></td>
      <td><div align="center"> 
          <input name="s_pwd<%=i%>" type="password" value="">
        </div></td>
      <td><div align="center"> 
          <select name="self<%=i%>" size="1" value="">
            <option value=""></option>
            <option value="2">信息管理员</option>
            <option value="4">订单管理员</option>
            <option value="8">超级管理员</option>
          </select>
        </div></td>
      <td><div align="center"> 
          <input name="memo<%=i%>" type="text" value="<%=memo%>">
        </div></td>
    </tr>
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='35' colspan="2" align='right'>&nbsp;</td>
      <input type="hidden" name="maxi" value="<%=i%>">
      <td colspan="4"><input type="submit" name="Submit2" value=" 保存 "> &nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="reset" name="Submit3" value=" 重置 "> </td>
    </tr>
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';">
      <td height='35' colspan="2" align='right'><div align="left">如果要增加操作员，请在最后一行填写，并保存。每次保存后，都会自动 
          在最后新增一个空行，可让您不断增加新的操作员的。</div></td>
      <td colspan="4">1 超级管理员-拥有全部功能管理权。<br>
        2 信息管理员-拥有除会员及订单外的管理员权，如网站基本信息，栏目管理，商品管理，支付管理等前台所能看到的信息。<br>
        3 订单管理员-拥有会员及订单管理权。</td>
    </tr>
  </form>
</table>
<br><br><br>

</body>
</html>
<%
call connclose
%>

