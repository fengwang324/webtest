<!--#include file="conn.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>礼品列表</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<style>
<!--
td{font-size:9pt;line-height: 160%;}
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
</head>

<body bgcolor="#fcfcfc">

<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" id="tabcss">
  <tr bgcolor="#E1EFFF"> 
    <td height="25" colspan="7"><b>礼品设置：</b></td>
  </tr>
  <tr bgcolor="#eeeeee"> 
    <td height="25" width=16%>
<div align="center">图片</div></td>
    <td width=10%> 
      <div align="center">编号</div></td>
    <td width=27%>礼品名称</td>
    <td width=7%> <div align="center">所需积分</div></td>
    <td width=7%><div align="center">是否显示</div></td>
    <td width=7%> <div align="center">顺序</div></td>
    <td width=25%> <div align="center"><b><a href='lipinadd.asp'>增 加</a></b></div></td>
  </tr>
  <%
i=0

sql="select * from x_lipin order by num"

set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof or rs.eof then
	response.write "<tr><td colspan=7><div align=center>没有记录! </div></td></tr>"
else
	do while not rs.eof 
	
	id=rs("id")
	picpath=rs("picpath")
	'picpath2=rs("picpath2")
	lipinno=rs("lipinno")
	lipinname=rs("lipinname")
	zhifen=rs("zhifen")
	showflag=rs("showflag")
	num=rs("num")
	'memo=rs("memo")
	
	
	
%>
  <tr onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
    <td align=center><img width="90" src='<%=picpath%>' border='0'></td>
    <td><%=lipinno%></td>
    <td><b><%=lipinname%></b></td>
    <td><div align="right"><%=zhifen%>&nbsp;</div></td>
    <td><div align="center"><input type="checkbox" name="showflag" value="1" <%if showflag="1" then response.write "checked"%>></div></td>
    <td><%=num%></td>
    <td align="center">
	<a href="../lipinshow.asp?id=<%=id%>" target="_blank">查看</a>&nbsp;&nbsp;&nbsp;
	 <a href="lipinadd.asp?id=<%=id%>">修改</a>&nbsp;&nbsp;&nbsp;&nbsp; 
      <a href="lipindel.asp?id=<%=id%>" onClick="return confirm('真的删除此记录?')">删除</a> 
    </td>
  </tr>
  <%
	rs.movenext
	i=i+1
	loop
end if

rs.close
set rs=nothing
call connclose

%>
</table>
<br><br>
</body>
</html>


