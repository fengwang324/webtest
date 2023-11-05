<!--#include file="conn.asp"-->
<!--#include file="check2.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>左导航栏列表</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<style>
<!--
td{font-size:9pt}
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
	line-height: 24px;
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
<table height="370" width="100%" align="center" >
  <tr>
    <td  valign="top"> 
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" id="tabcss">
    <tr bgcolor="#BFDFFF"> 
          <td height="25" align="left" colspan="6"> 
				<a href="conlist.asp"><strong>全部记录</strong></a>&nbsp;&nbsp;
				<strong> 
				
				<%
				set rs3=server.CreateObject("adodb.recordset")
				sql3="select * from e_left where (menukind='普通文章列表' or menukind='文章列表并指定链接' or menukind='单个说明')  order by bigmenu desc,l_num "
				rs3.open sql3,conn,1,1
				if rs3.bof or rs3.eof then
				else
				do while not rs3.eof
				l_id=rs3("l_id")
				l_title=rs3("l_title")
				bigmenu=rs3("bigmenu")
				if bigmenu="是" then
				response.write "<a href='conlist.asp?l_id="&l_id&"'><font color='#ff7777'>"&l_title&"</font></a> &nbsp;"
				else
				response.write "<a href='conlist.asp?l_id="&l_id&"'><font color='#7777ff'>"&l_title&"</font></a> &nbsp;"
				end if
				rs3.movenext
				loop
				end if
				rs3.close
				set rs3=nothing
				%>
				</strong>
			  </td>
  </tr>

        <tr> 
          <td><span class="style3">导航栏</span></td>
   			 <form name="form1" method="post" action="conlist.asp">
            <td><span class="style3"> 标题　　关键字 
              <input name="keyword" type="text" size="15" maxlength="15">
              <input type="submit" name="Submit2" value="搜索">     
              <input type="button" name="Submit3" value="新增此类新闻" onClick="window.location.href='conadd.asp?newskind=<%=request("l_id")%>'">
              </span></td>
    		</form>
          <td align=center>顺序</td>
          <td> </td>
		  <td> </td>
		  <td> </td>
		  
  </tr>

<%
old="0"
i=0
nowpage=request("page")
keyword=request("keyword")
newskind=request("l_id")

if newskind<>"" then
	sql="select * from article where l_id = "&newskind&" order by l_title desc,c_num desc,c_addtime desc,c_id desc"
else
	if keyword="" then
		sql="select * from article order by l_title desc,c_num desc,c_addtime desc,c_id desc"
	else
		sql="select * from article where c_title like '%"&keyword&"%' order by l_title desc,c_num desc,c_addtime desc,c_id desc"
	end if
end if

set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof or rs.eof then
	response.write "<div align=center><br>没有记录!</div>"
else
rs.pagesize=15
rs.absolutepage=1
if request("page")<>"" then rs.absolutepage=request("page")
rowcount=rs.pagesize
	do while not rs.eof and rowcount>0
	c_id=rs("c_id")
	num=rs("id")
	title=rs("title")
	l_title=rs("l_title")&"&nbsp;"
	c_title=rs("c_title")
	c_num=rs("c_num")
	c_addtime=rs("c_addtime")

if cstr(old)<>cstr(l_title) and i>0 then 
	response.write "<tr><td colspan=5 height=10 bgcolor=#f7f7f7></td></tr>"
	old = l_title
else
	old = l_title
end if
%>
        <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';">
          <td><%=l_title%>  </td>
          <td style="WIDTH: 550; WORD-WRAP: break-word"><%=c_title%> <font color=#999999>(<%=c_addtime%>)</font></td>
		  <td align=center><%=c_num%></td>
          <td align=center><a href="conshow.asp?c_id=<%=c_id%>&newskind=<%=newskind%>">修改</a></td>
          <td align=center><a href="../show_all.asp?pkid=<%=num%>&c_id=<%=c_id%>" target="_blank">查看</a> </td>
          <td align=center><div align="center"><a href="condel.asp?c_id=<%=c_id%>&newskind=<%=newskind%>" onClick="return confirm('是否删除 <%=c_title%>?')">删除</a></div></td>
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
response.write "<table border=0 width='90%' align=center><tr><td align=right>"
		backnexturl="keyword="&keyword&"&l_id="&newskind&"&none=&"
		if nowpage="" then nowpage="1"
		lastpage=rs.pagecount
		nowpath=request.ServerVariables("SCRIPT_NAME")
		response.write "共"&rs.recordcount&"条记录&nbsp;每页15条&nbsp;现在 "&nowpage&"/"&rs.pagecount &" 页&nbsp;"
		response.write "<input type='button' name='Button1' value='首页' onclick=window.location.href='"&nowpath&"?"&backnexturl&"page=1'>&nbsp;"
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
		
		response.write "&nbsp;&nbsp;<a href='javascript:window.scroll(0,0)'>TOP↑</a>"
response.write "</td></tr></table>"
%><br><br>
    </td>
  </tr></table>
</body>
</html>
<%
rs.close
set rs=nothing
call connclose
%>




