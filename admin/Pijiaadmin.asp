<!--#include file="conn.asp"-->
<!--#include file="check4.asp"-->

<%
action=request.querystring("action")
id=request.querystring("id")
if action="del" and id<>"" then
conn.execute("delete * from z_pijia where id="&id)
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>会员列表</title>
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
.input3 {
      	BORDER-RIGHT: #666666 1px solid;BORDER-TOP: #666666 0px solid; height:23px;FONT-SIZE: 10pt; BORDER-LEFT: #666666 0px solid; BORDER-BOTTOM: #666666 1px solid; FONT-FAMILY: "宋体"; BACKGROUND-COLOR: #bbbbbb;padding-top: 4px;
}
-->
</style>
</head>

<body bgcolor="#fcfcfc">

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" id="tabcss">
  <tr bgcolor="#BFDFFF">
    <form name="form1" method="post" action="pijiaadmin.asp">
      <td height=35 colspan=7> <font color=#0000ff><strong></strong></font>查找： 
        <select name="selectkind">
          <option value="productname">商品名称</option> 
          <option value="yourname">姓名</option>
          <option value="tel">联系方式</option>
		  <option value="memo">评论内容</option>
		  <option value="rememo">回复内容</option>
        </select> <input name="keyword" type="text" size="15" maxlength="15"> 
        <input type="submit" name="Submit" value="查 找"></td>
    </form>
  </tr>
  <tr bgcolor="#E1E1E1"> 
    <td height="24"> <div align="center">图片</div></td>
    <td> <div align="center">商品名称</div></td>
    <td> <div align="center">姓名</div></td>
    <td> <div align="center">联系方式</div></td> 
    <td > <div align="center">评论内容</div></td>
    <td > <div align="center">回复</div></td>
	<td >&nbsp;</td>
  </tr>
  <%
i=0
nowpage=request("page")
selectkind=trim(request("selectkind"))
keyword=trim(request("keyword"))

if keyword="" then
	sql="select * from view_pijia where id>0 order by id desc"
else
	sql="select * from view_pijia where id>0 and "&selectkind&" like '%"&keyword&"%' order by id desc"
end if


set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof or rs.eof then
	response.write "<tr><td colspan=7><div align=center><br>没有记录!</div></td></tr></table>"
else
rs.pagesize=20
rs.absolutepage=1
if request("page")<>"" then rs.absolutepage=request("page")
rowcount=rs.pagesize
	do while not rs.eof and rowcount>0
		id=rs("id")
		productid=rs("productid")
		smallpicpath=rs("smallpicpath")
		productname=rs("productname")
		
		yourname=addspace(rs("yourname"))
		tel=addspace(rs("tel"))
		memo=addspace(rs("memo")) 
		rememo=addspace(rs("rememo"))
		addtime=rs("addtime") 
		check=rs("check") 

%>
  <tr onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
    <td height="23" align="center"><a href="../show.asp?pkid=<%=productid%>" target="_blank"><IMG src="../product/<%=smallpicpath%>" border=0 height='50'></a></td>
    <td><%=productname%></td>
    <td><%=yourname%></td>
    <td style="WIDTH: 120; WORD-WRAP: break-word"><%=tel%></td>
    <td style="WIDTH: 220; WORD-WRAP: break-word"><%=memo%> [<%=addtime%>]</td>
    <td style="WIDTH: 200; WORD-WRAP: break-word"><%=rememo%> <input type="button" name="Button" value="回复" class="input3" onClick="window.open('pijiare.asp?id=<%=id%>','','width=550,height=500,top=120,left=250')"></td>
    <td align=center>
	<%
	  if check="0" then
		response.Write"　　　未审核 <a href=pinjia_check.asp?all=1&check=1&id="& id &" onClick=""return confirm('真的审核通过吗？')""><font color=#ff0000>点击审核</font></a>"
	  else
		response.Write"　　　已审核 <a href=pinjia_check.asp?all=1&check=0&id="& id &" onClick=""return confirm('真的要反审核吗？')""><font color=#ff0000>点击反审</font></a>"
	  end if
	%>
	<br>
	<a href="pijiaadmin.asp?action=del&id=<%=id%>" onClick="return confirm('确认删除吗?')">删除</a></td>
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


