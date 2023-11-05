<!--#include file="conn.asp"-->
<!--#include file="check4.asp"-->
<%
set rs=server.CreateObject("adodb.recordset")
sql="select * from x_huiyuan where id = "&request.QueryString("id")
rs.open sql,conn,1,3
if rs.bof or rs.eof then
	response.write "此会员不存在，可能被其他管理人员刚刚删除。请后退，刷新，重试。"
else
	rs.delete
	response.redirect (request.cookies("hypage"))
end if

rs.close
set rs=nothing

call connclose

%>
