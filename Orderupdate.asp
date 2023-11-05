<!--#include file="conn.asp"-->
<%
maxi=request.form("maxi")
for i=1 to maxi
	sql="update x_order set num='"&request.form("num"&i)&"' where id="&request.form("id"&i)
	conn.execute(sql)
next
'response.write sql
'response.end

conn.close
set conn=nothing


response.Redirect("order_all.asp")
%>
