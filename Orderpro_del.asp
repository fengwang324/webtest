<!--#include file="conn.asp"-->
<%
if request.QueryString("kind")="qingkong" then
sql="delete * from x_order where notemp='"&request.Cookies("notemp")&"'"
response.Cookies("notemp")=""
else
sql="delete * from x_order where id="&request.QueryString("id")
end if
conn.execute(sql)


conn.close
set conn=nothing

response.Redirect("order_all.asp")
%>
