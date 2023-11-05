<!--#include file="conn.asp"-->
<!--#include file="check4.asp"-->
<%
id = request.QueryString("id")
if id<>"" and isnumeric(id) then

	sql ="delete * from x_bill where id="&id
	conn.execute(sql)
	sql2="delete * from x_bill_product where billid="&id
	conn.execute(sql2)

end if
conn.close
response.redirect (request.cookies("ddpage"))
%>
