<!--#include file="conn.asp"-->
<!--#include file="check4.asp"-->
<%
id = request.QueryString("id")
if id<>"" and isnumeric(id) then

	sql2="update x_huanlipin set sendstat='已发货' where id="&id
	conn.execute(sql2)

end if
conn.close
response.redirect ("huihuan.asp")
%>
