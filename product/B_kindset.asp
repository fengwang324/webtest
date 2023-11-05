<!-- #include file="conn.asp" -->
<!--#include file="check2.asp"-->
<%
kindnum=request("kindnum")
v=request("v")

sql="update sh_kind set indexshow='"&v&"' where kindnum like '"&kindnum&"%' "
conn.execute(sql)

response.Redirect("b_kind.asp")

%>
