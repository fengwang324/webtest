<!--#include file="conn.asp"-->
<%
dim sql 
dim rs
sql="delete from khly where id="&request("id")&" "
set rs=conn.execute(sql)
conn.close
set conn=nothing
'response.write"<script>alert('信息删除成功!');navigate('khly.asp');<'/script>"
response.redirect "khly.asp?l_id="&request.QueryString("newskind")
%>