<!-- #include file="conn.asp" -->
<!--#include file="check2.asp"-->
<%
good=request("good")
pkid=request("pkid")
kind=request("kind")
keyword=request("keyword")
nowpage=request("page")

sql="update e_product set good='"&good&"',goodtime='"&now()&"' where pkid="&pkid
conn.execute(sql)

response.write "<table align=center width=500 height=50><tr><td>"
'response.write "设置成功。 <a href='"&&"'>返回商品列表</a>"
response.write "<script language=JavaScript>" & "window.location.href='"&request.cookies("oldpage")&"';" & "</script>"
response.write "</td></tr></table>"

%>
