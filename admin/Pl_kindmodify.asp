<!-- #include file="conn.asp" -->
<!--#include file="check2.asp"-->
<%

dim i,strid,sql,rs
dim sql2,rs2
dim newsid1,newsid2,news1,news2
allkind=request.form("allkind")

for i=1 to 30
	strid=request.form("checkbox"&i)
	if strid<>"" then
		sql=" update e_product set kind='"&allkind&"' where pkid="&strid&""
		conn.execute(sql)
	end if
next


response.write "<script language=JavaScript>" & "alert('成功修改所选商品分类！');" & "window.location.href='pl_modify.asp?page="&request("page")&"&selectkind="&request("selectkind")&"&keyword="&request("keyword")&"&no=';" & "</script>"


%>










