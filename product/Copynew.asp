<!-- #include file="conn.asp" -->
<!--#include file="check2.asp"-->
<%

dim i,strid,sql,rs
dim sql2,rs2
dim newsid1,newsid2,news1,news2

for i=1 to 30
	strid=request.form("checkbox"&i)
	if strid<>"" then

		sql="insert into e_product(kind,kind2,model,productname,price1,price2,price3,price4,price5,price6,smallpicpath,bigpicpath,bigpicpath2,bigpicpath3,bigpicpath4,bigpicpath5,disc,pipai,hot,hottime,good,goodtime,hit,color,psize,updown,updowntime,oneweight,kuchushu) select kind,kind2,model,productname,price1,price2,price3,price4,price5,price6,smallpicpath,bigpicpath,bigpicpath2,bigpicpath3,bigpicpath4,bigpicpath5,disc,pipai,hot,hottime,good,goodtime,hit,color,psize,updown,updowntime,oneweight,kuchushu from e_product where pkid="&strid&""
		'response.write sql
		'response.end
		conn.execute(sql)
		
	end if
next


response.redirect "productlist.asp?menuid="&request("menuid")


%>










