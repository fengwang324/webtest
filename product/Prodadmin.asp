<!-- #include file="conn.asp" -->
<!--#include file="check2.asp"-->
<%

kind=trim(request.form("kind"))
kind2=trim(request.form("kind2"))
model=trim(request.form("model"))
productname=trim(request.form("productname"))
price1=trim(request.form("price1"))
price2=trim(request.form("price2"))
price3=trim(request.form("price3"))
price4=trim(request.form("price4"))
price5=trim(request.form("price5"))
price6=trim(request.form("price6"))
smallpicpath=trim(request.form("smallpicpath"))
bigpicpath=trim(request.form("bigpicpath"))
bigpicpath2=trim(request.form("bigpicpath2"))
bigpicpath3=trim(request.form("bigpicpath3"))
bigpicpath4=trim(request.form("bigpicpath4"))
bigpicpath5=trim(request.form("bigpicpath5"))
disc=trim(request.form("disc"))
pipai=trim(request.form("pipai"))
color=trim(request.form("color"))
psize=trim(request.form("psize"))


fahoudi=trim(request.form("fahoudi"))
kuchumemo=trim(request.form("kuchumemo"))
oneweight=trim(request.form("oneweight"))
kuchushu=trim(request.form("kuchushu"))
if oneweight="" then oneweight="0"
addtime=trim(request.form("addtime"))
if not isdate(addtime) then addtime=date()

if color<>"" then
	color=color&"，"
	color=replace(color,",","，")
	color=replace(color,"，，","，")
	color=replace(color,"，，","，")
	color=replace(color,"，，","，")
end if
if psize<>"" then
	psize=psize&"，"
	psize=replace(psize,",","，")
	psize=replace(psize,"，，","，")
	psize=replace(psize,"，，","，")
	psize=replace(psize,"，，","，")
end if

 
	
	sql="select  * from e_product where model='"&model&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	if rs.bof or rs.eof then
		session("submit")=""  '防重复提交
		rs.addnew
		
		rs("kind")=kind
		rs("kind2")=kind2
		rs("model")=model
		rs("productname")=productname
		rs("price1")=price1
		rs("price2")=price2
		rs("price3")=price3
		rs("price4")=price4
		rs("price5")=price5
		rs("price6")=price6
		rs("smallpicpath")=smallpicpath
		rs("bigpicpath")=bigpicpath
	
		rs("disc")=disc
		rs("pipai")=pipai
	

		'rs("fahoudi")=fahoudi
		'rs("kuchumemo")=kuchumemo
		rs("oneweight")=oneweight
		rs("kuchushu")=kuchushu
		rs("addtime")=addtime
		rs("punit")=trim(request.form("punit"))
		'rs("selt1")=trim(request.form("selt1"))
		rs("updowntime") = now()
		
		rs.update
		
		response.redirect("productlist.asp")
	else
		response.write "此商品代码已存在。 <a href='javascript:window.history.back()'>返回</a>"
	end if
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 

%>
