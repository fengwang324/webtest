<!-- #include file="conn.asp" -->
<!--#include file="check2.asp"-->
<%

pkid=trim(request.form("pkid"))
kind=trim(request.form("kind"))
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
disc=trim(request.form("disc"))
pipai=trim(request.form("pipai"))


fahoudi=trim(request.form("fahoudi"))
kuchumemo=trim(request.form("kuchumemo"))
oneweight=trim(request.form("oneweight"))
kuchushu=trim(request.form("kuchushu"))
if oneweight="" then oneweight="0"
addtime=trim(request.form("addtime"))
if not isdate(addtime) then addtime=date()


sql="select  * from e_product where pkid="&pkid

set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
if rs.bof or rs.eof then
		response.write "不能修改，此商品已不存在！"
else
	rs("kind")=kind
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
	
	rs.update

'	'on error resume next
'	smallpicpathold=request.form("smallpicpathold")
'	bigpicpathold=request.form("bigpicpathold")
'
'	set fso= server.CreateObject("scripting.fileSystemObject")
'	if smallpicpathold<>smallpicpath and smallpicpathold<>"" and smallpicpathold<>bigpicpath then
'		set f1=fso.getfile(server.mappath(""&smallpicpathold&""))
'		f1.delete
'	end if
'	if bigpicpathold<>bigpicpath and bigpicpathold<>"" and bigpicpathold<>smallpicpath then
'		set f2=fso.getfile(server.mappath(""&bigpicpathold&""))
'		f2.delete
'	end if
	
	response.redirect(request.cookies("oldpage"))
	
end if

rs.close
conn.close

%>




