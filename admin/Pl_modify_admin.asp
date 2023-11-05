<!-- #include file="conn2.asp" -->

<%

for i=1 to 50

	pkid=request.form("pkid"&i)
	kind=request.form("kind"&i)
	model=trim(request.form("model"&i))
	productname=trim(request.form("productname"&i))
	
	pipai=trim(request.form("pipai"&i))
	color=trim(request.form("color"&i))
	psize=trim(request.form("psize"&i))
	
	price1=trim(request.form("price1_"&i))
	price2=trim(request.form("price2_"&i))
	price3=trim(request.form("price3_"&i))
	price4=trim(request.form("price4_"&i))
	price5=trim(request.form("price5_"&i))
	price6=trim(request.form("price6_"&i))
	oneweight=trim(request.form("oneweight"&i))
	kuchushu=trim(request.form("kuchushu"&i))
	hit=trim(request.form("hit"&i))
	saleshu=trim(request.form("saleshu"&i))
	if hit="" or not isnumeric(hit) then hit=0
	if saleshu="" or not isnumeric(saleshu) then saleshu=0
	
	
	if oneweight="" or not isnumeric(oneweight) then
	oneweight=0
	end if

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

	if pkid<>"" then
		sql="select * from e_product where pkid="&pkid&""
		set rs=server.CreateObject("adodb.recordset")
		rs.open sql,conn,1,3
		if rs.bof or rs.eof then
		else
			rs("kind")=kind
			rs("model")=model
			rs("productname")=productname
			
			rs("pipai")=pipai
			rs("color")=color
			rs("psize")=psize
			
			rs("price1")=price1
			rs("price2")=price2
			rs("price3")=price3
			rs("price4")=price4
			rs("price5")=price5
			rs("price6")=price6

			rs("oneweight")=oneweight
			rs("kuchushu")=kuchushu
			rs("hit")=hit
			rs("saleshu")=saleshu			
	
			rs.update
		end if
		rs.close
		set rs=nothing
	end if
	
next



'response.redirect "pl_modify.asp?page="&request("page")
response.write "<script language=JavaScript>" & "alert('成功保存此页的修改！');" & "window.location.href='pl_modify.asp?page="&request("page")&"&selectkind="&request("selectkind")&"&keyword="&request("keyword")&"&no=';" & "</script>"

%>










