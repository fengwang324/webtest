<!-- #include file="conn.asp" -->
<%
dim vipno,truename,province,city,telphone1,telphone2,fax,mobile,address,postno,email
dim paystat
paystat=""

dim sql,rs
if trim(request.form("pkid1"))="" then
	conn.close
	response.write "<script language=JavaScript>" & "alert('提交失败！您的购车是空的。请点击确定后选购商品。');" & "window.location.href='productlist.asp'" & "</script>"
	response.end
end if 

vipno=trim(request.form("vipno")) 
truename=trim(request.form("truename")) 
province=trim(request.form("province")) 
city=trim(request.form("city"))
area=trim(request.form("area"))
telphone1=trim(request.form("telphone1")) 
telphone2=trim(request.form("telphone2")) 
fax=trim(request.form("fax")) 
mobile=trim(request.form("mobile")) 
address=trim(request.form("address")) 
postno=trim(request.form("postno")) 
email=trim(request.form("email")) 
customid=trim(request.form("customid"))
if customid="" then customid="0"

if province=city then
	address=city&" "&area&" "&address
else
	address=province&" "&city&" "&area&" "&address
end if
if address<>"" then
	address=replace(address,"省直辖行政单位","")
	address=replace(address,"省直辖县级行政单位","")
end if

'billdate
billno=request.Cookies("notemp")
memo=trim(request.form("memo"))

sendtype=trim(request.form("sendtype"))
sendmoney=trim(request.form("sendmoney"))
allmoney=trim(request.form("allmoney"))
youhjmoney=trim(request.form("youhjmoney"))
total=trim(request.form("total"))
paytype=trim(request.form("paytype"))
'sendornot

'-------------------------打积分比例BEGIN------------------------------
sql4="select * from zhfbili"
set rs4=server.createobject("adodb.recordset")
rs4.open sql4,conn,1,1
if rs4.bof or rs4.eof then
		zhfbili=0
else
		zhfbili=rs4("zhfbili")
end if
rs4.close
set rs4=nothing
'--------------------------打积分比例END-------------------------------

'----------------通过后台确认账户余额是否足够BEGIN----------------------
if paytype="账户余额" then
		sql6="SELECT id, username,(select sum(czmoney) from sh_chongzhi where sh_chongzhi.customid=a.id and confirmflag='是') AS sumcz, "&_
		"(select sum(total) from x_bill where x_bill.customid=a.id and paytype='账户余额') AS sumkj,  "&_
		"(IIf(IsNull(sumcz),0,sumcz)-IIf(IsNull(sumkj),0,sumkj)) AS leftye  "&_
		"FROM x_huiyuan AS a where username='"&session("username")&"' "
		
		set rs6=server.createobject("adodb.recordset")
		rs6.open sql6,conn,1,1
		if rs6.bof or rs6.eof then
			leftye=0
		else
			leftye=rs6("leftye")
		end if
		rs6.close
		set rs6=nothing
		if cdbl(leftye)<cdbl(total) then
			conn.close
			set conn=nothing
			response.write "<br>您的账户余额不足够支付本次交易。<a href=javascript:history.back()>返回</a>"
			response.end
		end if
		paystat="已付款"
end if
'----------------通过后台确认账户余额是否足够  END----------------------

sql="select * from x_bill where billno='"&billno&"'"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
if rs.bof or rs.eof then
	rs.addnew

	rs("vipno")=vipno 
	rs("truename")=truename 
	rs("province")=province 
	rs("city")=city 
	rs("telphone1")=telphone1 
	rs("telphone2")=telphone2 
	rs("fax")=fax 
	rs("mobile")=mobile 
	rs("address")=address 
	rs("postno")=postno 
	rs("email")=email 
	rs("customid")=customid
	'billdate
	rs("billno")=billno
	rs("memo")=memo
	
	rs("sendtype")=sendtype
	rs("whopaysendmoney")="买家支付"
	
	rs("allmoney")=allmoney    	 '商品金额
	rs("sendmoney")=sendmoney  	 '运输费
	
	total0=cdbl(allmoney) + cdbl(sendmoney)
	rs("total")=total0   		 '商品金额+运输费
	
	rs("youhjmoney")=0  '优惠券抵扣金额
	'youhuimoney				 '卖家优惠的金额：负数为优惠 正数为附加金额。像淘宝里修改价格。卖家在后台修改的。
	rs("lastneedpay")=total      '最后所需付金额
	
	
	rs("paytype")=paytype
	if session("m_gold")<>"" and session("s_stat")="" then
	rs("backop")=session("m_user")
	end if
	rs("paystat")=paystat
	'sendornot
	rs("billzhf")=allmoney*zhfbili
	
	rs.update
	
	'===============以上保存基本信息===================
	'取得这个单的id
	cussql="select id from x_bill where billno='"&billno&"'"
	set cusrs=conn.execute(cussql)
	nowid=cusrs("id")
	cusrs.close
	set cusrs=nothing
	'===============以下保存商品信息===================
	for i =1 to request.form("maxi")
		pkid=request.form("pkid"&i)
		colorsize=request.form("colorsize"&i)
		price=request.form("price"&i)
		num=request.form("num"&i)
		money=request.form("money"&i)
		oneweight=request.form("oneweight"&i)
		allweight=request.form("allweight"&i)
		
		sql0="insert into x_bill_product(billid,productpkid,colorsize,num,price,[money],oneweight,allweight) values('"&nowid&"','"&pkid&"','"&colorsize&"','"&num&"','"&price&"','"&money&"','"&oneweight&"','"&allweight&"')"
		conn.execute(sql0)
	next
	
	nowtime=now()
	maxyy=request.form("maxyy")
	if maxyy="" then maxyy=0
	for j=1 to maxyy
		dkbox=request.form("dkbox"&j)
		if dkbox<>"" then
			idbox=request.form("idbox"&j)
			sql0="update x_bill set youhj_used='是',youhj_useddate='"&nowtime&"',youhj_used_billid='"&nowid&"' where id="&idbox
			conn.execute(sql0)
		end if
	next

	sql8="delete * from x_order where notemp='"&request.Cookies("notemp")&"'"
	conn.execute(sql8)
	response.Cookies("notemp")=""
	response.Cookies("allmoney")=""
	
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	

	response.redirect "orderjs3.asp?billno="&billno
	response.end
		
else
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "<script language=JavaScript>" & "alert('订单保存失败。请点击确定重新提交订单。如果重试都不能保存，请空购物车，重新选择商品。');" & "window.history.back()" & "</script>"
	response.end
end if


	
%>


