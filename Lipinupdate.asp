<!-- #include file="conn.asp" -->
<%
	dim allzhifen,lipindate,username
	lipindate=now()
	allzhifen=0
	username=session("username")

if session("username")="" or session("s_stat")="" then
else
	'-----------------计算剩余积分BEGIN----------------
	sql6="SELECT id, username,(select sum(billzhf) from x_view_bill where x_view_bill.username=a.username and sendornot='已发货') AS sumzhf, "&_
	"(select sum(kjzhf) from x_huanlipin where x_huanlipin.username=a.username) AS sumkjzhf,  "&_
	"(IIf(IsNull(sumzhf),0,sumzhf)-IIf(IsNull(sumkjzhf),0,sumkjzhf)) AS leftzhf  "&_
	"FROM x_huiyuan AS a where username='"&username&"' "
	
	set rs6=server.createobject("adodb.recordset")
	rs6.open sql6,conn,1,1
	if rs6.bof or rs6.eof then
		leftzhf=0
	else
		leftzhf=rs6("leftzhf")
	end if
	rs6.close
	set rs6=nothing
	'-----------------计算剩余积分 END ----------------

	for i =1 to request.form("maxi")
	
		id=request.form("id"&i)
		shu=request.form("shu"&i)
		
		if shu<>"" and isnumeric(shu) then
			sql="select lipinname,zhifen from x_lipin where id="&id
			set rs=server.CreateObject("adodb.recordset")
			rs.open sql,conn,1,1
			if rs.bof or rs.eof then
			else
				lipinname=rs("lipinname")
				zhifen=rs("zhifen")
				allzhifen=allzhifen+zhifen*shu
				sql0="insert into x_huanlipin(lipindate,username,lipinname,zhifen,shu,kjzhf) values('"&lipindate&"','"&username&"','"&lipinname&"','"&zhifen&"','"&shu&"','"&zhifen*shu&"')"
				conn.execute(sql0)
			end if
			rs.close
			set rs=nothing
		end if
		
	next
	
	
	if allzhifen>leftzhf then  '后台重新确认是否够积分
		sql0="delete * from x_huanlipin where lipindate=#"&lipindate&"# and username='"&username&"'"
		conn.execute(sql0)
		response.write "<script language=JavaScript>" & "alert('保存失败。请重新进行兑换。');" & "window.location.href='lipinlist.asp'" & "</script>"
	else
		response.write "<script language=JavaScript>" & "alert('礼品兑换成功。请尽快购买商品，并在该订单附言中说明“有礼品一同发送”字样。');" & "window.location.href='lipinlist.asp'" & "</script>"
	end if


end if

conn.close
set conn=nothing
	
%>


