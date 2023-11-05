
<!-- #include file="conn.asp" -->
<%
id=trim(request.form("id"))
picpath=trim(request.form("smallpicpath"))
picpath2=trim(request.form("bigpicpath"))
lipinno=trim(request.form("lipinno"))
lipinname=trim(request.form("lipinname"))
zhifen=trim(request.form("zhifen"))
showflag=trim(request.form("showflag"))
num=trim(request.form("num"))
memo=trim(request.form("memo"))

if id="" then
	sql="select  * from x_lipin "
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
		rs.addnew
		rs("picpath")=picpath
		rs("picpath2")=picpath2
		rs("lipinno")=lipinno
		rs("lipinname")=lipinname
		rs("zhifen")=zhifen
		rs("showflag")=showflag
		rs("num")=num
		rs("memo")=memo
		rs.update
	rs.close
	conn.close
else
	sql="select  * from x_lipin where id="&id&" "
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
		rs("picpath")=picpath
		rs("picpath2")=picpath2
		rs("lipinno")=lipinno
		rs("lipinname")=lipinname
		rs("zhifen")=zhifen
		rs("showflag")=showflag
		rs("num")=num
		rs("memo")=memo
		rs.update
	rs.close
	conn.close
end if

response.redirect("lipin.asp")
%>
