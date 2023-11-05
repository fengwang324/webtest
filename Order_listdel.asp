<!--#include file="conn.asp"-->
<%
if session("username")="" then
	conn.close
	set conn=nothing
	response.Redirect("hyzq.asp")
	response.end
end if
%>
<%
id = request.QueryString("id")
if id<>"" and isnumeric(id) then
	sql4="select id from x_huiyuan where username='"&session("username")&"'"
	set rs4=conn.execute(sql4)
	customid=rs4("id")
	rs4.close
	set rs4=nothing


	sql="select * from x_bill where customid="&customid&" and id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	if rs.bof or rs.eof then
	else
		sendornot=rs("sendornot")
		if sendornot="未处理" then
		rs("sendornot")="会员自行取消"
		rs.update
		end if
		
	end if
	rs.close
	
end if
conn.close
response.Redirect("order_list.asp")
%>
