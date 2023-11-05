<!--#include file="conn.asp"-->
<!--#include file="check2.asp"-->
<%
set rs=server.CreateObject("adodb.recordset")
sql="select * from e_contect where c_id ="&request("c_id")
rs.open sql,conn,1,3
if rs.bof or rs.eof then
	response.write "此栏目不存在。"
	call history
else
	rs.delete
	response.redirect "conlist.asp?l_id="&request.QueryString("newskind")
end if

rs.close
set rs=nothing

call connclose

%>
