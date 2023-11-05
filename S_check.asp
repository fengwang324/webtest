<!--#include file="conn.asp"-->
<%
if session("m_user")<>"" then
	id=request.QueryString("id")
	check=request.QueryString("check")
	set rs=server.createobject("adodb.recordset")
	rs.open "select * from message where id="&id,conn,1,3
	rs("check")=check
	rs.update
	rs.close
	set rs=nothing
	response.write "<script language=JavaScript>" & "alert('²Ù×÷³É¹¦£¡');" &"window.location.href='showmessage.asp?id="&id&"';" & "</script>"
end if

conn.close
set conn=nothing

%>
