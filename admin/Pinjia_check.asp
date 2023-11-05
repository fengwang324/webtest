<!--#include file="conn.asp"-->
<!--#include file="check4.asp"-->
<%
if session("m_user")<>"" then
	id=request.QueryString("id")
	check=request.QueryString("check")
	conn.execute("update z_pijia set check='"&check&"' where id="&id)
	response.write "<script language=JavaScript>" & "alert('²Ù×÷³É¹¦£¡');" &"window.location.href='pijiaadmin.asp';" & "</script>"
end if

conn.close
set conn=nothing

%>
