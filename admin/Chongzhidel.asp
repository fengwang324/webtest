<!--#include file="conn.asp"-->
<!--#include file="check4.asp"-->
<%
dim menuid
dim i,strid,sql
menuid=request.form("menuid")

if session("righturl")="1" then

	strid=request("strid")
	if strid<>"" then
	sql="delete from sh_chongzhi where id="&strid
	conn.execute(sql)
	end if
end if

response.redirect ("chongzhi.asp?menuid="&menuid&"&sublist="&request.Cookies("sublist"))

%>


