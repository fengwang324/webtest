<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%
m_gold=session("m_gold")
if m_gold="2" or m_gold="8" then
else
	response.write "对不起，您没有此权限。"
	response.end
end if
%>
