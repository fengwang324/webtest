<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'on error resume next
Response.CharSet = "utf-8"
Session.CodePage = 65001
%><!-- #include file="../dataname.asp" -->
<!-- #include file="md5.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" c/>
<%
response.write "<style>td{font-size:9pt}</style>"
m_user=session("m_user")
m_gold=session("m_gold")
 

db="../data/"&mydataname
set conn = server.createobject("adodb.connection")
connstr="driver={microsoft access driver (*.mdb)};uid=admin;pwd="&mydatapsw&";dbq=" & server.mappath(""&db&"")
conn.open connstr

sub connclose
	conn.close
	set conn=nothing
end sub
sub history
	response.write " <a href=javascript:history.back()>返回</a> "
end sub

function addspace(strfield)	'列表显示	
	if strfield="" or isnull(strfield) or strfield="未知" then 
		addspace="&nbsp;"
	else
		addspace=strfield
	end if
end function
%>

