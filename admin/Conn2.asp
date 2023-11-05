<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = 65001
%><!-- #include file="../Dataname.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" c/>
<%
'on error resume next
dim db,conn,connstr
db="../data/"&mydataname
Set conn = Server.CreateObject("ADODB.Connection")
connstr="driver={Microsoft Access Driver (*.mdb)};uid=admin;pwd="&mydatapsw&";dbq=" & Server.MapPath(""&db&"")
conn.Open connstr

sub connclose
	conn.close
	set conn=nothing
end sub

function addspace(strfield)	'列表显示	
	if strfield="" or isnull(strfield) or strfield="未知" then 
		addspace="&nbsp;"
	else
		addspace=strfield
	end if
end function
%>
