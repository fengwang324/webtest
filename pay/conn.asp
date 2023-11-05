<!-- #include file="../Dataname.asp" -->
<%
Response.Buffer = True 
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache"
db="../data/"&mydataname
Set conn = Server.CreateObject("ADODB.Connection")
connstr="driver={Microsoft Access Driver (*.mdb)};uid=admin;pwd="&mydatapsw&";dbq=" & Server.MapPath(""&db&"")
conn.Open connstr


%>
