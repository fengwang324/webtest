<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" c/>
<%
'on error resume next
Response.CharSet = "utf-8"
Session.CodePage = 65001
%>
<!-- #include file="Dataname.asp" -->
<%
dim db,connstr
db="data/"&mydataname
Set conn = Server.CreateObject("ADODB.Connection")
'connstr="driver={Microsoft Access Driver (*.mdb)};uid=admin;pwd="&mydatapsw&";dbq=" & Server.MapPath(""&db&"")
ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)

conn.Open connstr

if session("customkind")="" or session("customkindname")="" then
	session("customkind")="2"
	session("customkindname")="普通会员"
end if

'---------------选定显示格式begin------------------
dim s_sql,s_rs,s_col,s_row,s_picwidth,s_picheight,s_colorsize,s_news,s_mes,s_hit,s_productkind,s_head,s_foot,s_diaocha,s_youqing,s_zishu,s_qq,s_msn,s_wangwang,s_pinpai,s_pinpai_l_r,s_pinpaiw,s_pinpaih
dim s_hot_line,s_cuxiao_line,s_new_line
dim s_model,s_pipai,s_kuchu
s_sql="select * from siteshow"
set s_rs=server.createobject("adodb.recordset")
s_rs.open s_sql,conn,1,1
if s_rs.bof or s_rs.eof then
	
else
	s_col=s_rs("s_col")
	s_row=s_rs("s_row")
	s_picwidth=s_rs("s_picwidth")
	s_picheight=s_rs("s_picheight")
	s_colorsize=s_rs("s_colorsize")

	s_model=s_rs("s_model")
	s_pipai=s_rs("s_pipai")
	s_kuchu=s_rs("s_kuchu")
	
	s_hot_line=s_rs("s_hot_line")
	s_cuxiao_line=s_rs("s_cuxiao_line")
	s_new_line=s_rs("s_new_line")
    s_content=s_rs("s_content")
	s_news=s_rs("s_news")
	s_mes=s_rs("s_mes")
	s_hit=s_rs("s_hit")
	s_hitnum=s_rs("s_hitnum")
	s_productkind=s_rs("s_productkind")
	s_head=s_rs("s_head")
	s_foot=s_rs("s_foot")
	s_diaocha=s_rs("s_diaocha")
	s_bbs=s_rs("s_bbs")
	s_youqing=s_rs("s_youqing")
	s_zishu=s_rs("s_zishu")
	s_qq=s_rs("s_qq")
	s_msn=s_rs("s_msn")
	s_wangwang=s_rs("s_wangwang")
	s_pinpai=s_rs("s_pinpai")
	s_pinpai_l_r=s_rs("s_pinpai_l_r")
	s_pinpaiw=s_rs("s_pinpaiw")
	s_pinpaih=s_rs("s_pinpaih")
end if
s_rs.close
set s_rs=nothing
'--------------选定显示格式end---------------------
%>


