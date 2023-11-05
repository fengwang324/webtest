<!--#include file="conn.asp"-->
<!--#include file="check2.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>设定栏目</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
<!--
td{font-size:9pt}
body {
	margin-top: 5px;margin-left: 2px;
}
.style3 {color: #A84E22; font-weight: bold; }
-->
</style>
<%
dim sql4,rs4,id,num,pkid

for j=1 to request.form("i")
	l_id = request.form("l_id"&j)
	mylock = request.form("mylock"&j)
	menukind = request.form("menukind"&j)
	l_title = trim(request.form("l_title"&j))
	l_num = trim(request.form("l_num"&j))
	bigmenu = request.form("bigmenu"&j)
	zhidingurl = trim(request.form("zhidingurl"&j))
	showflag = request.form("showflag"&j)
	newopen = request.form("newopen"&j)
	
	if l_num="" or not isnumeric(l_num) then l_num=0
	
	if isnumeric(l_id) and l_title<>"" and mylock="是" then
		sql2="update e_left set l_title='"&l_title&"',l_num='"&l_num&"',bigmenu='"&bigmenu&"',showflag='"&showflag&"',newopen='"&newopen&"' where l_id="&l_id
		conn.execute(sql2)
	elseif isnumeric(l_id) and l_title<>"" then
		sql2="update e_left set menukind='"&menukind&"',l_title='"&l_title&"',l_num='"&l_num&"',bigmenu='"&bigmenu&"',zhidingurl='"&zhidingurl&"',showflag='"&showflag&"',newopen='"&newopen&"' where l_id="&l_id
		conn.execute(sql2)
	elseif l_id="" and l_title<>"" then
		sql2="insert into e_left (menukind,l_title,l_num,bigmenu,zhidingurl,showflag,newopen) values('"&menukind&"','"&l_title&"',"&l_num&",'"&bigmenu&"','"&zhidingurl&"','"&showflag&"','"&newopen&"')"
		conn.execute(sql2)
	elseif isnumeric(l_id) and l_title="" then
		sql2="delete * from e_left where l_id="&l_id
		conn.execute(sql2)
	else
	end if
next

response.write "<script language=JavaScript>" & "alert('成功保存栏目设置！');" & "window.top.location.href='index.asp';"  & "</script>"
%>
</head>

<%
call connclose
%>

