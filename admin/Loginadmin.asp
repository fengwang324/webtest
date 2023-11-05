<!--#include file="conn2.asp"-->
<!--#include file="md5.asp"-->
<% 
sub passwrong()
	response.write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
	response.write "<title>登陆失败请返回重新输入</title><br><br>"
	response.write "<table align=center border=1 cellspacing=0 cellpadding=20 bordercolorlight=#cccccc bordercolordark=#ffffff><tr>"
	response.write "<td width=400 align=left height=160 bgcolor=#eeeeee style='font-size:12px;'>"
	response.write "<font style='font-size:14px;font-weight:bold;color:#ff0000;'>系统提示：</font><br><br><br>"
	response.write "<a href='javascript:window.history.back()'><b>登陆后台失败。点击这里重新输入用户名及密码。</b></a>"
	response.write "</td></tr></table>"
	response.end
end sub
if cstr(session("getcode"))<>cstr(trim(request("verifycode"))) then
response.Write "<script LANGUAGE='javascript'>alert('请输入正确的验证码！');history.go(-1);</script>"
response.end
end if
dim User_name,User_pwd,confirmcode
dim del_date,sql0


User_name=trim(request.form("h_name"))
User_pwd=md5(trim(request.form("h_pwd")))

if User_pwd="" or User_name="" then
	conn.close
	set conn=nothing
	call passwrong()
end if
if InStr(User_name,"'")>0 or InStr(User_name,"--")>0 or InStr(User_name,"(")>0 or InStr(User_name,";")>0 or InStr(User_name,"replace")>0 or InStr(User_name,"eval")>0 then
	conn.close
	set conn=nothing
	call passwrong()
END IF


set rs=server.createobject("adodb.recordset")
	sql="select * from Admin where s_user='" & User_name & "'"
	rs.open sql,conn,1,3
	if rs.bof or rs.eof then
		call passwrong()
	else
		if User_pwd=md5(rs("s_pwd")) then
			t=rs("lasttime")
			ip=rs("lastip")
			session("m_user")=rs("s_user")
			session("m_gold")=rs("self1")
			
		    session("username")=rs("s_user")
			'session("s_stat")="8"
			rs("lasttime")=now()
			rs("lastip")=request.ServerVariables("REMOTE_ADDR")
			rs.update
			del_date=date()-60
			sql0="delete * from [x_order] where adddate<#"&del_date&"# "
			'response.write del_sql
			'conn.execute(sql0)
			
			response.cookies("shoplogintime")=t
			response.cookies("shoploginip")=ip
			
			response.Redirect("index.asp")
		else
			call passwrong()
		end if
	end if
	
rs.close
conn.close
%>
