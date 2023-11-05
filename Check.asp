<!--#include file="conn.asp"-->
<!--#include file="md5.asp"-->
<%
Response.CacheControl = "no-cache"

if session("sure")<>"1" then 
	response.write "请“后退”->“刷新”，重新输入用户名及密码。"
	conn.close
	set conn=nothing
	response.end
end if
'session.Abandon()

sub passwrong()
	response.write "<script language=JavaScript>" & "alert('对不起，登录失败。您所填写的“用户名”与密码不符，或此用户名被停用。\n请点击确定，重新填写用户名及密码。');" & "window.history.back()" & "</script>"
	response.end
end sub
sub connend()
	conn.close
	set conn=nothing
end sub

username=trim(request.form("username"))
password=trim(request.form("password"))



if password="" or username="" then
	conn.close
	set conn=nothing
	call passwrong
end if

if InStr(username,"'")>0 or InStr(username,"--")>0 or InStr(username,"(")>0 or InStr(username,";")>0 or InStr(username,"replace")>0 or InStr(username,"eval")>0 then
	conn.close
	set conn=nothing
	call passwrong
END IF
if InStr(password,"'")>0 or InStr(password,"--")>0 or InStr(password,"(")>0 or InStr(password,";")>0 or InStr(password,"replace")>0 or InStr(password,"eval")>0 then
	conn.close
	set conn=nothing
	call passwrong
END IF

set rs=server.createobject("adodb.recordset")
	sql="select username,password,lasttime,lastip,s_stat,customkind from x_huiyuan where s_stat<>-1 and username='" & username & "'"
	rs.open sql,conn,1,3
 	if not(rs.bof and rs.eof) then
 		if md5(password)=md5(rs("password")) then
		    session("username")=rs("username")
			session("s_stat")=rs("s_stat")
			customkind=rs("customkind")
			session("customkind")=customkind
			
			if customkind="2" then 
			session("customkindname")="普通会员"
			elseif customkind="3" then 
			session("customkindname")="铜级会员"
			elseif customkind="4" then 
			session("customkindname")="银级会员"
			elseif customkind="5" then 
			session("customkindname")="金级会员"
			elseif customkind="6" then 
			session("customkindname")="钻级会员"
			else
			session("customkindname")=""
			end if
			
			
			rs("lasttime")=now()
			rs("lastip")=request.ServerVariables("REMOTE_ADDR")
            rs.update
			rs.close
			set rs=nothing
			call connend()
			if request.QueryString("jieshun")="1" then
				response.write "<script language=JavaScript>" & "alert('成功登录！您将可以享受“"&session("customkindname")&"”的优惠。');" & "window.location.href='orderjs.asp'" & "</script>"
			elseif request.QueryString("jieshun")="2" then
				response.write "<script language=JavaScript>" & "alert('成功登录！您将可以享受“"&session("customkindname")&"”的优惠。');" & "window.location.href='lipinlist.asp'" & "</script>"
			elseif InStr(request.serverVariables("Http_REFERER"),"hyzq.asp")>0 then
				response.write "<script language=JavaScript>" & "alert('成功登录！您将可以享受“"&session("customkindname")&"”的优惠。');" & "window.location.href='hyzq.asp'" & "</script>"
			else
				response.write "<script language=JavaScript>" & "alert('成功登录！您将可以享受“"&session("customkindname")&"”的优惠。');" & "window.location.href='index.asp'" & "</script>"
 			end if		
		else
			rs.close
			set rs=nothing
			call connend()
			call passwrong
 		end if
	else
		rs.close
		set rs=nothing
		call connend()
		call passwrong
	end if
%>




