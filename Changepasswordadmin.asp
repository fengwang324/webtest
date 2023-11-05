<%@ LANGUAGE = VBScript.Encode %>
<!-- #include file="conn.asp" -->
<%
dim oldpassword,newpassword,sql,rs

sub sub_safe(letter)
	letter=lcase(letter)
	if InStr(letter,"'")>0 or InStr(letter," ")>0 or InStr(letter,"--")>0 or InStr(letter,";;")>0 then
		response.write "<br>&nbsp;&nbsp;<font size=2><font color=#ff0000>系统提示：</font>你输入的字符或地址含有空格或特殊符号，请后退重新输入。"
		response.write "<a href=javascript:history.back()>返回</a> &nbsp;&nbsp;&nbsp;<a href=help.asp>查看帮助</a> </font>"
		response.end
	else
		if InStr(letter,"&")>0 or InStr(letter,"select")>0 or InStr(letter,"update")>0 or InStr(letter,"delete")>0 or InStr(letter,"asc(")>0 then
			response.write "<br>&nbsp;&nbsp;<font size=2><font color=#ff0000>系统提示：</font>你输入的字符或地址含有空格或特殊符号，请后退重新输入。"
			response.write "<a href=javascript:history.back()>返回</a> &nbsp;&nbsp;&nbsp;<a href=help.asp>查看帮助</a> </font>"
			response.end
		end if
	END IF
end sub


oldpassword=trim(request.form("oldpassword")) 
newpassword=trim(request.form("newpassword"))
call sub_safe(oldpassword)
call sub_safe(newpassword)

		sql="select password from x_huiyuan where username='"&session("username")&"' and password='"&oldpassword&"'"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,3
		if rs.bof or rs.eof then
			response.write "<br><div align=center>您的旧密码不对!<a href=javascript:window.history.back()>返回</a></div>"
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			response.end
		else
			rs("password")=newpassword
			rs.update
		end if
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		response.write "<title>修改密码</title><br><br><div align=center>密码修改成功。<a href='index.asp'>返回主页</a></div>"
%>

















