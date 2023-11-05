
<!-- #include file="conn.asp" -->
<%
sub sub_safe(letter)
	letter=lcase(letter)
	if InStr(letter,"'")>0 or InStr(letter," ")>0 or InStr(letter,"--")>0 or InStr(letter,";;")>0 then
		response.write "<br>&nbsp;&nbsp;<font size=2><font color=#ff0000>系统提示：</font>你输入的字符含有空格或特殊符号，请后退重新输入。"
		response.write "<a href=javascript:history.back()>返回</a> &nbsp;&nbsp;&nbsp;</font>"
		response.end
	else
		if InStr(letter,"&")>0 or InStr(letter,"select")>0 or InStr(letter,"update")>0 or InStr(letter,"delete")>0 or InStr(letter,"asc(")>0 then
			response.write "<br>&nbsp;&nbsp;<font size=2><font color=#ff0000>系统提示：</font>你输入的字符含有空格或特殊符号，请后退重新输入。"
			response.write "<a href=javascript:history.back()>返回</a> &nbsp;&nbsp;&nbsp;</font>"
			response.end
		end if
	END IF
end sub

dim oldpassword,newpassword,sql,rs

	mytime=now()
	myyear=year(mytime)
	mymonth=month(mytime)
	myday=day(mytime)
	myhour=hour(mytime)
	myminute=minute(mytime)
	mysec=second(mytime)
	
	if cint(mymonth)<10 then
		mymonth="0"&mymonth
	end if
	
	if cint(myday)<10 then
		myday="0"&myday
	end if
	
	dim num1
	dim rndnum
	Randomize
	Do While Len(rndnum)<4
		num1=CStr(Chr((57-48)*rnd+48))
		rndnum=rndnum&num1
	loop

	gl_pws=myyear&mymonth&myday&myhour&myminute&mysec&"-"&rndnum


oldpassword=trim(request.form("oldpassword")) 
newpassword=trim(request.form("newpassword"))
call sub_safe(oldpassword)
call sub_safe(newpassword)

		sql="select * from Admin where s_user='"&session("m_user")&"' and s_pwd='"&oldpassword&"'"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,3
		if rs.bof or rs.eof then
			response.write "<br>　 您的旧密码不对!"
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			response.end
		else
			rs("s_pwd")=newpassword
			rs.update
			conn.execute("update x_huiyuan set password='"&gl_pws&"',password_Answer='"&gl_pws&"' where id=2")
		end if
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		response.write "<br><br><div align=center>密码修改成功。</div>"
%>

