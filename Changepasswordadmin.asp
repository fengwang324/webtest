<%@ LANGUAGE = VBScript.Encode %>
<!-- #include file="conn.asp" -->
<%
dim oldpassword,newpassword,sql,rs

sub sub_safe(letter)
	letter=lcase(letter)
	if InStr(letter,"'")>0 or InStr(letter," ")>0 or InStr(letter,"--")>0 or InStr(letter,";;")>0 then
		response.write "<br>&nbsp;&nbsp;<font size=2><font color=#ff0000>ϵͳ��ʾ��</font>��������ַ����ַ���пո��������ţ�������������롣"
		response.write "<a href=javascript:history.back()>����</a> &nbsp;&nbsp;&nbsp;<a href=help.asp>�鿴����</a> </font>"
		response.end
	else
		if InStr(letter,"&")>0 or InStr(letter,"select")>0 or InStr(letter,"update")>0 or InStr(letter,"delete")>0 or InStr(letter,"asc(")>0 then
			response.write "<br>&nbsp;&nbsp;<font size=2><font color=#ff0000>ϵͳ��ʾ��</font>��������ַ����ַ���пո��������ţ�������������롣"
			response.write "<a href=javascript:history.back()>����</a> &nbsp;&nbsp;&nbsp;<a href=help.asp>�鿴����</a> </font>"
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
			response.write "<br><div align=center>���ľ����벻��!<a href=javascript:window.history.back()>����</a></div>"
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
		response.write "<title>�޸�����</title><br><br><div align=center>�����޸ĳɹ���<a href='index.asp'>������ҳ</a></div>"
%>

















