<!--#include file="conn.asp"-->
<!--#include file="check2.asp"-->

<%
dim sql4,rs4,id,num,pkid
	copyright=trim(request.form("copyright"))

		sql4="select * from siteright"
		set rs4=server.createobject("adodb.recordset")
		rs4.open sql4,conn,1,3
		if rs4.bof or rs4.eof then
				rs4.addnew
				rs4("copyright")=copyright
				rs4.update
		else
				rs4("copyright")=copyright
				rs4.update
		end if
		rs4.close
		set rs4=nothing
	response.write "<script language=JavaScript>" & "alert('保存成功!');"& "window.location.href='siteright.asp';" & "</script>"
%>

