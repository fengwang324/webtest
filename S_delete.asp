<!--#include file="conn.asp"-->
<%

if session("m_user")<>"" then

	if request("all")="1" then
		sql="delete * from message where id="& request("id") &" or  parentid="&request("id")&""
		conn.execute(sql)
		'response.Write"<br><br><br><div align=center><font color=ff0000>成功删除!</font>       <a href=message.asp><font color=ff0000>返回留言主页</font></a></div>"
		response.write "<script language=JavaScript>" & "alert('成功删除！');" &"window.location.href='message2.asp';" & "</script>"
	else
		sql="delete * from message where id="& request("id") &""
		conn.execute(sql)
		
		id=request("parentid")
		set rs=server.createobject("adodb.recordset")
		sql="select count(id) as allre from message where parentid="& id &""
		rs.open sql,conn,1,1
		if rs.bof or rs.eof then
			allre="0"
		else
			allre=rs("allre")
			if allre="" or isnull(allre) then allre="0"
		end if
		rs.close
		set rs=nothing
		conn.execute("update message set recount='"&allre&"' where id="&id)
		'response.Write"<br><br><br><div align=center><font color=ff0000>成功删除!</font>       <a href=showmessage.asp?id="&id&"><font color=ff0000>返回留言</font></a></div>"
		response.write "<script language=JavaScript>" & "alert('成功删除！');" &"window.location.href='message2.asp';" & "</script>"
	end if
	
end if

conn.close
set conn=nothing

%>
