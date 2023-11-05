
<!-- #include file="conn.asp" -->
<!--#include file="check2.asp"-->
<%

pkid=request("pkid")




sql="select * from e_product where pkid="&pkid&""
set rs=conn.execute(sql)
if rs.bof or rs.eof then
else
	smallpicpathold=rs("smallpicpath")
	bigpicpathold=rs("bigpicpath")

	sql="delete * from e_product where pkid="&pkid&""
	conn.execute(sql)

	'on error resume next
		set fso= server.CreateObject("scripting.fileSystemObject")
		if smallpicpathold<>"" then
		set f1=fso.getfile(server.mappath(""&smallpicpathold&""))
		f1.delete
		end if
		if bigpicpathold<>"" then
		set f2=fso.getfile(server.mappath(""&bigpicpathold&""))
		f2.delete
		end if
end if
rs.close
conn.close

	response.write "<script language=JavaScript>" & "alert('成功删除！');" & "window.location.href='productlist.asp'" & "</script>"


%>
