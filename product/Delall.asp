<!-- #include file="conn.asp" -->
<!--#include file="check2.asp"-->
<%

dim i,strid,sql,rs
dim sql2,rs2
dim newsid1,newsid2,news1,news2

for i=1 to 30
	strid=request.form("checkbox"&i)
	if strid<>"" then


		sql="select * from e_product where pkid="&strid&""
		set rs=conn.execute(sql)
		if rs.bof or rs.eof then
		else
			smallpicpathold=rs("smallpicpath")
			bigpicpathold=rs("bigpicpath")
		
			sql="delete * from e_product where pkid="&strid&""
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
		
	
	end if
next


response.redirect "productlist.asp?menuid="&request("menuid")


%>


