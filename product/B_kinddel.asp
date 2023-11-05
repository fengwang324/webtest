<!-- #include file="conn.asp" -->
<!--#include file="check2.asp"-->
<%

dim i,strid,sql,rs
dim sql2,rs2
dim newsid1,newsid2,news1,news2

for i=1 to 50
	strid=request.form("checkbox"&i)
	if strid<>"" then
		
		sql2="select kind from e_product where kind='"&strid&"'"
		set rs2=conn.execute(sql2)
		if rs2.bof or rs2.eof then
		
				sql="select kindnum from sh_kind where kindnum like '"&strid&"%'"
				set rs=server.createobject("adodb.recordset")
				rs.open sql,conn,1,3
				num=rs.recordcount
				if cint(num)=1 then
					rs.delete
				else
					newsid1="1"
					news1=news1 & strid &"、"
				end if
				rs.close
				set rs=nothing

		else
			newsid2="2"
			news2=news2 & strid &"、"
		end if
		rs2.close
		set rs2=nothing
	end if
next

if newsid1="" and newsid2="" then
	response.redirect "b_kind2.asp?menuid="&request("menuid")
else
	if newsid1="1" then
		response.write "<br><br><div align=center>"&news1 & "不能删除，因为此分类有下级分类。<a href=b_kind2.asp?menuid="&request("menuid")&">返回</a></div>"
	end if
	if newsid2="2" then
		response.write "<br><br><div align=center>"&news2 & "不能删除，因为此分类有商品相关联。<a href=b_kind2.asp?menuid="&request("menuid")&">返回</a></div>"
	end if	
end if
%>










