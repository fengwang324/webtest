<!--#include file="conn.asp"-->
<!--#include file="check2.asp"-->
<%
c_parent=request.form("c_parent")
c_title=trim(request.form("c_title"))
c_contect=trim(request.form("c_contect"))
c_num=request.form("c_num")
c_ubb=request.form("c_ubb")
url=trim(request.form("url"))
cuxflag=trim(request.form("cuxflag"))
if len(c_contect)>5 then 
detail="1"
else
detail="0"
end if
c_addtime=date()

arr=split(c_parent,"|")

set rs=server.CreateObject("adodb.recordset")
sql="select * from e_contect"
rs.open sql,conn,1,3

	rs.addnew
	rs("c_parent1")=arr(0)
	rs("c_parent2")=arr(1)
	rs("c_title")=c_title
	rs("c_contect")=c_contect
	rs("c_num")=c_num
	rs("c_ubb")=c_ubb
	rs("url")=url
	rs("c_addtime")=c_addtime
	rs("detail")=detail
	rs("cuxflag")=cuxflag
	rs.update


rs.close
set rs=nothing

call connclose

	response.redirect "conlist.asp?l_id="&request.QueryString("newskind")
%>
