<!-- #include file="conn.asp" -->
<!--#include file="check2.asp"-->

<%

dim kindnum,menuid,sql2,rs2
dim kindnumlen,maxkindnum,a

dim sql,rs

kindnum=trim(request.form("kindnum"))
kindname=trim(request.form("kindname"))
menuid=trim(request("menuid"))
kindnumlen=len(kindnum)+4

if kindnumlen>12 then
	response.write "商品分类不能超出3级"
	response.end
end if


sql2="select max(kindnum) as maxkindnum from sh_kind where kindnum like '"&kindnum&"%' and len(kindnum)="&kindnumlen&""

set rs2=conn.execute(sql2)
if rs2.bof or rs2.eof then
	maxkindnum=rs2("maxkindnum")
else
	maxkindnum=rs2("maxkindnum")
end if

'response.write maxkindnum &"***"
'response.end

if isnull(maxkindnum) then
	kindnum=kindnum&"0001"
else
	a=cint(right(maxkindnum,4))
	a=a+1
	a=right("0000"&cstr(a),4)
	kindnum=kindnum&a
end if
rs2.close
set rs2=nothing

'response.write kindnum
'response.end

sql="select * from sh_kind where kindnum='"&kindnum&"'"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
if rs.bof or rs.eof then
	rs.addnew
	rs("kindnum")=kindnum
	rs("kindname")=kindname
	rs("whoaddkind")=session("username")
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	'response.Redirect("b_kind.asp?page="&request.cookies("nowpage_cook"))
	response.Redirect("b_kindadd.asp?kindnum_up="&trim(request.form("kindnum"))&"")
	
else
	response.write "<br><div align=center>此分类编号已存在，请点击 <a href='b_kindadd.asp?menuid="&menuid&"'>返回</a> 重新输入。</div>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if
%>










