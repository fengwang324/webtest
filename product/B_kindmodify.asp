<!-- #include file="conn.asp" -->
<!--#include file="check2.asp"-->
<%

dim kindnum,menuid,sql2,rs2
dim kindnumlen,maxkindnum,a

dim sql,rs

id=trim(request.form("id"))
kindnum=trim(request.form("kindnum"))
kindnumold=trim(request.form("kindnumold"))
kindname=trim(request.form("kindname"))
productflag=trim(request.form("productflag"))
menuid=trim(request("menuid"))

if len(kindnumold)<>len(kindnum) then
	response.write "不能修改。本次填写的编号与原来的<b>位数</b>不同。"
	response.end
end if
if len(kindnum) mod 4 >0 then
	response.write "不能修改。本次填写的编号并不是<b>4的倍数</b>。"
	response.end
end if

if kindnumold<>kindnum then

	sql0="select * from sh_kind where kindnum like '"&kindnum&"%'"
	set rs=server.createobject("adodb.recordset")
	rs.open sql0,conn,1,1
	if rs.bof or rs.eof then
	
		sql2="update sh_kind set kindnum='"&kindnum&"'+mid(kindnum,"&len(kindnum)+1&") where kindnum like '"&kindnumold&"%'"
		conn.execute(sql2)

		sql3="update e_product set kind='"&kindnum&"'+mid(kind,"&len(kindnum)+1&") where kind like '"&kindnumold&"%'"
		conn.execute(sql3)

	else
		response.write "不能修改。本次填写的编号和已有的编号一样，请填写其他编号。"
		response.end
	end if
	rs.close
	set rs=nothing

end if

call checkchar(kindname)

sub checkchar(info)
		info=replace(info,"	","")
		info=replace(info," ","")
		info=replace(info,"&","/")
	for i=1 to len(info)
	aa=mid(info,i,1)
	if aa=" " or aa="&" then
	response.write "<script language=JavaScript>" & "alert('类别或商品名称不能含有空格或者&符号，请用其他符号。');" & "history.back()" & "</script>"
	response.end
	end if
	next
end sub

sql="select * from sh_kind where id="&id
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
if rs.bof or rs.eof then
	response.write "<br><div align=center>保存失败，请点击 <a href='b_kindadd.asp?menuid="&menuid&"'>返回</a> 重试。</div>"

	rs.close
	set rs=nothing
	conn.close
	set conn=nothing

else

	'rs("kindnum")=kindnum
	rs("kindname")=kindname
	'rs("productflag")=productflag
	rs("whoaddkind")=session("username")
	rs.update
	rs.close
	set rs=nothing
	
	
	conn.close
	set conn=nothing
	response.Redirect("b_kind.asp?page="&request.cookies("nowpage_cook"))	
end if
%>










