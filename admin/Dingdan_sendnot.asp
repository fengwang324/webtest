<!--#include file="conn.asp"-->
<!--#include file="check4.asp"-->
<%
id=request.form("id")
username=trim(request.form("username"))
allmoney=trim(request.form("allmoney"))
sendornot=request.form("sendornot")

set rs=server.CreateObject("adodb.recordset")
sql="select * from x_bill where id = "&request.form("id")
rs.open sql,conn,1,3
if rs.bof or rs.eof then
	response.write "此单号不存在，可能被其他管理人员刚刚删除。请后退，刷新，重试。"
else

	rs("paystat")=request.form("paystat")
	rs("sendornot")=request.form("sendornot")
	rs("senddate")=trim(request.form("senddate"))
	rs("sendcom")=trim(request.form("sendcom"))
	rs("sendno")=trim(request.form("sendno"))
	rs("sendmemo")=trim(request.form("sendmemo"))
	
	youhj_id=0
	if username<>"" and username<>"非会员" and sendornot="已发货" then
		sql2=" select id from sh_youhj where yhj_qiyong='是' and (yhj_mz_jine<"&allmoney&" or yhj_mz_jine="&allmoney&") and yhj_zs_time>=#"&date()&"# order by yhj_dk_jine desc "
		set rs2=server.CreateObject("adodb.recordset")
		rs2.open sql2,conn,1,1
		if rs2.bof or rs2.eof then
			youhj_id=0
		else
			youhj_id=rs2("id")
		end if
		rs2.close
		set rs2=nothing
	end if
	
	if clng(youhj_id)>0 then
		rs("youhj_id")=youhj_id
		rs("youhj_getdate")=date()
	end if

	rs.update
	response.write "<script language=JavaScript>" & "window.opener.location.reload();window.close();" & "</script>"
	
end if

rs.close
set rs=nothing

call connclose

%>
