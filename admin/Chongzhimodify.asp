<!--#include file="conn.asp"-->
<!--#include file="check4.asp"-->
<%
dim id,customid,operator,chongzhidate,constract,actionkind,actionglass,actionsubject,actionaddr,alertyesno,alertdate
dim sql,rs
	id=request.form("id")
	menuid=request("menuid")
czdate=trim(request.form("czdate"))
customid=trim(request.form("customid"))

czby=trim(request.form("czby"))
czmoney=trim(request.form("czmoney"))

czmemo=trim(request.form("czmemo"))

confirmflag=trim(request.form("confirmflag"))
operator=trim(request.form("operator"))

sql="select * from sh_chongzhi where id="&id
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3

	rs("czdate")=czdate
	rs("customid")=customid
	
	rs("czby")=czby
	rs("czmoney")=czmoney
	
	rs("czmemo")=czmemo
	rs("confirmflag")=confirmflag
	rs("operator")=session("m_user")&"("&now()&")"
	
	rs.update
	rs.close
	set rs=nothing
conn.close
set conn=nothing

response.Redirect("chongzhi.asp?menuid="&menuid&"&sublist="&request.Cookies("sublist"))	

%>








