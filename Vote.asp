<!--#include file="conn.asp"-->
<%
dim votenews
votenews=""
Submit=request.form("Submit")
rad=request.form("rad")

'response.write Submit & "pp"&rad
'response.end

'response.write request.cookies("votestat")
'response.end

if rad<>"" and request.cookies("votestat")<>"Y" then
	sql="update e_diaocha2 set shuliang=shuliang+1 where id="&rad
	'response.write sql
	'response.end
	conn.execute(sql)
	votenews="1"
	response.cookies("votestat")="Y"
end if
%>

<%
if s_head="head4.asp" or s_productkind="4" then
	response.write "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">"
else
	response.write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">"
end if
%>
<html>
<head>
<title>.::<%=application("sitename")%>::. 最新调查结果</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<style>
<!--
td{font-size:9pt}
body {
	margin-top: 0px;
}
.style3 {color: #A84E22; font-weight: bold; }
-->
</style>

</head>

<body>
<table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="38"><font color="#FF0000" size="3"><strong><%=application("sitename")%>最新调查</strong></font></td>
  </tr>
  <tr>
    <td height="20">
	<%
	if votenews="1" then
		response.write "谢谢您的一票。谢谢您的关注。"
	else
		response.write "您已经投过一票。谢谢您的关注。"
	end if
	%>
	</td>
  </tr>

<tr>
        
    <td width="602"  valign="top"> 
      <%
sql="select top 1 * from e_diaocha1 order by id desc"
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,1

if rs.bof or rs.eof then
	response.write "还没有调查记录。"
else
	dcid=rs("id")
	dctitle=rs("dctitle")
	adddate=rs("adddate")
%>
      <table width="585" border="0" cellpadding="0" cellspacing="0" bgcolor="#D2FFD2">
        <tr> 
          <td width="447" height="35"><font color=FF0000><b>&nbsp;题目：<%=dctitle%></b></FONT> 
          </td>
          <td width="138">开始日期：<%=adddate%></td>
        </tr>
      </table>
      <table width="585" border="0" cellpadding="0" cellspacing="0" bgcolor="#F5F5F5">
        <%
sql2="select sum(shuliang) as allshuliang from e_diaocha2 where contect<>'' and dcid="&dcid&" "
set slrs=conn.execute(sql2)
allshuliang=slrs("allshuliang")
slrs.close
set slrs=nothing
if allshuliang="" or isnull(allshuliang) or allshuliang=0 then 
	response.write "对不起，还没有投票数。"
	response.end
end if

i=1
sql2="select * from e_diaocha2 where contect<>'' and dcid="&dcid&" order by id"
set rs2=server.CreateObject("adodb.recordset")
rs2.open sql2,conn,1,1
if rs2.bof or rs2.eof then
else
	do while not rs2.eof
		id=rs2("id")
		contect=rs2("contect")
		shuliang=rs2("shuliang")
		percent=shuliang/allshuliang*100
%>
        <tr> 
          <td width="32" height="32" align="right"><b><%=i%></b>&nbsp;</td>
          <td width="193"><%=contect%> [<%=shuliang%>票 <font color="#FF0000"><%=cint(percent)%>%</font>]</td>
		  <td width="337"> 
            <table width="<%=percent%>%" height="15" border="0" cellpadding="0" cellspacing="0">
				<tr><td bgcolor="#0099FF"></td></tr>
				</table>
		 </td>
		  <td width="23"> </tr>
        <%
		rs2.movenext
		i=i+1
	loop
end if
rs2.close
set rs2=nothing
%>
        <tr> 
          <td width="32" height="32">&nbsp;</td>
          <td width="193">&nbsp;</td>
		  <td width="337">&nbsp; </td>
        </tr>
      </table>
      <%
end if
%>
<br>
<div align="center"><a href=javascript:window.close()><b>关闭窗口</b></a></div>
    </td>
  </tr>
</table>
<%
conn.close
set conn=nothing
%>
</body>
</html>


