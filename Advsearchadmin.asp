<!--#include file="conn.asp"-->
<!--#include file="inc_product_list.asp"-->
<%
if s_head="head4.asp" or s_productkind="4" then
	response.write "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">"
else
	response.write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">"
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=application("sitename")%></title>
<link href="i.css" type=text/css rel=stylesheet>
<style type="text/css">
<!--

.style1 {font-size: 14px}
.style11 {
	font-size: 14px;
	color: #990000;
	font-weight: bold;
}
.style12 {color: #990000}
-->
</style>

<script language="javascript">
<!--
function SetBgColor(Menu)
{
		Menu.style.background="#F3F3F5";
}
function RestoreBgColor(Menu)
{
		Menu.style.background="#ffffff";
}
-->
</script>

</head>

<body>
<!-- #include file="head.asp" -->
<TABLE width="960" border=0 align="center" cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
  <tr><td>

	
            <TABLE width="960" border=0 align="center" cellPadding=0 cellSpacing=2 class=a1>
              <TBODY>
          <TR>
            <TD  height=30 colSpan=3><TABLE width="100%" border=0 cellPadding=0 cellSpacing=0 background="images/bg4.gif">
				<TBODY>
                <TR>
                          <TD width=11% height=30 background=images/hotnewpro.gif>　　　<span class="style11">商品列表</span></TD>
                          <TD width="88%"><strong>&gt;&gt; 高级搜索结果：</strong></TD>
                          <TD width=1%>&nbsp;</TD>
                </TR>
              </TBODY>
            </TABLE></TD>
          </TR>
		  </tbody>
		  </table>
<%
dim kind,nowpage
dim pkid,model,productname,smallpicpath,price1,price2,pipai
dim searchkind,keyword

'kind=request("kind")
nowpage=request("page")

'searchkind=trim(request("searchkind"))
'keyword=trim(request("keyword"))
kindnum=trim(request("kindnum"))
pipai=trim(request("pipai"))
model=trim(request("model"))
productname=trim(request("productname"))
price11=trim(request("price11"))
price12=trim(request("price12"))
price21=trim(request("price21"))
price22=trim(request("price22"))
if price11="" then price11="0"
if price12="" then price12="9999999"
if price21="" then price21="0"
if price22="" then price22="9999999"

sqlcon=" where pkid>0 and updown='1' "
if kindnum<>"" then
	sqlcon = sqlcon &" and kind like '"&kindnum&"%' "
end if
if pipai<>"" then 
	sqlcon = sqlcon &" and pipai like '%"&pipai&"%' "
end if
if model<>"" then
	sqlcon = sqlcon &" and model like '%"&model&"%' "
end if
if productname<>"" then
	sqlcon = sqlcon &" and productname like '%"&productname&"%' "
end if

sql="select pkid,model,productname,smallpicpath,price1,price"&session("customkind")&",kindname,pipai,addtime from view_product "&sqlcon&" order by pkid desc"
'response.write sql

set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof or rs.eof then
	response.write "没有商品记录！"
else




i=1
%>
                          <TABLE align=center border=0 cellPadding=2 cellSpacing=0 width="100%">
                            <TBODY>
                              <TR> 
                                
<%
	rs.pagesize=20
	rs.absolutepage=1
	if request("page")<>"" then rs.absolutepage=request("page")
	rowcount=rs.pagesize
	if not rs.eof then
		do while not rs.eof and rowcount>0
		pkid=rs("pkid")
		model=rs("model")
		productname=rs("productname")
		smallpicpath=rs("smallpicpath")
		price1=rs("price1")
		price2=rs("price"&session("customkind"))
		pipai=rs("pipai")
		kindname=rs("kindname")
		addtime=rs("addtime")

		call product_list(2)
		
		if i=s_col then
			response.write "</tr><tr><td height=8></td></tr><tr>"
			i=1
		else
			i=i+1
		end if

		rs.movenext
		rowcount=rowcount-1
		loop
	end if

%>
							   
                              </TR>
                              <TR> 
                                <TD height="5" class="a1"></TD>
                              </TR>
                            </TBODY>
                          </TABLE>

<%
if nowpage="" then nowpage=1
nowpath=request.servervariables("script_name")
kind=server.URLEncode(kind)
response.write "<table width='550' align='center'>"
  response.write "<tr> "
	response.write "<td height='22' align=right style='font-size:9pt;'> "
	kindnum=trim(request("kindnum"))
	pipai=trim(request("pipai"))
	model=trim(request("model"))
	productname=trim(request("productname"))
	kind="&kindnum="&kindnum&"&pipai="&pipai&"&model="&model&"&productname="&productname&""	
		response.write "共"&rs.recordcount&"条记录&nbsp;每页20条&nbsp;现在 "&nowpage&"/"&rs.pagecount &" 页&nbsp;"	
        response.write "<input type='button' name='Button3' value='' style='background-image:url(images/z1.jpg);width:37px;height:19px;border:0px' onclick=window.location.href='"&nowpath&"?page=1"&kind&"&none='>&nbsp;"
        if clng(nowpage)>1 then
        response.write "<input type='button' name='Button4' value='' style='background-image:url(images/z2.jpg);width:45px;height:19px;border:0px' onclick=window.location.href='"&nowpath&"?page="&nowpage-1&kind&"&none='>&nbsp;"
        else
        response.write "<input type='button' name='Button4' value='' style='background-image:url(images/z2.jpg);width:45px;height:19px;border:0px'>&nbsp;"
        end if
		if clng(nowpage)<>clng(rs.pagecount) then
        response.write "<input type='button' name='Button5' value='' style='background-image:url(images/z3.jpg);width:45px;height:19px;border:0px' onclick=window.location.href='"&nowpath&"?page="&nowpage+1&kind&"&none='>&nbsp;"
        else
        response.write "<input type='button' name='Button5' value='' style='background-image:url(images/z3.jpg);width:45px;height:19px;border:0px'>&nbsp;"
		end if
		response.write "<input type='button' name='Button6' value='' style='background-image:url(images/z4.jpg);width:37px;height:19px;border:0px' onclick=window.location.href='"&nowpath&"?page="&rs.pagecount&kind&"&none='>"
    response.write "</td>"
  response.write "</tr>"
response.write "</table>"

end if
rs.close
set rs=nothing

%> 

    </td>
  </tr>

</table>
<!-- #include file="foot.asp" -->
<%
conn.close
set conn=nothing
%>
</body>
</html>

