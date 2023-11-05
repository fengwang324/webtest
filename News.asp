<!--#include file="conn.asp"-->
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
.style13 {color: #333333}
.style15 {font-size: 14px; color: #669933; font-weight: bold; }
-->
</style></head>

<body>
<!-- #include file="head.asp" -->
<TABLE width="1020" border=0 align="center" cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
  <tr>
    <td> 

    <table width="1020" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr>
          <td width="210" valign="top">
		  <!-- #include file="inc_left_all.asp" -->
		  </td>
          <td width="750" valign="top">
	
            <TABLE class=text width="98%" align=center cellSpacing=0 cellPadding=0 border=0>
              <TBODY>
                <TR> 
                  <TD vAlign=bottom height=32><IMG height=19 src="images/dian04.gif" width=10 align=absMiddle> <STRONG>&nbsp;<FONT color=#ff6600>
<%
sql3="select l_title from e_left where l_id="&request.QueryString("l_id")&""
set rs3=server.CreateObject("adodb.recordset")
rs3.open sql3,conn,1,1

if rs3.bof or rs3.eof then
else
	l_title=rs3("l_title")
	response.write l_title
end if
rs3.close
set rs3=nothing
%>
			</FONT></STRONG></TD>
                </TR>
                <TR> 
                  <TD vAlign=top height=25> <DIV align=left><IMG height=25 src="images/line.gif" width=457></DIV></TD>
                </TR>
			  <form name="form5" method="post" action="news.asp?l_id=<%=request.QueryString("l_id")%>">
                <TR> 
                  <TD height=40 valign="top">
				&nbsp; &nbsp; &nbsp; &nbsp;搜索：&nbsp;关键字 <input name="searchkey" type="text" id="searchkey" size="40">&nbsp;&nbsp;
				  <input  type=submit name="submit" value=" 查询 "  style="BACKGROUND-COLOR: #cccccc; BORDER: #AAA 1px groove; COLOR: #000000; ">
                      
				  </TD>
                </TR>
			  </form>
                <TR> 
                  <TD height="320" vAlign=top>
                    <%

searchkey=request("searchkey")
searchkind=request("searchkind")

if searchkey<>"" then
	sql3="select * from e_contect where c_parent2="&request.QueryString("l_id")&" and (c_title like '%"&searchkey&"%' or c_contect like '%"&searchkey&"%') order by c_num desc,c_addtime desc"
else
	sql3="select * from e_contect where c_parent2="&request.QueryString("l_id")&" order by c_num desc,c_addtime desc"
end if
'response.write sql3
'response.end

set rs3=server.CreateObject("adodb.recordset")
rs3.open sql3,conn,1,1

if rs3.bof or rs3.eof then
	response.write "&nbsp;&nbsp;没有符合条件的信息。"
else
	rs3.pagesize=25
	rs3.absolutepage=1
	if request("page")<>"" then rs3.absolutepage=request("page")
	rowcount =rs3.pagesize
	
	response.write "<table border=0  cellpadding='0' cellspacing='0' width='100%' align=center>"
		
		do while not rs3.eof and rowcount>0
		'set c_parent2=rs3("c_parent2")
		set c_id=rs3("c_id")
		set c_title=rs3("c_title")
		set c_addtime=rs3("c_addtime")
		set detail=rs3("detail")
		'c_addtime=c_addtime
		if detail="1" then
			response.Write "<tr><td height=25 width='3%' align=center><img src=images/king_822.jpg></td>"
			response.Write "<td width='90%'><a href=show_all.asp?c_id="&c_id &" class=large>"&c_title &"</a></td>"
			response.Write "<td width='17%' align=center><font color=#666666>("&c_addtime&")</font></td></tr>"
			response.Write "<tr><td colspan=3 background=images/newline.gif height=2></td></tr>"
		else
			response.Write "<tr><td height=25 width='3%' align=center><img src=images/dian02.gif></td>"
			response.Write "<td width='90%'><a href=#>"&c_title &"</a></td>"
			response.Write "<td width='17%' align=center><font color=#666666>("&c_addtime&")</font></td></tr>"
			response.Write "<tr><td colspan=3 background=images/newline.gif height=2></td></tr>"

		end if
		rs3.movenext
		rowcount=rowcount-1
		loop
		
		response.write "<tr><td colspan=3 height=8></td></tr><tr><td colspan=3 align=right>每页25条 页码："
		 nowpage=request("page")
		if nowpage="" then nowpage=1
		for i=1 to rs3.pagecount 
		if cint(nowpage)=i then
			response.write "<font color=#ff0000>[<b>" & i & "</b>]</font>&nbsp;"
		else
			response.write "<a href='news.asp?page=" & i & "&l_id=" & request("l_id") & "&c_p="&request("c_p")&"&searchkey="&request("searchkey")&"&searchkind="&request("searchkind")&"'>[<b>" & i & "</b>]</a>&nbsp;"
		end if
		next 
		response.write "</td></tr>"
	response.write "</table>"
end if
rs3.close
set rs3=nothing

%>
                  </TD>
                </TR>
                <TR> 
                  <TD height="25">&nbsp;</TD>
                </TR>
              </TBODY>
            </TABLE>
            
          </td>
  </tr>
</table>


    </td>
  </tr>

<TR><TD>

</TD></TR>
</table>
<!-- #include file="foot.asp" -->
<%
conn.close
set conn=nothing
%>
</body>
</html>



