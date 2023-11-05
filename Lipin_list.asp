<!--#include file="conn.asp"-->
<%
Response.Buffer = True 
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
%>
<%
if session("username")="" then
	conn.close
	set conn=nothing
	response.Redirect("hyzq.asp")
	response.end
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
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=application("sitename")%></title>
<link href="i.css" type=text/css rel=stylesheet>
<style type="text/css">
<!--

.style1 {font-size: 14px}
.style12 {color: #990000}
-->
</style></head>

<body>
<!-- #include file="head.asp" -->
<TABLE width="960" border=0 align="center" cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
  <tr>
    <td> 


<table width="960" height="352" border="0" align="center" cellpadding="0" cellspacing="0">
<tr> 
          <td height="184" valign="top" bgcolor="#F7F7F7"> 
<table width="958"  border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td height=1></td>
              </tr>
            </table>
            <TABLE width="958" border=0 align="center" cellPadding=0 cellSpacing=2 class=a1>
              <TBODY>
                <TR> 
                  <TD height=30 colSpan=3>
				  <TABLE width="100%" border=0 cellPadding=0 cellSpacing=0 background="images/bg4.gif">
<TBODY>
                        <TR> 
                          <TD width="621" height=30> <strong>我的礼品：</strong></TD>
                          <TD width="137"><a href="javascript:window.history.back()">&lt;&lt; 
                            返回</a></TD>
                        </TR>
                      </TBODY>
                    </TABLE></TD>
                </TR>
              </tbody>
            </table>
            <table width="958" border="0" align="center" cellpadding="2" cellspacing="1">
              <tr bgcolor="#5C9153"> 
                <td width="32" height="23" > 
					<div align="center"><font color="#FFFFFF"><strong>序号</strong></font></div></td>

                <td width="55"> 
                  <div align="center"><font color="#FFFFFF"><strong>填写时间</strong></font></div></td>
                <td width="55"> 
                  <div align="center"><font color="#FFFFFF"><strong>用户名</strong></font></div></td>
                <td width="53"> 
                  <div align="center"><font color="#FFFFFF"><strong>礼品名称</strong></font></div></td>
                <td width="53"> 
                  <div align="center"><font color="#FFFFFF"><strong>单个积分</strong></font></div></td>
                <td width="53"> 
                  <div align="center"><font color="#FFFFFF"><strong>数量</strong></font></div></td>
                <td width="57"> 
                  <div align="center"><font color="#FFFFFF"><strong>所需积分</strong></font></div></td>
                <td width="65"> 
                  <div align="center"><font color="#FFFFFF"><strong>处理状态</strong></font></div></td>

              </tr>
              <%
		dim sql4,rs4,id,num,pkid
		dim sql5,rs5,model,productname,price2,price,money
		
		kind=request("kind")
		nowpage=request("page")

		sql4="select * from x_huanlipin where username='"&session("username")&"' order by id desc"
		set rs4=server.createobject("adodb.recordset")
		rs4.open sql4,conn,1,1
		if rs4.bof or rs4.eof then
			response.write "没有礼品记录！"
		else
		i=1
		rs4.pagesize=20
		rs4.absolutepage=1
		if request("page")<>"" then rs4.absolutepage=request("page")
		rowcount=rs4.pagesize
		
			do while not rs4.eof and rowcount>0
				id=rs4("id")
				lipindate=rs4("lipindate")
				username=rs4("username")
				lipinname=rs4("lipinname")
				zhifen=rs4("zhifen")
				shu=rs4("shu")
				kjzhf=rs4("kjzhf")
				
				sendstat=rs4("sendstat")
		
				response.write "<tr bgcolor='#Ffffff' onMouseOver=""this.style.background='#f1f1f1';"" onMouseOut=""this.style.background='#ffffff';"">"
				response.write "<td height='23' align='center'>"&i&"</td>"
				response.write "<td>"&lipindate&"</td>"
				response.write "<td>"&username&"</td>"
				response.write "<td>"&lipinname&"</td>"

				response.write "<td>"&zhifen&"</td>"
				response.write "<td>"&shu&"</td>"
				response.write "<td>"&kjzhf&"</td>"
				response.write "<td>"&sendstat&"</td>"
				
				response.write "</tr>"&vbcrlf
				rs4.movenext
				i=i+1
				rowcount=rowcount-1
			loop
		end if
			  %>
            </table>
            <table width="956" border="0" align="center" cellpadding="0" cellspacing="0">
              <TR bgColor=#ffffff> 
                <TD  height="32" align=middle colSpan=6>
<%
if nowpage="" then nowpage=1
nowpath=request.servervariables("script_name")
kind=server.URLEncode(kind)
response.write "<table width='550' align='center'>"
  response.write "<tr> "
   response.write " <td height='22' align=right style='font-size:9pt;'> "
		response.write "共"&rs4.recordcount&"条记录&nbsp;每页20条&nbsp;现在 "&nowpage&"/"&rs4.pagecount &" 页&nbsp;"	
        response.write "<input type='button' name='Button3' value='首页' onclick=window.location.href='"&nowpath&"?kind="&kind&"&none='>&nbsp;"
        if clng(nowpage)>1 then
        response.write "<input type='button' name='Button4' value='上页' onclick=window.location.href='"&nowpath&"?page="&nowpage-1&"&kind="&kind&"&none='>&nbsp;"
        else
        response.write "<input type='button' name='Button4' value='上页'>&nbsp;"
        end if
		if clng(nowpage)<>clng(rs4.pagecount) then
        response.write "<input type='button' name='Button5' value='下页' onclick=window.location.href='"&nowpath&"?page="&nowpage+1&"&kind="&kind&"&none='>&nbsp;"
        else
        response.write "<input type='button' name='Button5' value='下页'>&nbsp;"
		end if
		response.write "<input type='button' name='Button6' value='末页' onclick=window.location.href='"&nowpath&"?page="&rs4.pagecount&"&kind="&kind&"&none='>"
    response.write "</td>"
  response.write "</tr>"
response.write "<tr><td height=60>&nbsp;</td></tr></table>"
%>
				</TD>
              </TR>
			  
            </table>
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

