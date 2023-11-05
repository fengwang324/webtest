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
.style12 {color: #990000}
-->
</style>
<%
sub subsel(inname,invalue)
	if invalue<>"" then
		response.write "<select name='"&inname&"' style='width:68px'><option value=''></option>"
		
		arr=split(invalue,"，")
		if Ubound(arr)=1 then one="selected"
		for i=0 to Ubound(arr)
			if arr(i)<>"" then
				response.write "<option value='"&arr(i)&"' "&one&">"&arr(i)&"</option>"
			end if
		next
		response.write "</select>"
	end if
end sub
%>
<script language="JavaScript">
<!--
function Juge()
{
<%if s_colorsize="是" then%>
	if (formshow.color.value == "")
	{
		alert("请选择颜色!");
		document.formshow.color.focus();
		return false;
	}
	if (formshow.psize.value == "")
	{
		alert("请选择尺码!");
		document.formshow.psize.focus();
		return false;
	}
<%end if%>
}
//-->
</script>
</head>

<body>
<!-- #include file="head.asp" -->
<TABLE width="960" border=0 align="center" cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
  <tr>
    <td> 



<table width="960" height="352" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
    <td width="210" height="184" valign="top">
		<!-- #include file="inc_left_all.asp" -->
      </td>
    <td width="750" valign="top">
	<table width="99%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height=2></td>
      </tr>
    </table>
            <TABLE width="700" border=0 align="center" cellPadding=0 cellSpacing=2 class=a1>
              <TBODY>
          <TR>
            <TD height=30 colSpan=3><TABLE width="100%" border=0 cellPadding=0 cellSpacing=0 background="images/bg4.gif">
<TBODY>
                <TR>
                  <TD width=41 
                            height=30 background=images/hotnewpro.gif>　　　</TD>
                          <TD width="448"><a href="index.asp"><strong>礼品</strong></a><strong>详细</strong></TD>
                          <TD width=51>&nbsp;</TD>
                </TR>
              </TBODY>
            </TABLE></TD>
          </TR>
		  </tbody>
		  </table>
            <%
dim pkid,model,productname,smallpicpath,price1,price2,pipai

pkid=request("id")
sql="select  * from x_lipin where id = "&pkid

set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof or rs.eof then
	response.write "没有商品记录！"
else

		id=rs("id")
		'picpath=rs("picpath")
		bigpicpath=rs("picpath2")
		lipinno=rs("lipinno")
		lipinname=rs("lipinname")
		zhifen=rs("zhifen")
		'showflag=rs("showflag")
		'num=rs("num")
		memo=rs("memo")
		
		if memo<>"" then
			memo = replace(memo,vbcrlf,"<br>")
			memo = replace(memo,"shop/","")
			'memo = replace(memo," ","&nbsp;")
			'memo = replace(memo,"IMG&nbsp;src=","IMG src=")
			'memo = replace(memo,"&nbsp;border="," border=")
		end if
		'---------更新点击次数begin---------
		conn.execute("update e_product set hit=hit+1 where pkid="&pkid)
		'---------更新点击次数end-----------
%>

		    <TABLE width=700 border=0 align="center" cellPadding=2 cellSpacing=0><TBODY>
              <TR> 
                <TD colSpan=2>&nbsp;</TD>
              </TR>
              <TR> 
                <TD width=415 align=left vAlign=top> <TABLE width=406 height=250 border=0 align="center" cellPadding=0 cellSpacing=1 bgcolor="#CCCCCC">
                    <TBODY>
                      <TR> 
                        <TD align=middle vAlign=top bgcolor="#ffffff" ><DIV align=center><a href="<%=managerfolder%>/<%=bigpicpath%>" target="_blank"><IMG 
                  src="<%=managerfolder%>/<%=bigpicpath%>" name=mainImage  width=340 hspace="2" vspace="6" border="0" id=Image1></a></DIV></TD>
                      </TR>
                    </TBODY>
                  </TABLE></TD>
                <TD width="277" align=left vAlign=top> 
				<TABLE width=270 border=0 align=center cellPadding=0  cellSpacing=0 >
                    <TBODY>
                      <TR> 
                          <TD valign="top" style="PADDING-LEFT: 8px; line-height:180%;"> 
                            <b><span id=cname><font color=ff0000><%=lipinname%></font></span> </b><br>
                          商品编号：<b><span id=memocode><%=lipinno%></span> </b> <br>
                          所需积分：<B><%=zhifen%></B><br> </TD>
                      </TR>
                    </TBODY> </TABLE></TD>
              </TR>
              <TR> 
                <TD colSpan=2>&nbsp;</TD>
              </TR>
              <TR> 
                <TD style="BACKGROUND-REPEAT: repeat-x" background=images/talbe_2_back.jpg 
          colSpan=2 height=1></TD>
              </TR>
              <TR> 
                <TD colSpan=2 height=5></TD>
              </TR>
              <TR> 
                <TD></TD>
              </TR>
              <TR> 
                <TD colSpan=2 height=20> <TABLE cellSpacing=0 cellPadding=0 width="100%" 
            background="images/product5-bg1.gif" border=0>
                    <TBODY>
                      <TR> 
                        <TD width=2><IMG height=31 
                  src="images/product5-c1.gif" width=2></TD>
                        <TD vAlign=top align=middle><IMG height=31 
                  src="images/product5-tittle.gif" 
                width=119></TD>
                        <TD width=2><IMG height=31 
                  src="images/product5-c2.gif" 
              width=2></TD>
                      </TR>
                    </TBODY>
                  </TABLE>
                  <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                    <TBODY>
                      <TR> 
                        <TD width=1 bgColor=#95e3d5><IMG height=1 
                  src="images/space.gif" width=1></TD>
                        <TD align=middle> <TABLE cellSpacing=0 cellPadding=0 width="100%" 
                  background="images/product5-dotline.gif" 
                  border=0>
                            <TBODY>
                              <TR> 
                                <TD><IMG height=1 
                        src="images/space.gif" 
                    width=1></TD>
                              </TR>
                            </TBODY>
                          </TABLE>
                          <TABLE cellSpacing=0 cellPadding=0 border=0>
                            <TBODY>
                              <TR> 
                                <TD><IMG height=5 
                        src="images/space.gif" 
                    width=5></TD>
                              </TR>
                            </TBODY>
                          </TABLE>
                          <TABLE cellSpacing=0 cellPadding=10 width="100%" border=0>
                            <TBODY>
                              <TR> 
                                <TD align=left><%=memo%>
                                  <p>&nbsp;</p></TD>
                              </TR>
                            </TBODY>
                          </TABLE></TD>
                        <TD width=1 bgColor=#95e3d5><IMG height=1 
                  src="images/space.gif" 
            width=1></TD>
                      </TR>
                    </TBODY>
                  </TABLE>
                  <TABLE cellSpacing=0 cellPadding=0 width="100%" 
            background="images/product5-bg2.gif" border=0>
                    <TBODY>
                      <TR> 
                        <TD align=left><IMG height=8 src="images/product5-c3.gif" width=9></TD>
                        <TD align=right><IMG height=8 src="images/product5-c4.gif" width=9></TD>
                      </TR>
                    </TBODY>
                  </TABLE>
				  <br>
                  <br>
                </TD>
              </TR></TBODY>
            </TABLE>
		 <%	
		end if
		rs.close
		set rs=nothing

		 %>


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

