<style type="text/css">
<!--
.style1 {font-weight: bold}
.style2 {
	color: #993333;
	font-weight: bold;
}
-->
</style>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<TABLE width="960" border=0 align="center" cellPadding=0 cellSpacing=0>
  <!--DWLayoutTable-->
<TBODY>
    <TR>
      <TD width="4%" background=images/copyrightbg.gif height=30><DIV align=center></DIV></TD>
      <TD width="88%" background=images/copyrightbg.gif>
	  <DIV align=center>
             <%
		sql4="select l_id,menukind,l_title,zhidingurl,newopen from e_left where showflag='是' and bigmenu='否' order by bigmenu desc,l_num"
		set rs4=server.createobject("adodb.recordset")
		rs4.open sql4,conn,1,1
		if rs4.bof or rs4.eof then
		else
			i=1
			
			do while not rs4.eof
				l_id=rs4("l_id")
				menukind=rs4("menukind")
				l_title=rs4("l_title")
				zhidingurl=rs4("zhidingurl")
				newopen=rs4("newopen")
				
				if i>1 then response.write " | "
				
				if len(zhidingurl)>1 then
						response.write "<a href='"&zhidingurl&"'"
				else
					if menukind="产品列表" then
						response.write "<a href='productlist.asp?l_id="&l_id&"'"
					elseif menukind="普通文章列表" then
						response.write "<a href='news.asp?l_id="&l_id&"'"
					elseif menukind="文章列表并指定链接" then
						response.write "<a href='"&zhidingurl&"'"
					elseif menukind="单个说明" then
						response.write "<a href='showone.asp?l_id="&l_id&"'"
					elseif menukind="指定链接" then
						response.write "<a href='"&zhidingurl&"'"
					else
					end if
				end if
				if newopen="是" then response.write " target='_blank' "
				response.write " style='font-size: 12px;line-height: 150%;'>"&l_title&"</a>"
				
				
			rs4.movenext
  			i=i+1
  			loop			
		end if
	rs4.close
	set rs4=nothing		
%>
	  </DIV></TD>
      <TD width="8%" background=images/copyrightbg.gif><a href='javascript:window.scroll(0,0)'><IMG 
      src="images/top.gif" align=absMiddle border=0></a></TD>
    </TR>
  </TBODY>
</table>
	
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr> 
    <td width="249" rowspan="2" valign="bottom"><img src="images/wei1.gif" width="249" height="190"></td>
    <td height="120" colspan="2"><TABLE width="100%" border=0 align="center" cellPadding=0 cellSpacing=0>
<TBODY>
          <TR> 
            <TD width="33%" height=120> 
<DIV class=a1 align=center> 
                <p align="right"><img src="images/logo.jpg"></p>
              </DIV></TD>
            <TD width="2%" height=105>&nbsp;</TD>
            <TD width="65%" height=105> 
              <%
			sql4="select * from siteright"
			set rs4=server.createobject("adodb.recordset")
			rs4.open sql4,conn,1,1
			if rs4.bof or rs4.eof then
					copyright=""
			else
					copyright=rs4("copyright")
			end if
			rs4.close
			set rs4=nothing
			response.write copyright
			%>
            </TD>
          </TR>
        </TBODY>
      </TABLE></td>
  </tr>
  <tr> 
    <td width="373"><img src="images/wei2.gif" width="373" height="78"></td>
    <td width="338" background="images/wei3.gif">&nbsp;</td>
  </tr>
</table>
