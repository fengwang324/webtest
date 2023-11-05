
<%
nowpath = Request.ServerVariables("SCRIPT_NAME")
if instr(nowpath,"news.asp")>0 then
else
%>
 <TABLE width="210" border=0 cellPadding=2 cellSpacing=0 class="weitab"  style="margin-top:5px;">
  <tr>
	  <td height="32" align="center" background="images/index_png24_a1.png"><font color="#999999"><strong>最新商訊</strong></font></td>
  </tr>
			<%
			sql="select top 10 c_id,c_title,detail from e_contect where c_parent2=30 order by c_num desc,c_addtime desc,c_id desc"
			set rsv=server.CreateObject("adodb.recordset")
			rsv.open sql,conn,1,1
			if rsv.bof or rsv.eof then
				response.write "<tr><td><div align=center>沒有記錄!</div></td></tr>"
			else
			
				do while not rsv.eof
					set c_id=rsv("c_id")
					set c_title=rsv("c_title")
					if len(c_title)>29 then 
					c_title2=left(c_title,28)&"..."
					else
					c_title2=c_title
					end if
					set detail=rsv("detail")
					if detail="1" then
						response.write "<TR><TD style='padding-left:4px;padding-right:4px;padding-top:3px;'>·<a href='show_all.asp?c_id="&c_id &"' title='"&c_title&"'>"&c_title2&"</a></td></tr>"&vbcrlf   
					else
						response.write "<TR><TD style='padding-left:4px;padding-right:4px;padding-top:3px;'>·<a href='news.asp?l_id=30' title='"&c_title&"'>"&c_title2&"</a></td></tr>"&vbcrlf
					end if
					rsv.movenext
				loop
					response.write "<TR>" 
					response.write "<TD height='20' align=right><a href='news.asp?l_id=30'>更多>></a>&nbsp;</TD>"
					response.write "</TR>"
			end if
			rsv.close
			set rsv=nothing
		    %>
</table>
<%end if%>


<TABLE width="210" border=0 cellPadding=2 cellSpacing=0 class="weitab" style="margin-top:10px;">
<tr>
  <td height="32" colspan=2 align="center" background="images/index_png24_a1.png"><font color="#999999"><strong>最近流覽過的商品</strong></font></td>
</tr>
<%
haveshow = request.cookies("haveshow")
'response.write haveshow
if haveshow<>"" then
	arr=split(haveshow,",")
	
	for j=0 to Ubound(arr)
	if isnumeric(arr(j)) then
	
		sql="select pkid,productname,smallpicpath,hit,price1,price"&session("customkind")&" from e_product where pkid=" & arr(j)
		set rsv=server.CreateObject("adodb.recordset")
		rsv.open sql,conn,1,1
		if rsv.bof or rsv.eof then
		else
				set pkid=rsv("pkid")
				set productname=rsv("productname")
				set smallpicpath=rsv("smallpicpath")
				set hit=rsv("hit")
				set price1=rsv("price1")
				set price2=rsv("price"&session("customkind"))
%>
		  <TR> 
			<TD align=center vAlign=top width='25%'><A href="show.asp?pkid=<%=pkid%>"><IMG src="product/<%=smallpicpath%>" width='45' border="0" style="margin:2px;"></A></TD>
			<TD width='75%'>
				<A href="show.asp?pkid=<%=pkid%>" target=_blank><%=productname%></A>
				<DIV align="left" style="padding-bottom:3px;"><b>＄<%=price2%></b>&nbsp;&nbsp;<del>＄<%=price1%></del></DIV>
			</TD>
		  </TR>
		  <TR> 
			<TD align=middle height=5></TD>
		  </TR>
<%
		end if
		rsv.close
		set rsv=nothing
		if j=8 then
			exit for
		end if
	end if	
	next
	
else
	response.write "<tr><td height='25'>還沒有流覽過的商品。</td></tr>"
end if
%>
</TABLE>




