<%sub product_list(showkind)%> 
	<%if showkind=1 then%>
		<TD vAlign=top align=center>
		<div align="center" style="width:<%=s_picwidth+8%>px; padding:2px;"  Onmouseover='return SetBgColor(this);' Onmouseout='return RestoreBgColor(this);'>
			<div align="center" style="padding-top:3px;"> 
				<A href="show.asp?pkid=<%=pkid%>" <%=newwindow%>> <IMG src="product/<%=smallpicpath%>" width='<%=s_picwidth%>' height='<%=s_picheight%>' border="0" class=img1></A>
			</div>
	
	
			<DIV align="left" style="MARGIN-TOP: 2px; WIDTH: <%=s_picwidth%>px; height:38px;overflow: hidden;">
				<A class=font_link href="show.asp?pkid=<%=pkid%>" <%=newwindow%> title="<%=productname%>"><%=productname%></A>
			</DIV>
			<DIV align="left"><font class='nowprice'>＄<%=price2%></font>&nbsp;&nbsp;<FONT class='oldprice'><del>＄<%=price1%></del></FONT></DIV>
			<!--
			<DIV align="left" style="MARGIN-TOP: 2px; WIDTH: <%=s_picwidth%>px; height:38px;">上市时间：<%=addtime%></DIV>
			-->
			<DIV align="left" style="MARGIN-TOP: 5px;MARGIN-bottom: 2px; WIDTH: <%=s_picwidth%>px">
			<%if s_colorsize="是" then%>
			<a href="show.asp?pkid=<%=pkid%>" <%=newwindow%>><img src="images/icon_car.gif" border=0></a>
			<%else%>
			<a href="show.asp?pkid=<%=pkid%>&num=1&tk=shop7z" <%=newwindow%>><img src="images/goumai.gif" border=0></a>
			<%end if%>
			</div>
		</div>
		</TD>
	<%end if%>
	<%if showkind=2 then%>
		<TD vAlign=top align=center>
		<div align="center" style="width:<%=s_picwidth+8%>px; padding:2px;"  Onmouseover='return SetBgColor(this);' Onmouseout='return RestoreBgColor(this);'>
			<div align="center" style="padding-top:3px;"> 
				<A href="show.asp?pkid=<%=pkid%>" <%=newwindow%>> <IMG src="product/<%=smallpicpath%>" width='<%=s_picwidth%>' height='<%=s_picheight%>' border="0" class=img1></A>
			</div>
			<%if s_model="是" then%>
			<DIV align="left" style="MARGIN-TOP: 5px; WIDTH: <%=s_picwidth%>px">
			商品號：<a href="show.asp?pkid=<%=pkid%>" <%=newwindow%>><%=model%></a>
			</div>
			<%end if%>
			<%if s_pipai="是" then%>
			<DIV align="left" style="MARGIN-TOP: 2px; WIDTH: <%=s_picwidth%>px">
			品　牌：<%=pipai%>
			</div>	
			<%end if%>
			<DIV align="left" style="MARGIN-TOP: 2px; WIDTH: <%=s_picwidth%>px; height:38px;overflow: hidden;">
				<A class=font_link href="show.asp?pkid=<%=pkid%>" <%=newwindow%> title="<%=productname%>"><%=productname%></A>
			</DIV>
			<DIV align="left"><font class='nowprice'>＄<%=price2%></font>&nbsp;&nbsp;<FONT class='oldprice'><del>＄<%=price1%></del></FONT></DIV>
			<!--
			<DIV align="left" style="MARGIN-TOP: 2px; WIDTH: <%=s_picwidth%>px; height:38px;">上市时间：<%=addtime%></DIV>
			-->
			<DIV align="left" style="MARGIN-TOP: 5px;MARGIN-bottom: 2px; WIDTH: <%=s_picwidth%>px">
			<%if s_colorsize="是" then%>
			<a href="show.asp?pkid=<%=pkid%>" <%=newwindow%>><img src="images/icon_car.gif" border=0></a>
			<%else%>
			<a href="show.asp?pkid=<%=pkid%>&num=1&tk=shop7z" <%=newwindow%>><img src="images/goumai.gif" border=0></a>
			<%end if%>
			</div>
		</div>
		</TD>
	<%end if%>
<%end sub%>
