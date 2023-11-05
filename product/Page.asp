<%
sub page(pagerecord,allrecord,allpage,backandnexturl,addbuttonurl)
%>
<SCRIPT language=JavaScript>
<!--
function jugenum(theFormnum)
{	
	  if (isNaN(theFormnum.page.value))
	  {
		alert("页数要为数字!");
		theFormnum.page.focus();
		return (false);
	  }
}
-->
</script>
<%
'backandnexturl="advsearchadmin.asp?choosedate="&choosedate&"&choosedate2="&choosedate2&""
'backandnexturl="advsearchadmin.asp?none="

  if nowpage="" then nowpage="1"	
  response.Write"<table width='650' border=0 align=center cellpadding=0 cellspacing=0>"
  	response.Write"<tr><td colspan=2 height=5></td></tr>"
    response.Write"<tr>"
      response.Write"<td width='550'> <div align=right>每页"&pagerecord&"条 共"&allrecord&"条 第"&nowpage&"/"&allpage&"页&nbsp;"
	  
			call actionbutton(request("menuid"),"3",addbuttonurl)
			call actionbutton(request("menuid"),"4","")
			
			response.Write"&nbsp;&nbsp;<input type=button name=Button3 value='首页' onClick=window.location.href='"&backandnexturl&"&page=1'>"
			if nowpage>1 then
			response.Write"&nbsp;<input type=button name=Button4 value='上页' onClick=window.location.href='"&backandnexturl&"&page="&nowpage-1&"'>"
			else
			response.Write"&nbsp;<input type=button name=Button4 value='上页'>"
			end if
			if clng(nowpage)<clng(allpage) then
			response.Write"&nbsp;<input type=button name=Button5 value='下页' onClick=window.location.href='"&backandnexturl&"&page="&nowpage+1&"'>"
			else
			response.Write"&nbsp;<input type=button name=Button5 value='下页'>"
			end if
			
			response.Write"&nbsp;<input type=button name=Button6 value='末页' onClick=window.location.href='"&backandnexturl&"&page="&allpage&"'>&nbsp;"
			response.Write"</td></form>"
			response.Write"<form name='theFormnum' method='post' action='"&backandnexturl&"' onSubmit='return jugenum(this)'>"
			response.Write"<td width='100'><div align=left>"
			'response.Write"页数<input name='page' type='text' size='4' maxlength='8' value='"&nowpage&"'>"
			'response.Write"<input name='pagego' type='SUBMIT' value='GO'>"
			response.Write"&nbsp;&nbsp;<a href='javascript:window.scroll(0,0)'>TOP↑</a>&nbsp;"
			response.Write"</div></div></td></form>"
	response.Write"</tr>"
  response.Write"</table>"
  
response.Cookies("jumppage")=backandnexturl&"&page="&nowpage
end sub
%>
