<!--#include file="CONN.asp"-->
<!--#include file="../common/Function1.asp"-->
<!--#include file="System_Function.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>无标题文档</title>
    
<link href="images/skin/Css_<%=Session("Admin_Style_Num")%>/<%=Session("Admin_Style_Num")%>.css" rel="stylesheet" type="text/css"> 
<script type="text/javascript" src="js/jquery.js"></script>
<script type="text/javascript">
//$(document).ready(function() {
//$(".titledaohang").hover(
//	function(){
//		$(this).css('background-image','url(images/skin/Css_<%=Session("Admin_Style_Num")%>/main_men_3.jpg)')
//		}, function(){)
//})
//						   })
</script>

<style type="text/css">
<!--
body {
	margin-left: 4px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;

}
-->
</style></head>

<body class="tree" style="margin:0px;" scroll=yes>
<table width="150" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-top:10px; margin-left:7px;">
  <tr>
    <td align="left"><img src="images/skin/Css_<%=Session("Admin_Style_Num")%>/main_men_top.jpg" width="147" height="55"></td>
  </tr>
</table>
<% if Trim(Session("AdminID"))="52EB650AF9A0F59ECBEF39348ED5E1E1" or Trim(Session("AdminID"))="5EFEA2FE33DB1A7AA433B2318EFE1215" then  %>
<TABLE class=leftframetable cellSpacing=0 cellPadding=0 width=145 align=center 
border=0>
<TBODY>
<TR classid="RecyleManage">
<TD onMouseUp="opencat('18001');" class=titledaohang id=OrdinaryMan 
onmouseover="this.style.backgroundImage='url(images/skin/Css_<%=Session("Admin_Style_Num")%>/main_men_3.jpg)'" 
style="BACKGROUND-IMAGE: url(images/skin/Css_<%=Session("Admin_Style_Num")%>/main_men_1.jpg); CURSOR: hand" 
onmouseout="this.style.backgroundImage='url(images/skin/Css_<%=Session("Admin_Style_Num")%>/main_men_1.jpg)'" 
width=144>参数配置</TD></TR>
<TR id=18001 style="display:none">
<TD colSpan=2>
<TABLE cellSpacing=0 cellPadding=0 width=145 align=center border=0>
<TBODY>
<TR class=RecyleManage parentid="OrdinaryMan" allparentid="OrdinaryMan">
<TD language=javascript id=item$pval[CatID]) style="CURSOR: hand" vAlign=center 
width="97%">
<TABLE cellSpacing=0 cellPadding=0 width="98%" align=center border=0>
<TBODY>
<TR>
<TD align=right width="4%">&nbsp;<IMG src="images/skin/Css_3/dh.gif"> </TD>
<TD align=left width="96%">&nbsp;<A class=lefttop 
href="admin/system.asp" 
target=mainframer>系统配置</A></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=145 align=center border=0>
<TBODY>
<TR class=RecyleManage parentid="OrdinaryMan" allparentid="OrdinaryMan">
<TD language=javascript id=item$pval[CatID]) style="CURSOR: hand" vAlign=center 
width="97%">
<TABLE cellSpacing=0 cellPadding=0 width="98%" align=center border=0>
<TBODY>
<TR>
<TD align=right width="4%">&nbsp;<IMG src="images/skin/Css_3/dh.gif"> </TD>
<TD align=left width="96%">&nbsp;<A class=lefttop 
href="admin/channellist.asp" 
target=mainframer>栏目配置</A></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=145 align=center border=0>
<TBODY>
<TR class=RecyleManage parentid="OrdinaryMan" allparentid="OrdinaryMan">
<TD language=javascript id=item$pval[CatID]) style="CURSOR: hand" vAlign=center 
width="97%">
<TABLE cellSpacing=0 cellPadding=0 width="98%" align=center border=0>
<TBODY>
<TR>
<TD align=right width="4%">&nbsp;<IMG src="images/skin/Css_3/dh.gif"> </TD>
<TD align=left width="96%">&nbsp;<A class=lefttop 
href="QQsever/admin.asp" 
target=mainframer>QQ设置</A></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
</TD></TR></TBODY></TABLE>
<% End If %>
<% ChannelMenuTree 1000 %><%
Sub ChannelMenuTree(MainID)
	Set rss = Server.CreateObject("Adodb.Recordset")
	rss.Open(" Select P.[ID],P.[Title],P.[ChannelModel], "&_
			" (Select Count(*) From [ChannelInfo] Where [ParentID]=P.[ID]) AS [SON] "&_
			" From [ChannelInfo] P Where P.[ParentID]=0 AND P.[Locked]='0' Order by P.[OIndex] ASC "),conn,1,1
	while not rss.eof		 
		MainID = MainID + 100
		if checkPowered(Trim(Session("AdminPower")),rss("id"),5)=true then
			if rss("ChannelModel") = 0 or rss("SON") > 0 then
				%><table width="145" border="0" align="center" cellpadding="0" cellspacing="0" class="leftframetable">
  <tr classid="RecyleManage">
    <td width="144" class="titledaohang"  id="OrdinaryMan" style="CURSOR: hand; "  onmouseup="opencat('<%= MainID %>');" onMouseOut="this.style.backgroundImage='url(images/skin/Css_<%=Session("Admin_Style_Num")%>/main_men_1.jpg)'"
    onMouseOver="this.style.backgroundImage='url(images/skin/Css_<%=Session("Admin_Style_Num")%>/main_men_3.jpg)'"><%=GotTopic(rss("title"),16)%></td>
  </tr>
  <% if rss("SON") > 0 then	 %>
	<% call MunuTree(rss("id"),1,MainID) %>
  <% end if %>
</table><%
			end if
		end if
		rss.Movenext
	wend
	rss.Close
	Set rss = Nothing 
End Sub
sub MunuTree(ID,LevelID,MainID)
	if ID > 0 then
	  Set rsss = Server.CreateObject("Adodb.Recordset")
	  Total = Conn.Execute("Select count(ID) AS total from ChannelInfo Where [ParentID]="&ID)(0)
	  rsss.Open(" Select P.[ID],P.[Title],P.[ChannelModel],P.[url],P.[ChannelTree],"&_
				" (Select Count(*) From [ChannelInfo] Where [ParentID]=P.[ID]) AS [SON] "&_
				" From [ChannelInfo] P Where P.[ParentID]="&ID&" AND P.[Locked]='0' Order by P.[OIndex] ASC "),conn,1,1
	  NMainID = MainID * 1000
	  Num=1
	  %> <tr style="display:none" id="<%= MainID %>">
    <td colspan="2"><%
	  while not rsss.eof
	  TJ=ubound(split(rsss(4),","))
	SpaceStr = ""
		For k = 1 To TJ
		  If k = 1 And k <> TJ Then
		  SpaceStr = SpaceStr &"<img src='menu/line.gif'>"
		  ElseIf k = TJ Then
			If Num = Total Then
				if rsss("SON") = 0 then
				 SpaceStr = SpaceStr &"<img src='images/skin/Css_"&Session("Admin_Style_Num")&"/dh.gif'> "
				 else
				 SpaceStr = SpaceStr &"<img src='menu/plusbottom.gif' class='"&NMainID+1&"'> "
				 end if
			Else
				if rsss("SON") = 0 then
				 SpaceStr = SpaceStr &"<img src='images/skin/Css_"&Session("Admin_Style_Num")&"/dh.gif'> "
				 else
				 SpaceStr = SpaceStr &"<img src='menu/plus.gif' class='"&NMainID+1&"'> "
				 end if
			End If
		  Else
		   SpaceStr = SpaceStr &"<img src='menu/line.gif'>"
		  End If
		Next
		  if checkPowered(Trim(Session("AdminPower")),rsss("id"),0)=true then		 
			  NMainID = NMainID + 1
			  
			  if rsss("ChannelModel") = 0 or rsss("SON") > 0 then
				  %>
     <table width="145" border="0" align="center" cellpadding="0" cellspacing="0" class="<%= MainID %>">
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" class="RecyleManage">
          <td width="87%"  valign="middle" style="CURSOR: hand"  onMouseUp="opencat('<%= NMainID %>');"  language=javascript><table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td align="right" width="4%">&nbsp;<%= SpaceStr %></td>
    <td width="96%" align="left">&nbsp;<%=GotTopic(rsss("title"),16)%></td>
  </tr>
</table>
</td>
        </tr>
		<% if rsss("SON") > 0 then	
					call MunuTree(rsss("id"),LevelID+1,NMainID)
end if%>
  </table><%
			  else
				  ModelLinkUrl = getModelLink(rsss("ChannelModel")) & rsss("ID")&"&cid="&ID
				  if rsss("ChannelModel")=4 then
				  	if instr(rsss("url"),"?")>0 then
				  	 ModelLinkUrl=rsss("url")&"&ChannelID="& rsss("ID")&"&cid="&ID
					 else
					 ModelLinkUrl=rsss("url")&"?ChannelID="& rsss("ID")&"&cid="&ID
					end if
				  end if
				  %>
       <table width="145" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" class="RecyleManage">
          <td width="97%" valign="middle"   id=item$pval[CatID]) style="CURSOR: hand"  language=javascript><table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td align="right" width="4%">&nbsp;<%= SpaceStr %></td>
    <td width="96%" align="left">&nbsp;<a href="<%= ModelLinkUrl %>" class="lefttop" target="mainframer" title="<%= rsss("title") %>"><%=GotTopic(rsss("title"),16)%></a></td>
  </tr>
</table></td>
        </tr>
      </table>
<%
			  end if
			  	
		  end if			
		  rsss.Movenext
		  Num=Num+1
	  wend
	  %></td>
  </tr><%
	  rsss.Close
	  Set rsss = Nothing
  end if	
end sub
%>
<TABLE class=leftframetable cellSpacing=0 cellPadding=0 width=145 align=center 
border=0>
<TBODY>
<TR classid="RecyleManage">
<TD onMouseUp="opencat('18000');" class=titledaohang id=OrdinaryMan 
onmouseover="this.style.backgroundImage='url(images/skin/Css_<%=Session("Admin_Style_Num")%>/main_men_3.jpg)'" 
style="BACKGROUND-IMAGE: url(images/skin/Css_<%=Session("Admin_Style_Num")%>/main_men_1.jpg); CURSOR: hand" 
onmouseout="this.style.backgroundImage='url(images/skin/Css_<%=Session("Admin_Style_Num")%>/main_men_1.jpg)'" 
width=144>访问统计</TD></TR>
<TR id=18000 style="display:none">
<TD colSpan=2>
<TABLE cellSpacing=0 cellPadding=0 width=145 align=center border=0>
<TBODY>
<TR class=RecyleManage parentid="OrdinaryMan" allparentid="OrdinaryMan">
<TD language=javascript id=item$pval[CatID]) style="CURSOR: hand" vAlign=center 
width="97%">
<TABLE cellSpacing=0 cellPadding=0 width="98%" align=center border=0>
<TBODY>
<TR>
<TD align=right width="4%">&nbsp;<IMG src="images/skin/Css_3/dh.gif"> </TD>
<TD align=left width="96%">&nbsp;<A class=lefttop 
href="../SmartStat/Admin_simple.asp?ChannelID=359&amp;cid=358" 
target=mainframer>综合统计</A></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=145 align=center border=0>
<TBODY>
<TR class=RecyleManage parentid="OrdinaryMan" allparentid="OrdinaryMan">
<TD language=javascript id=item$pval[CatID]) style="CURSOR: hand" vAlign=center 
width="97%">
<TABLE cellSpacing=0 cellPadding=0 width="98%" align=center border=0>
<TBODY>
<TR>
<TD align=right width="4%">&nbsp;<IMG src="images/skin/Css_3/dh.gif"> </TD>
<TD align=left width="96%">&nbsp;<A class=lefttop 
href="../SmartStat/Admin_onlineuser.asp?ChannelID=360&amp;cid=358" 
target=mainframer>在线用户</A></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=145 align=center border=0>
<TBODY>
<TR class=RecyleManage parentid="OrdinaryMan" allparentid="OrdinaryMan">
<TD language=javascript id=item$pval[CatID]) style="CURSOR: hand" vAlign=center 
width="97%">
<TABLE cellSpacing=0 cellPadding=0 width="98%" align=center border=0>
<TBODY>
<TR>
<TD align=right width="4%">&nbsp;<IMG src="images/skin/Css_3/dh.gif"> </TD>
<TD align=left width="96%">&nbsp;<A class=lefttop 
href="..//SmartStat/Admin_hour.asp?ChannelID=361&amp;cid=358" 
target=mainframer>24小时排行</A></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=145 align=center border=0>
<TBODY>
<TR class=RecyleManage parentid="OrdinaryMan" allparentid="OrdinaryMan">
<TD language=javascript id=item$pval[CatID]) style="CURSOR: hand" vAlign=center 
width="97%">
<TABLE cellSpacing=0 cellPadding=0 width="98%" align=center border=0>
<TBODY>
<TR>
<TD align=right width="4%">&nbsp;<IMG src="images/skin/Css_3/dh.gif"> </TD>
<TD align=left width="96%">&nbsp;<A class=lefttop 
href="..//SmartStat/Admin_from.asp?ChannelID=362&amp;cid=358" 
target=mainframer>来路统计</A></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=145 align=center border=0>
<TBODY>
<TR class=RecyleManage parentid="OrdinaryMan" allparentid="OrdinaryMan">
<TD language=javascript id=item$pval[CatID]) style="CURSOR: hand" vAlign=center 
width="97%">
<TABLE cellSpacing=0 cellPadding=0 width="98%" align=center border=0>
<TBODY>
<TR>
<TD align=right width="4%">&nbsp;<IMG src="images/skin/Css_3/dh.gif"> </TD>
<TD align=left width="96%">&nbsp;<A class=lefttop 
href="../SmartStat/Admin_ip.asp?ChannelID=363&amp;cid=358" 
target=mainframer>IP统计</A></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=145 align=center border=0>
<TBODY>
<TR class=RecyleManage parentid="OrdinaryMan" allparentid="OrdinaryMan">
<TD language=javascript id=item$pval[CatID]) style="CURSOR: hand" vAlign=center 
width="97%">
<TABLE cellSpacing=0 cellPadding=0 width="98%" align=center border=0>
<TBODY>
<TR>
<TD align=right width="4%">&nbsp;<IMG src="images/skin/Css_3/dh.gif"> </TD>
<TD align=left width="96%">&nbsp;<A class=lefttop 
href="../SmartStat/Admin_browser.asp?ChannelID=364&amp;cid=358" 
target=mainframer>浏览器统计</A></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=145 align=center border=0>
<TBODY>
<TR class=RecyleManage parentid="OrdinaryMan" allparentid="OrdinaryMan">
<TD language=javascript id=item$pval[CatID]) style="CURSOR: hand" vAlign=center 
width="97%">
<TABLE cellSpacing=0 cellPadding=0 width="98%" align=center border=0>
<TBODY>
<TR>
<TD align=right width="4%">&nbsp;<IMG src="images/skin/Css_3/dh.gif"> </TD>
<TD align=left width="96%">&nbsp;<A class=lefttop 
href="../SmartStat/Admin_keyword.asp?ChannelID=365&amp;cid=358" 
target=mainframer>关键词统计</A></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=145 align=center border=0>
<TBODY>
<TR class=RecyleManage parentid="OrdinaryMan" allparentid="OrdinaryMan">
<TD language=javascript id=item$pval[CatID]) style="CURSOR: hand" vAlign=center 
width="97%">
<TABLE cellSpacing=0 cellPadding=0 width="98%" align=center border=0>
<TBODY>
<TR>
<TD align=right width="4%">&nbsp;<IMG src="images/skin/Css_3/dh.gif"> </TD>
<TD align=left width="96%">&nbsp;<A class=lefttop 
href="../SmartStat/Admin_operating.asp?ChannelID=366&amp;cid=358" 
target=mainframer>操作系统</A></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=145 align=center border=0>
<TBODY>
<TR class=RecyleManage parentid="OrdinaryMan" allparentid="OrdinaryMan">
<TD language=javascript id=item$pval[CatID]) style="CURSOR: hand" vAlign=center 
width="97%">
<TABLE cellSpacing=0 cellPadding=0 width="98%" align=center border=0>
<TBODY>
<TR>
<TD align=right width="4%">&nbsp;<IMG src="images/skin/Css_3/dh.gif"> </TD>
<TD align=left width="96%">&nbsp;<A class=lefttop 
href="../SmartStat/Admin_page.asp?ChannelID=367&amp;cid=358" 
target=mainframer>页面统计</A></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>


<table width="145" border="0" align="center" cellpadding="0" cellspacing="0" class="leftframetable">
  <tr allparentid="OrdinaryMan" parentid="OrdinaryMan" class="RecyleManage"> 
    <td width="98%" class="titledaohang">系统信息</td>
  </tr>
  </table>
  <table width="145" border="0" align="center" cellpadding="0" cellspacing="0" class="Manage" style="margin-left:4px;">
    <tr>
      <td height="50" valign="top" style="padding:10px;">&nbsp;</td>
    </tr>
    <tr>
      <td><img src="images/skin/Css_<%=Session("Admin_Style_Num")%>/main_men_foot.jpg" width="145" height="18"></td>
    </tr>
  </table>
<script language="JavaScript" type="text/JavaScript">
$(document).ready(function(){})
function opencat(cat)
{
	var cat1=cat
	var cat=document.getElementById(cat)
	  if(cat.style.display=="none")
	  {
	  $("."+cat1+"").attr("src","menu/minus.gif")
		 cat.style.display="";
	  }
	  else
	  {
		 cat.style.display="none"; 
		 $("."+cat1+"").attr("src","menu/plus.gif")
	  }
}
</script>
<%Call CloseConn()%> 
</body>
</html>
