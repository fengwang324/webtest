<!--#include file="conn.asp"-->
<!--#include file="check2.asp"-->
<%
s_sql="select s_colorsize from siteshow"
set s_rs=server.createobject("adodb.recordset")
s_rs.open s_sql,conn,1,1
if s_rs.bof or s_rs.eof then
	
else
	s_colorsize=s_rs("s_colorsize")
end if
s_rs.close
set s_rs=nothing

on error resume next
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>快速修改商品属性</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../product/FixTable.css" rel="Stylesheet" type="text/css" />
<style>
<!--
td{font-size:9pt;line-height:140%;}
body {
	margin-top: 3px;margin-left: 3px;
}
.style3 {color: #A84E22; font-weight: bold; }
#tabcss{
	BORDER-COLLAPSE: collapse;
	border-top-width: 1px;
	border-left-width: 1px;
	border-top-style: solid;
	border-left-style: solid;
	border-top-color: #dddddd;
	border-left-color: #dddddd;
}
#tabcss td {
	line-height: 21px;
	padding-left: 3px;
	padding-top: 2px;
	padding-bottom: 2px;
	color: #333333;
	border-right-width: 1px;
	border-bottom-width: 1px;
	border-right-style: solid;
	border-bottom-style: solid;
	border-right-color: #dddddd;
	border-bottom-color: #dddddd;
}
.input1 {
	BACKGROUND-COLOR: #FEF3E7;BORDER-RIGHT: #FF0000 1px groove; BORDER-TOP: #FF0000 1px groove
}
form{margin:1px;}
-->
</style>

<SCRIPT LANGUAGE="JavaScript"> 
<!-- 
function allcheck() 
{
	for(i=document.theForm.elements.length-1;i>0;i--)
	{
		document.theForm.elements[i-1].checked=true;
	}
}

function nocheck() 	
{
	for(i=document.theForm.elements.length-1;i>0;i--)
	{
		document.theForm.elements[i-1].checked=false;
	}
}

   // --> 
</SCRIPT> 
<script language="javascript">
<!--
function SetBgColor(Menu,Menucolor)
{
	if (Menu.style.background.toUpperCase()!="#FFDDBB")
	{
	Menu.style.background=Menucolor;
	}
}
function RestoreBgColor(Menu,Menucolor)
{
	if (Menu.style.background.toUpperCase()!="#FFDDBB")
	{
	Menu.style.background=Menucolor;
	}
}

function ClickBgColor(Menu)
{
	if (Menu.style.background.toUpperCase()!="#FFDDBB")
	{
	Menu.style.background="#FFDDBB";
	}
	else
	{
	Menu.style.background="#E7E7E7";
	}
}

-->
</script>
</head>

<body bgcolor="#fcfcfc" onLoad="tab.style.height=window.screen.height-258">

<table border="0" width="100%" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <form name="form6" method="post" action="pl_modify.asp">
      <td height=30 colspan=15> <font color=#0000ff><strong></strong></font>查找： 
        <select name="selectkind" style="width:120px;">
          <option value="productname">商品名称</option>
		  <option value="kindname">种类</option>
		  <option value="model">货号</option>
          <option value="pipai">品牌</option>
		  <option value="color">颜色</option>
		  <option value="psize">尺寸</option>
          <option value="disc">商品详细</option>
        </select> <input name="keyword" type="text" size="15" maxlength="15"> 
        <input type="submit" name="Submit" value="查 找">&nbsp;&nbsp;&nbsp;&nbsp;<a href=pl_modify.asp>全部</a>&nbsp;&nbsp;&nbsp;&nbsp;
		(温馨提示：如窗口较小，操作不便，按F11键最大化窗口)
		</td>
    </form>
  </tr>

<%
i=1
nowpage=request("page")
if nowpage="" then nowpage="1"
selectkind=trim(request("selectkind"))
keyword=trim(request("keyword"))
%>
</table>

<form name="theForm" method="post">

<div id=tab class="fixbox">
<table  class="fixtable" width="1450" border="0" align="center" cellpadding="1" cellspacing="0">

  <tr> 
    <th class="lockcolumn" height="24">No.</th>
	<th class="lockcolumn">图片</th>
    <th class="lockcolumn">商品名称</th>
	<th class="lockcolumn" bgcolor="#cca" width="1"></th>
	
    <th class="thbg">种类</th>
    <th class="thbg">货号</th>
    <th class="thbg">品牌</th>
	<%if s_colorsize="是" then%>
    <%end if%>
    <th class="thbg">市场价</th>
	<th class="thbg">会员价</th>
	<th class="thbg">铜级价</th>
	<th class="thbg">银级价</th>
	<th class="thbg">金级价</th>
	<th class="thbg">钻级价</th>
	<th class="thbg">重量(克)</th>
	<th class="thbg">库存情况</th>
	<th class="thbg">点击数</th>
	<th class="thbg">销售数</th>
  </tr>
  <%

if keyword="" then
	sql="select pkid,kind,kindname,model,productname,smallpicpath,pipai,color,psize,price1,price2,price3,price4,price5,price6,oneweight,kuchushu,hit,saleshu from view_product order by pkid desc"
else
	sql="select pkid,kind,kindname,model,productname,smallpicpath,pipai,color,psize,price1,price2,price3,price4,price5,price6,oneweight,kuchushu,hit,saleshu from view_product where "&selectkind&" like '%"&keyword&"%' order by pkid desc"
end if


set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof or rs.eof then
	response.write "<div align=center><br>没有商品记录!</div>"
else
rs.pagesize=30
rs.absolutepage=1
if request("page")<>"" then rs.absolutepage=request("page")
rowcount=rs.pagesize
	do while not rs.eof and rowcount>0
	
		pkid=rs("pkid")
		kind=rs("kind")
		kindname=rs("kindname")
		
		model=rs("model")
		productname=rs("productname")
		smallpicpath=rs("smallpicpath")
		pipai=rs("pipai")
		color=rs("color")
		psize=rs("psize")
		price1=rs("price1")
		price2=rs("price2")
		price3=rs("price3")
		price4=rs("price4")
		price5=rs("price5")
		price6=rs("price6")
		oneweight=rs("oneweight")
		kuchushu=rs("kuchushu")
		hit=rs("hit")
		saleshu=rs("saleshu")
		
		if color<>"" then
		color=replace(color,"，",",")
		end if
		if psize<>"" then
		psize=replace(psize,"，",",")
		end if
		
		if kindname<>"" then
			grade=len(kind)/4
			'p=1
			gradestr=""
			'for p=1 to grade-1
				gradestr=gradestr & String(grade-1, "—")
			'next
			kindname="—"&gradestr&kindname
		end if
'response.write hit &"yyy"
'response.end
%>
  <tr Onmouseover="return SetBgColor(this,'#E1E1E1');" Onmouseout="return RestoreBgColor(this,'#FFFFFF');" Onclick="return ClickBgColor(this);"> 
   <td class="lockcolumn" height="22" align="center"><%=(nowpage-1)*30+i%><input type='checkbox' name='checkbox<%=i%>' value='<%=pkid%>'></td>
   <TD class="lockcolumn"><a href="../show.asp?pkid=<%=pkid%>" target="_blank"><img src="../product/<%=smallpicpath%>" height="35" border="0"></a></TD>
   <td class="lockcolumn"><textarea name="productname<%=i%>" cols="24" rows="2"><%=productname%></textarea></td>
   <td class="lockcolumn" bgcolor="#CCCCCC" width="1"></td>
   <td><input type='hidden' name='pkid<%=i%>' value='<%=pkid%>'><input type='hidden' name='kind<%=i%>' value='<%=kind%>'><input type='input' name='kindname<%=i%>' value='<%=kindname%>'  
	size='15' readonly class='input1' title='点击后面…选择分类'><input type='button' name='butt<%=i%>' value='…' title='点击这里选择分类' onClick="choosekind('<%=i%>')"></td>
    <td><input type='input' name='model<%=i%>' value='<%=model%>' size='8'></td>
	
	<td>
        <select name='pipai<%=i%>' style="z-index:2; width:90px;">
          <option value="<%=pipai%>" selected><%=pipai%></option>
		  <%
			response.write "<option value=''></option>"
			c_sql="select memo from x_pinpai order by memo"
			set c_rs=conn.execute(c_sql)
			if not c_rs.eof then
			do while not c_rs.eof
				memo2=c_rs("memo")
				response.write "<option value='"&memo2&"'>"&memo2&"</option>"
				c_rs.movenext
			loop
			end if
			c_rs.close
			set c_rs=nothing
		  %>
        </select></td>
	
	
	
	<%if s_colorsize="是" then%>
	<%end if%>
	<td><input type='input' name='price1_<%=i%>' value='<%=cdbl(price1)%>' size='5'></td>
	<td><input type='input' name='price2_<%=i%>' value='<%=cdbl(price2)%>' size='5'></td>
	<td><input type='input' name='price3_<%=i%>' value='<%=cdbl(price3)%>' size='5'></td>
	<td><input type='input' name='price4_<%=i%>' value='<%=cdbl(price4)%>' size='5'></td>
	<td><input type='input' name='price5_<%=i%>' value='<%=cdbl(price5)%>' size='5'></td>
	<td><input type='input' name='price6_<%=i%>' value='<%=cdbl(price6)%>' size='5'></td>
	<td><input type='input' name='oneweight<%=i%>' value='<%=oneweight%>' size='5'></td>
	<td><input type='input' name='kuchushu<%=i%>' value='<%=kuchushu%>' size='10'></td>
	<td><input type='input' name='hit<%=i%>' value='<%=hit%>' size='5'></td>
	<td><input type='input' name='saleshu<%=i%>' value='<%=saleshu%>' size='5'></td>
  </tr>
<%
	rs.movenext
	rowcount=rowcount-1
	i=i+1
	loop
end if
%>
</table>
</div>

<%
response.write "<table border=0 width='100%' align=center cellpadding='2' cellspacing='0'><tr>"
response.write "<td width='37%' height=30>"
response.write "<input type='button' name='check1' value='全选' class='inputss' onclick='return allcheck()'> "
response.write "<input type='button' name='check2' value='不选' class='inputss' onclick='return nocheck()'>&nbsp;"
response.write "<select name='allkind' style='width:153px; font-size:12px;'>"
	dim rsk,sql,maxkind,kindnum,kindname,grade,i,gradestr
	response.write "<option value=''>将选中的批量分类到...</option>"
		sql="select * from sh_kind order by kindnum"
		set rsk=server.createobject("adodb.recordset")
		rsk.open sql,conn,1,1
		if rsk.bof or rsk.eof then
		else
		do while not rsk.eof
			kindnum=rsk("kindnum")
			kindname=rsk("kindname")
			grade=len(kindnum)/4
			'p=1
			gradestr=""
			'for p=1 to grade-1
				gradestr=gradestr & String(grade-1, "—")
			'next
			response.write "<option value='"&kindnum&"'>|—"&gradestr&kindname&"</option>"
		rsk.movenext
		loop
		end if
		rsk.close
		set rsk=nothing
		response.write "</select> <input type='button' name='Submit1' value='确定' onClick='savekind()'></td>"
		response.write "<td align=right>"
		response.write "<input type='button' name='Submit2' value='保存修改' onClick='saveall()'></td></form>"
		
		response.write "<td align=right>"
		
		backnexturl="keyword="&keyword&"&selectkind="&selectkind&"&none=&"
		if nowpage="" then nowpage="1"
		lastpage=rs.pagecount
		nowpath=request.ServerVariables("SCRIPT_NAME")
		
		response.write "共"&rs.recordcount&"条 30条/页 "&nowpage&"/"&rs.pagecount &"页 "

		response.write "<input type='button' name='Button1' value='首页' onclick=window.location.href='"&nowpath&"?"&backnexturl&"page=1'>&nbsp;"
		if nowpage>1 then
			response.write"<input type='button' name='Button2' value=上页 onclick=window.location.href='"&nowpath&"?"&backnexturl&"page="&nowpage-1&"'>&nbsp;"
		else
			response.write"<input type='button' name='Button3' value=上页>&nbsp;"
		end if
		if clng(nowpage) < clng(lastpage) then
			response.write"<input type='button' name='Button4' value=下页 onclick=window.location.href='"&nowpath&"?"&backnexturl&"page="&nowpage+1&"'>&nbsp;"
		else
			response.write"<input type='button' name='Button5' value=下页>&nbsp;"
		end if
		response.write"<input type='button' name='Button6' value=尾页 onclick=window.location.href='"&nowpath&"?"&backnexturl&"page="&clng(lastpage)&"'>&nbsp;"
		response.write "</td>"
		
		response.write "<form name='form1' method='post' action="&nowpath&"?"&backnexturl&">"
		response.write "<td>第<input type=text name='page' size=3>页<input type='submit' name='Submit' value='确定'>"
		response.write "</td></form>"
response.write "</tr></table>"


%>

<SCRIPT LANGUAGE="JavaScript">
function savekind()
{

	if (theForm.allkind.value == "")
	{
		alert("请选择批量分类的类别!");
		document.theForm.allkind.focus();
		return false;
	}
	var getp="0";
	
	for(p=1;p<<%=i%>;p++)
	{
		if (theForm.all("checkbox"+p).checked==true)
		{
			document.theForm.all("kind"+p).value=document.theForm.allkind.value;
			document.theForm.all("kindname"+p).value=document.theForm.allkind.options[theForm.allkind.selectedIndex].text;
			theForm.all("checkbox"+p).checked=false;
			getp="1";
		}
	}
	if (getp=="0")
	{
		alert("还没有选中任何商品！请在商品前面打勾，然后选择批量类别，再点击"确定"。");
		return false;
	}

}

function saveall()
{
	for(p=1;p<<%=i%>;p++)
	{
		if (theForm.all("kind"+p).value == "")
		{
			alert("请选择第 "+p+" 行类别!");
			document.theForm.all("kindname"+p).focus();
			return false;
		}
	}

	if(confirm("确认保存修改吗？"))
	{
		document.theForm.action = 'pl_modify_admin.asp?page=<%=nowpage%>&selectkind=<%=selectkind%>&keyword=<%=keyword%>&no=';
		document.theForm.submit();
	}
	else
	return false;  
}
</SCRIPT>
<SCRIPT>
<!--
function choosekind(myi) 
{ 
  var k;
  k=window.open("openkind.asp?i="+myi+"&mykind="+theForm.all("kind"+myi).value,"choosek","top=50,left=450,width=500,height=600,scrollbars=yes");
  k.focus();
} 
// -->
</SCRIPT>

</body>
</html>
<%
rs.close
set rs=nothing
call connclose
%>


