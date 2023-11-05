<!--#include file="conn.asp"-->
<!--#include file="check2.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>设定栏目</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
<!--
td{font-size:9pt}
body {
	margin-top: 5px;margin-left: 2px;
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
}#tabcss td {
	line-height: 24px;
	padding-left: 10px;
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
-->
</style>

</head>

<body bgcolor="#fcfcfc">

<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0" id="tabcss">
  <tr bgcolor="#BFDFFF">
      
    <td height=30 colspan=9>&nbsp;<b>设定栏目：</b></td> 
  </tr>

		<form name="theForm" method="post" action="sitemenuadmin.asp"  onsubmit="return Juge(this)">
		<tr bgcolor='#Ffffff'><td height=23 align=center>&nbsp;</td>
		<td colspan=8><B>首页 <font color="#FF0000">(时尚版已不在这里定义头部菜单，而是采用了商品分类作为头部菜单及下拉菜单) </font></b></td>
		</tr>
		<tr bgcolor='#E1E1E1'>
			<td height=23 align='center'>NO</td>
			<td>头部栏目</td>
			<td>顺序</td>
			<td>栏目类别</td>
			<td>栏目名称</td>
			<td>指定的URL</td>
			<td>新开窗口</td>
			<td>系统固定</td>
			<td>是否显示</td>
		</tr>
  <%
		
		sql4="select * from e_left order by bigmenu desc,l_num"
		set rs4=server.createobject("adodb.recordset")
		rs4.open sql4,conn,1,1
		if rs4.bof or rs4.eof then
			response.write "还没有栏目！"
		else
		i=1
			do while not rs4.eof
				l_id=rs4("l_id")
				menukind=rs4("menukind")
				l_title=rs4("l_title")
				l_num=rs4("l_num")
				bigmenu=rs4("bigmenu")
				zhidingurl=rs4("zhidingurl")
				mylock=rs4("mylock")
				showflag=rs4("showflag")
				newopen=rs4("newopen")
				
				if bigmenu="是" then
%>
		  <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFc4CA';" onMouseOut="this.style.background='#ffffff';"> 
		 <%else%>
		 <tr bgcolor='#BbEcB9' onMouseOver="this.style.background='#FFc4CA';" onMouseOut="this.style.background='#BbEcB9';"> 
		 <%end if%>
			<td height='23' align='center'><%=i%><input type=hidden name='l_id<%=i%>' value="<%=l_id%>"><input type=hidden name='mylock<%=i%>' value="<%=mylock%>"></td>
			<td>
			<select name="bigmenu<%=i%>">
				<option value="<%=bigmenu%>"><%=bigmenu%></option>
				<option value=""></option>
				<option value="是">是</option>
				<option value="否">否</option>
			</select>
			</td>
			<td>
			<input type=text name='l_num<%=i%>' value="<%=l_num%>" size="4">
			</td>			
			<td>
			<select name="menukind<%=i%>" style='width:150px;'>
				<option value="<%=menukind%>"><%=menukind%></option>
				<option value=""></option>
				<option value="产品列表">产品列表</option>
				<option value="普通文章列表">普通文章列表</option>
				<option value="文章列表并指定链接">文章列表并指定链接</option>
				<option value="单个说明">单个说明</option>
				<option value="指定链接">指定链接</option>
			</select>
			</td>
			
			<td>
			<input type=text name='l_title<%=i%>' value="<%=l_title%>" size="15">
			</td>
			
			<td>
			<input type=text name='zhidingurl<%=i%>' value="<%=zhidingurl%>" size="40" <%if mylock="是" then response.write " readonly style='BACKGROUND-COLOR: #FDEEDB;'"%>>
			</td>
			<td>
			<select name="newopen<%=i%>">
				<option value="<%=newopen%>"><%=newopen%></option>
				<option value="是">是</option>
				<option value="否">否</option>
			</select>
			</td>
			
			<td align='center'><%=mylock%></td>
			<td>
			<select name="showflag<%=i%>">
				<option value="<%=showflag%>"><%=showflag%></option>
				<option value="是">是</option>
				<option value="否">否</option>
			</select>
			</td>
		  </tr>

  <%		rs4.movenext
  			i=i+1
  			loop			
		end if
rs4.close
set rs4=nothing		
%>

		  <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
			<td height='23' align='center'><%=i%><input type=hidden name='l_id<%=i%>' value=""></td>
			<td>
			<select name="bigmenu<%=i%>">
				<option value=""></option>
				<option value="是">是</option>
				<option value="否">否</option>
			</select>
			</td>
			<td>
			<input type=text name='l_num<%=i%>' value="" size="4">
			</td>			
			<td>
			<select name="menukind<%=i%>" style='width:150px;'>
				<option value=""></option>
				<option value="产品列表">产品列表</option>
				<option value="普通文章列表">普通文章列表</option>
				<option value="文章列表并指定链接">文章列表并指定链接</option>
				<option value="单个说明">单个说明</option>
				<option value="指定链接">指定链接</option>
			</select>
			</td><td>
			<input type=text name='l_title<%=i%>' value="" size="15">
			</td>
			<td>
			<input type=text name='zhidingurl<%=i%>' value="" size="40">
			</td>
			<td>
			<select name="newopen<%=i%>">
				<option value="<%=newopen%>"><%=newopen%></option>
				<option value="是">是</option>
				<option value="否">否</option>
			</select>
			</td>
			<td>&nbsp;</td>
			<td>
			<select name="showflag<%=i%>">
				<option value="是">是</option>
				<option value="否">否</option>
			</select>
			</td>
		  </tr>
		  
  <tr bgcolor='#Ffffff'> 
      <td height='35' align='right'><input type=hidden name='i' value="<%=i%>"></td>  
	  <td colspan=8><input type="submit" name="Submit2" value=" 保存 ">&nbsp;&nbsp;&nbsp;&nbsp;<input type="reset" name="Submit3" value=" 重置 "> </td>
  </tr>
  
</form>
</table>
<table width="97%" border="0" align="center" cellpadding="2" cellspacing="1">
  <tr bgcolor="#FFFFFF">
    <td height="63" valign="top" style="line-height:170%"><strong><font color="#FF0000">说明：</font></strong><br>
      1 如果设定了“指定的URL”，栏目连接将会以指定的URL优先。否则系统根据“栏目类别”性质生成链接。<br>
      2 如果是“锁定”的栏目，则是后台预先设定的，只可修改“栏目名称”、顺序，而其他属性不可修改。<br>
	  3 如果是“文章列表”及“单个说明”的具体信息内容即可通过菜单“添加栏目信息”来新增信息或通过“栏目信息列表”来修改信息。<br>
      4 如果是“指定链接”类的，即是点击此菜单后跳转到指定的链接。如果不是指向<strong>本站</strong>的链接，请使用http://开头的链接（如：指定的URL中填写http://www.Sunsnu.com/）<br>
      5 在这里新增加的栏目，在左边菜单不显示时，可按F5键刷新后，在栏目管理即可看到新增的菜单。 <br>
      6 如果“是否显示”栏中选择了“否”，则在网店前台菜单中将不显示此栏目。<br>
      7 如果要<strong>删除</strong>非锁定的栏目，下拉框选择为空白，删除输入框的内容，再保存，即可删除此栏目。 <br>
      <br></td>
  </tr>
</table>


<SCRIPT language=JavaScript>
<!--
function Juge()
{


<%for k=1 to i%>

if (document.theForm.l_title<%=k%>.value!="")
{
	if (theForm.menukind<%=k%>.value == "")
	{
	alert("请选择栏目类别。");
	theForm.menukind<%=k%>.focus();
	return (false);
	}
	
	if (theForm.l_num<%=k%>.value == "")
	{
	alert("请填写顺序号。");
	theForm.l_num<%=k%>.focus();
	return (false);
	}
	
	if (theForm.bigmenu<%=k%>.value == "")
	{
	alert("请选择是否为头部栏目。");
	theForm.bigmenu<%=k%>.focus();
	return (false);
	}

	if (theForm.menukind<%=k%>.value == "文章列表并指定链接" || theForm.menukind<%=k%>.value == "指定链接" )
	{
		if (theForm.zhidingurl<%=k%>.value == "")
		{
		alert("请填写指定的链接地址。");
		theForm.zhidingurl<%=k%>.focus();
		return (false);
		}
	}	
	
}

<%next%>




if(confirm("确认要保存栏目设置吗?"))
  return true
else		
  return false;
}
-->
</script>


</body>
</html>
<%
call connclose
%>

