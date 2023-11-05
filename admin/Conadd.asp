<!--#include file="conn.asp"-->
<!--#include file="check2.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>增加内容</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../inc/i.css" type=text/css rel=stylesheet>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>

<script language="JavaScript">
<!--
function Juge()
{
  if (theForm.c_parent.value == "")
  {
    alert("请选择栏目名称!");
    document.theForm.c_parent.focus();
    return false;
  }
  if (theForm.c_num.value == "")
  {
    alert("请用数字填写顺序!");
    document.theForm.c_num.focus();
    return false;
  }
  if (isNaN(theForm.c_num.value))
  {
    alert("请用数字填写顺序!");
    document.theForm.c_num.focus();
    return false;
  }
}
//-->
</script>
 <%


Dim htmlData

htmlData = Request.Form("c_contect")

Function htmlspecialchars(str)
	str = Replace(str, "&", "&amp;")
	str = Replace(str, "<", "&lt;")
	str = Replace(str, ">", "&gt;")
	str = Replace(str, """", "&quot;")
	htmlspecialchars = str
End Function
%>
<script type="text/javascript" charset="utf-8" src="../kindeditor.js"></script>
	<script type="text/javascript">
		KE.show({
			id : 'c_contect',
			urlType : 'absolute',
			imageUploadJson : '../../asp/upload_json.asp',
			fileManagerJson : '../../asp/file_manager_json.asp',
			allowFileManager : true
			 
		});
	</script>
<style type="text/css">
<!--
body {
	margin-top: 5px;
}
.style1 {color: #CCCCCC}
.style2 {color: #FF0000}
.style4 {
	font-size: 14px;
	font-weight: bold;
}
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
<body bgcolor="#fcfcfc">


<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0" id="tabcss">
  <form name="theForm" method="post" action="conaddadmin.asp?newskind=<%=request("newskind")%>" onSubmit="return Juge()">
    <tr bgcolor="#BFDFFF"> 
      <td height="33" colspan="2"><strong>增加信息发布</strong></td>
    </tr>
    <tr> 
      <td><div align="right">栏目</div></td>
      <td > <select name="c_parent" size="1" style="width:250px;">
        
          <%
		set rs2=server.CreateObject("adodb.recordset")
		sql2="select * from e_board"
		rs2.open sql2,conn,1,1
		if rs2.bof or rs2.eof then
		else
		do while not rs2.eof
		id=rs2("id")
		title=rs2("title")
		'response.write "<option value=''>"&title&"</option>"

				set rs3=server.CreateObject("adodb.recordset")
				sql3="select * from e_left where (menukind='普通文章列表' or menukind='文章列表并指定链接' or menukind='单个说明')  order by bigmenu desc,l_num "
				rs3.open sql3,conn,1,1
				if rs3.bof or rs3.eof then
				else
				do while not rs3.eof
				l_id=rs3("l_id")
				l_title=rs3("l_title")
				menukind=rs3("menukind")
				if request("newskind")=cstr(l_id) then
				response.write "<option value='"&id&"|"&l_id&"' selected>--"&l_title&" </option>"
				else
				response.write "<option value='"&id&"|"&l_id&"'>--"&l_title&" </option>"
				end if
				rs3.movenext
				loop
				end if
				rs3.close
				set rs3=nothing

		rs2.movenext
		loop
		end if
		rs2.close
		set rs2=nothing
%>
        </select></td>
    </tr>
    <tr> 
      <td height="33"><div align="right">标题</div></td>
      <td><input name="c_title" type="text" size="60" maxlength="80">
        （80字符以内）</td>
    </tr>
    <tr> 
      <td height="24">&nbsp;</td>
      <td>标题颜色设置 如：&lt;font color=&quot;#FF0000&quot;&gt;标题&lt;/font&gt;</td>
    </tr>
    <tr> 
      <td height="24"><div align="right">内容</div></td>
      <td><p> 
        
           <textarea id="c_contect" name="c_contect" cols="100" rows="8" style="width:725px;height:400px;visibility:hidden;"></textarea>
        </p></td>
    </tr>
    <tr> 
      <td height="28"><div align="right">顺序</div></td>
      <td><input name="c_num" type="text" value="0" size="5" maxlength="5"> </td>
    </tr>
    <tr> 
      <td height="28"><div align="right">促销信息</div></td>
      <td><input type="radio" name="cuxflag" value="1" <%if cuxflag="1" then response.write "checked"%>>
        是　 
        <input type="radio" name="cuxflag" value="0" <%if cuxflag<>"1" then response.write "checked"%>>
        否 </td>
    </tr>

    <!--
	 <tr>
      <td height="33"><div align="right"></div></td>
      <td><span class="style2">内容情况二 (如果外链接，请填写链接地址) </span></td>
	 </tr>
	<tr>
      <td height="33"><div align="right">链接地址</div></td>
      <td><input name="url" type="text" id="url" size="60" maxlength="100"></td>
    </tr>
	<tr>
	  <td height="33">&nbsp;</td>
	  <td><span class="style2">内容情况三</span></td>
    </tr>
	<tr>
	  <td height="33">&nbsp;</td>
	  <td><a href="condown.asp" class="style3">增加下载页</a><a href="condown.asp" class="style3">内容</a></td>
    </tr>
-->
    <tr> 
      <td height="45"><div align="right"></div></td>
      <td><input type="hidden" name="c_ubb" value="1"> <input type="submit" name="Submit" value=" 提 交 "> </td>
    </tr>
  </form>
</table>
<!--
<a class="style1" style="Cursor:hand" onClick="MM_openBrWindow('upload.asp','','width=350,height=210')">上传图片</a>
-->
</body>
</html>
<%
conn.close
set conn=nothing
%>

