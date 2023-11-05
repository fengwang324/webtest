<!-- #include file="conn.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>礼品资料</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
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
  if (theForm.smallpicpath.value == "")
  {
    alert("请上传礼品小图片");
    return false;
  }
  if (theForm.bigpicpath.value == "")
  {
    alert("请上传礼品大图片");
    return false;
  }
  if (theForm.lipinno.value == "")
  {
    alert("请输入礼品编号.");
    document.theForm.lipinno.focus();
    return false;
  }
  if (theForm.lipinname.value == "")
  {
    alert("请输入礼品名称.");
    document.theForm.lipinname.focus();
    return false;
  }
  if (theForm.zhifen.value == "")
  {
    alert("请输入所需积分.");
    document.theForm.zhifen.focus();
    return false;
  }
  if (isNaN(theForm.zhifen.value))
  {
    alert("所需积分要为数字!");
    theForm.zhifen.focus();
    return (false);
  }

  if (theForm.num.value == "")
  {
    alert("请输入顺序.");
    document.theForm.num.focus();
    return false;
  }
  if (isNaN(theForm.num.value))
  {
    alert("顺序要为数字!");
    theForm.num.focus();
    return (false);
  }

}
//-->
</script>
<%

Dim htmlData

htmlData = Request.Form("memo")

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
			id : 'memo',
			urlType : 'absolute',
			imageUploadJson : '../../asp/upload_json.asp',
			fileManagerJson : '../../asp/file_manager_json.asp',
			allowFileManager : true
			 
		});
	</script>
<style>
td { font-size: 9pt; color: #000000;font-family: "宋体"}

.border {
	BORDER-BOTTOM: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid
}
input {
      	height=20px;BORDER-RIGHT: #999999 1px solid; BORDER-TOP: #999999 1px solid; FONT-SIZE: 9pt; BORDER-LEFT: #999999 1px solid; BORDER-BOTTOM: #999999 1px solid; FONT-FAMILY: "宋体"; BACKGROUND-COLOR: #ffffff
}
.style1 {
	font-weight: bold;
	color: #FF0000;
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

	padding-bottom: 2px;
	color: #333333;
	border-right-width: 1px;
	border-bottom-width: 1px;
	border-right-style: solid;
	border-bottom-style: solid;
	border-right-color: #dddddd;
	border-bottom-color: #dddddd;
}
</style>

</head>

<body bgcolor="#fcfcfc">
<%
id=request("id")
if id="" then
	showflag="1"
else
	sql="select * from x_lipin where id="&id&" "
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.bof or rs.eof then
		response.write "<div align=center>没有记录! </div>"
	else
		id=rs("id")
		picpath=rs("picpath")
		picpath2=rs("picpath2")
		lipinno=rs("lipinno")
		lipinname=rs("lipinname")
		zhifen=rs("zhifen")
		showflag=rs("showflag")
		num=rs("num")
		memo=rs("memo")
	end if
	rs.close 
	set rs=nothing
end if

	conn.close
	set conn=nothing
%>
<form name="theForm" method="post" action="lipinaddadmin.asp" onSubmit="return Juge()">

  <table width="97%" border="0" align="center" cellpadding="0" cellspacing="0" id="tabcss">
    <tr> 
      <td width="682" height="33" bgcolor="#E1EFFF"><div class="style1">礼品设置</div></td>
    </tr>
    <tr> 
      <td height="32"> <input name="smallpicpath" type="text" size="40" value="<%=picpath%>"> 
        <input type="button" name="Button" value="上传小图片" onClick="MM_openBrWindow('upload2.asp?size=small','','width=310,height=210,top=100,left=10')">
        (最佳约100 X 约100) </td>
    </tr>
    <tr> 
      <td height="32"> <input name="bigpicpath" type="text" size="40" value="<%=picpath2%>"> 
        <input type="button" name="Button" value="上传大图片" onClick="MM_openBrWindow('upload2.asp?size=big','','width=310,height=210,top=100,left=10')">
      最佳400-600px X 400-600px 100KB以内　图片名为英文或数字      </td>
    </tr>
    <tr> 
      <td height="33">礼品编号: 
        <input type="text" name="lipinno" size="30" value="<%=lipinno%>"></td>
    </tr>
    <tr> 
      <td height="33">礼品名称: 
        <input type="text" name="lipinname" size="30" value="<%=lipinname%>"></td>
    </tr>
    <tr> 
      <td height="33">所需积分: 
        <input name="zhifen" type="text" size="30" value="<%=zhifen%>"></td>
    </tr>
    <tr> 
      <td height="33">是否显示: 
        <input type="checkbox" name="showflag" value="1" <%if showflag="1" then response.write "checked"%>></td>
    </tr>
    <tr> 
      <td height="33">顺序: 
        <input type="text" name="num" size="5" value="<%=num%>">
        用于排列的先后。</td>
    </tr>
    <tr>
      <td height="33">详细: 
	<textarea id="memo" name="memo" cols="100" rows="8" style="width:725px;height:400px;visibility:hidden;"><%=htmlspecialchars(memo)%></textarea>
		 
		</td>
    </tr>
    <tr> 
      <td height="37"> <input type="hidden" name="id" value="<%=id%>" > 
        <%if id="" then%> <input type="submit" name="Submit" value=" 增 加 " > <%else%> <input type="submit" name="Submit" value=" 修 改 " > <%end if%> </tr>
  </table>
  <p>&nbsp;</p>
</form>
</body>
</html>

