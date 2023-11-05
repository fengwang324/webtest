<!--#include file="conn.asp"-->
<!--#include file="check2.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>设置版权信息</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../inc/i.css" type=text/css rel=stylesheet>
 <%

Dim htmlData

htmlData = Request.Form("copyright")

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
			id : 'copyright',
			urlType : 'absolute',
			imageUploadJson : '../../asp/upload_json.asp',
			fileManagerJson : '../../asp/file_manager_json.asp',
			allowFileManager : true
			 
		});
	</script>



<%
dim sql4,rs4,id,num,pkid


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
		


%>



</head>

<body bgcolor="#fcfcfc">
<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0" id="tabcss">
  <form name="theForm" method="post" action="siterightadmin.asp">
    <tr bgcolor="#E1EFFF"> 
      <td height="24" ><strong>版权信息设置（网页尾部信息）</strong></td>
    </tr>
    <tr> 
      <td height="24"><p> 
      <textarea id="copyright" name="copyright" cols="100" rows="8" style="width:720px;height:400px;visibility:hidden;"><%=htmlspecialchars(copyright)%></textarea>
           
          <br>
        </p></td>
    </tr>
    <tr> 
      <td height="54"> <input type="submit" name="Submit" value=" 保存 "> </td>
    </tr>
  </form>
</table>
<!--
<a class="style1" style="Cursor:hand" onClick="MM_openBrWindow('upload.asp','','width=350,height=210')">上传图片</a>
-->
</body>
</html>


