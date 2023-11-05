<!--#include file="conn.asp"-->
<html>
<head>
<title>上传图片</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
input{
      	BORDER-RIGHT: #FFCC66 1px solid; BORDER-TOP: #FFCC66 1px solid; FONT-SIZE: 9pt; BORDER-LEFT: #FFCC66 1px solid; BORDER-BOTTOM: #FFCC66 1px solid; FONT-FAMILY: "宋体"; BACKGROUND-COLOR: #FAFAFA
}
td { font-size: 12px; line-height: 160%;font-family: "宋体"}
</style>
<script language="vbScript">
<!--
function juge()

note=document.theForm.text1.value
session("note")=note

-->
</script>

</head>
<body  text="#000000" leftmargin="0" topmargin="0" >
<table width="336" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="336" height="74" valign="top" align=center><br>
<%
	response.write "上传图片(50K以内)"
%>
<br>
      选择的图片文件名要为字母或数字。 <br> 
      <table width="314" height="77" border="0" cellpadding="0" cellspacing="0">
<tr> 
          <form name="theForm" method="post" action="saveupload.asp" enctype="multipart/form-data">
            <td width="314" height="77"> 
              <div align="center"> 
			  	<input type="hidden" name="size" value="<%=request("size")%>" >
                <input type="hidden" name="filepath" value="uploadimages" >
                <input type="hidden" name="act" value="upload">
                <input type="file" name="file1" SIZE="32">
                <br>
                <br>
                <input type="submit" name="Submit" value=" 上传图片 " onclick="juge"><br>
			
              </div></td>
          </form>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="53" valign="top">&nbsp;</td>
  </tr>
</table>
</body>
</html>
