<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>批量上传图片/选择图片</title>
</head>
    <frameset rows="*" cols="0,*" framespacing="0" frameborder="0" border="false" id="frame" scrolling="yes">
        <frame name="UploadFile_left" scrolling="auto" src="Admin_UploadFile_Left.asp?ChannelID=<%=request("ChannelID")%>&CurrentDir=<%=request("CurrentDir")%>">
        <frame name="UploadFile_Main" scrolling="auto" src="Admin_UploadFile_Main.asp?ChannelID=<%=request("ChannelID")%>&CurrentDir=<%=request("CurrentDir")%>">
    </frameset>
<noframes>
  <body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
  <p>你的浏览器版本过低！！！本系统要求IE5及以上版本才能使用本系统。</p>
  </body>
</noframes>
</html>
