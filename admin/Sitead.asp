<!--#include file="conn.asp"-->
<!--#include file="check2.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>网站后台,广告设置</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
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
.STYLE4 {
	color: #FF0000;
	font-weight: bold;
}
-->
</style>
<%
dim sql4,rs4,id,num,pkid
dim biaoflag,lefturl,righturl,bannarflag,bannarurl,changeflag,curl1,curl2,curl3,curl4,curl5


Submit2 = trim(request.form("Submit2"))
if Submit2="保存" then

	biaoflag = trim(request.form("biaoflag"))
	
	lefturl = trim(request.form("lefturl"))
	leftlink = trim(request.form("leftlink"))
	righturl = trim(request.form("righturl"))
	rightlink = trim(request.form("rightlink"))
	
	bannarflag = trim(request.form("bannarflag"))
	bannarurl = trim(request.form("bannarurl"))
	bannarlink = trim(request.form("bannarlink"))
	
	changeflag = trim(request.form("changeflag"))
	curl1 = trim(request.form("curl1"))
	curl2 = trim(request.form("curl2"))
	curl3 = trim(request.form("curl3"))
	curl4 = trim(request.form("curl4"))
	curl5 = trim(request.form("curl5"))

	clink1 = trim(request.form("clink1"))
	clink2 = trim(request.form("clink2"))
	clink3 = trim(request.form("clink3"))
	clink4 = trim(request.form("clink4"))
	clink5 = trim(request.form("clink5"))

	ctext1 = trim(request.form("ctext1"))
	ctext2 = trim(request.form("ctext2"))
	ctext3 = trim(request.form("ctext3"))
	ctext4 = trim(request.form("ctext4"))
	ctext5 = trim(request.form("ctext5"))

	hotpicurl = trim(request.form("hotpicurl"))
	hotpiclink = trim(request.form("hotpiclink"))
	
	mid_flag = trim(request.form("mid_flag"))
	rightpicurl1 = trim(request.form("rightpicurl1"))
	rightpiclink1 = trim(request.form("rightpiclink1"))
	rightpicurl2 = trim(request.form("rightpicurl2"))
	rightpiclink2 = trim(request.form("rightpiclink2"))
	rightpicurl3 = trim(request.form("rightpicurl3"))
	rightpiclink3 = trim(request.form("rightpiclink3"))

	topnewspicurl = trim(request.form("topnewspicurl"))
	topnewspiclink = trim(request.form("topnewspiclink"))

	
	cuxpicurl = trim(request.form("cuxpicurl"))
	cuxpiclink = trim(request.form("cuxpiclink"))
	
	newpicurl = trim(request.form("newpicurl"))
	newpiclink = trim(request.form("newpiclink"))

		sql4="select * from ad"
		set rs4=server.createobject("adodb.recordset")
		rs4.open sql4,conn,1,3
		if rs4.bof or rs4.eof then
				rs4.addnew
				rs4("biaoflag")=biaoflag
				
				rs4("lefturl")=lefturl
				rs4("leftlink")=leftlink
				rs4("righturl")=righturl
				rs4("rightlink")=rightlink
				
				rs4("bannarflag")=bannarflag
				rs4("bannarurl")=bannarurl
				rs4("bannarlink")=bannarlink
				
				rs4("changeflag")=changeflag
				rs4("curl1")=curl1
				rs4("curl2")=curl2
				rs4("curl3")=curl3
				rs4("curl4")=curl4
				rs4("curl5")=curl5
			
				rs4("clink1")=clink1
				rs4("clink2")=clink2
				rs4("clink3")=clink3
				rs4("clink4")=clink4
				rs4("clink5")=clink5

				rs4("ctext1")=ctext1
				rs4("ctext2")=ctext2
				rs4("ctext3")=ctext3
				rs4("ctext4")=ctext4
				rs4("ctext5")=ctext5

				rs4("hotpicurl")=hotpicurl
				rs4("hotpiclink")=hotpiclink
				
				rs4("mid_flag")=mid_flag
				rs4("rightpicurl1")=rightpicurl1
				rs4("rightpiclink1")=rightpiclink1
				rs4("rightpicurl2")=rightpicurl2
				rs4("rightpiclink2")=rightpiclink2
				rs4("rightpicurl3")=rightpicurl3
				rs4("rightpiclink3")=rightpiclink3
				
				rs4("topnewspicurl")=topnewspicurl
				rs4("topnewspiclink")=topnewspiclink				
				
				
				rs4("cuxpicurl")=cuxpicurl
				rs4("cuxpiclink")=cuxpiclink
				
				rs4("newpicurl")=newpicurl
				rs4("newpiclink")=newpiclink

				rs4.update
		else
				rs4("biaoflag")=biaoflag
				
				rs4("lefturl")=lefturl
				rs4("leftlink")=leftlink
				rs4("righturl")=righturl
				rs4("rightlink")=rightlink
				
				rs4("bannarflag")=bannarflag
				rs4("bannarurl")=bannarurl
				rs4("bannarlink")=bannarlink
				
				rs4("changeflag")=changeflag
				rs4("curl1")=curl1
				rs4("curl2")=curl2
				rs4("curl3")=curl3
				rs4("curl4")=curl4
				rs4("curl5")=curl5
			
				rs4("clink1")=clink1
				rs4("clink2")=clink2
				rs4("clink3")=clink3
				rs4("clink4")=clink4
				rs4("clink5")=clink5

				rs4("ctext1")=ctext1
				rs4("ctext2")=ctext2
				rs4("ctext3")=ctext3
				rs4("ctext4")=ctext4
				rs4("ctext5")=ctext5

				rs4("hotpicurl")=hotpicurl
				rs4("hotpiclink")=hotpiclink
				
				rs4("mid_flag")=mid_flag
				rs4("rightpicurl1")=rightpicurl1
				rs4("rightpiclink1")=rightpiclink1
				rs4("rightpicurl2")=rightpicurl2
				rs4("rightpiclink2")=rightpiclink2
				rs4("rightpicurl3")=rightpicurl3
				rs4("rightpiclink3")=rightpiclink3

				rs4("topnewspicurl")=topnewspicurl
				rs4("topnewspiclink")=topnewspiclink
				
				rs4("cuxpicurl")=cuxpicurl
				rs4("cuxpiclink")=cuxpiclink
				
				rs4("newpicurl")=newpicurl
				rs4("newpiclink")=newpiclink

				rs4.update
		end if
		rs4.close
		set rs4=nothing
		response.write "<script language=JavaScript>" & "alert('保存成功!');"& "window.location.href='sitead.asp';" & "</script>"
end if

%>
<script language=javascript src=../images/mouse.js></script>
</head>

<body bgcolor="#fcfcfc">       
<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0" id="tabcss">
  <tr bgcolor="#E1EFFF"> 
    <td height=30 colspan=2>&nbsp;<b>设置网站广告：</b></td>
  </tr>

  <form name="form1" method="post" action="sitead.asp">
    <%
		
		sql4="select * from ad"
		set rs4=server.createobject("adodb.recordset")
		rs4.open sql4,conn,1,1
		if rs4.bof or rs4.eof then
			response.write "没有符合条件记录！"
		else
			biaoflag=rs4("biaoflag")
			
			lefturl=rs4("lefturl")
			leftlink=rs4("leftlink")
			righturl=rs4("righturl")
			rightlink=rs4("rightlink")
			
			bannarflag=rs4("bannarflag")
			bannarurl=rs4("bannarurl")
			bannarlink=rs4("bannarlink")
			
			changeflag=rs4("changeflag")
			curl1=rs4("curl1")
			curl2=rs4("curl2")
			curl3=rs4("curl3")
			curl4=rs4("curl4")
			curl5=rs4("curl5")
		
			clink1=rs4("clink1")
			clink2=rs4("clink2")
			clink3=rs4("clink3")
			clink4=rs4("clink4")
			clink5=rs4("clink5")

			ctext1=rs4("ctext1")
			ctext2=rs4("ctext2")
			ctext3=rs4("ctext3")
			ctext4=rs4("ctext4")
			ctext5=rs4("ctext5")

			hotpicurl=rs4("hotpicurl")
			hotpiclink=rs4("hotpiclink")
			
			mid_flag=rs4("mid_flag")
			rightpicurl1=rs4("rightpicurl1")
			rightpiclink1=rs4("rightpiclink1")
			rightpicurl2=rs4("rightpicurl2")
			rightpiclink2=rs4("rightpiclink2")
			rightpicurl3=rs4("rightpicurl3")
			rightpiclink3=rs4("rightpiclink3")
			
			topnewspicurl=rs4("topnewspicurl")
			topnewspiclink=rs4("topnewspiclink")

			cuxpicurl=rs4("cuxpicurl")
			cuxpiclink=rs4("cuxpiclink")
			
			newpicurl=rs4("newpicurl")
			newpiclink=rs4("newpiclink")

	%>

    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td width="20%" height='23' align='right'><strong><b>首页</b>切换广告图</strong>
        <input type="hidden" name="changeflag" value="1"></td>
      <td width="80%">上传图片1 
        <input name='curl1' type=text id="curl1" value="<%=curl1%>" size="40"> 
        <input type="button" name="Button2" value="上传图片" onClick="window.open('../images/uploadpic.asp?size=curl1','','width=410,height=350,top=120,left=281')">
        (670 X 300) <img src='../images/helpmemo.gif' alt="图片预览：<br><img src='../<%=curl1%>'>"></td>
    </tr>
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'>&nbsp;</td>
      <td>链接地址1 
        <input name='clink1' type=text id="clink1" value="<%=clink1%>" size="35">
		&nbsp;&nbsp;		</td>
    </tr>
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'>&nbsp;</td>
      <td>上传图片2 
        <input name='curl2' type=text id="curl2" value="<%=curl2%>" size="40"> 
        <input type="button" name="Button2" value="上传图片" onClick="window.open('../images/uploadpic.asp?size=curl2','','width=410,height=350,top=120,left=281')">
        (670 X 300) <img src='../images/helpmemo.gif' alt="图片预览：<br><img src='../<%=curl2%>'>"></td>
    </tr>
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'>&nbsp;</td>
      <td>链接地址2 
        <input name='clink2' type=text id="clink2" value="<%=clink2%>" size="35">
		&nbsp;		</td>
    </tr>
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'>&nbsp;</td>
      <td>上传图片3 
        <input name='curl3' type=text id="curl3" value="<%=curl3%>" size="40"> 
        <input type="button" name="Button2" value="上传图片" onClick="window.open('../images/uploadpic.asp?size=curl3','','width=410,height=350,top=120,left=281')">
        (670 X 300) <img src='../images/helpmemo.gif' alt="图片预览：<br><img src='../<%=curl3%>'>"></td>
    </tr>
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'>&nbsp;</td>
      <td>链接地址3 
        <input name='clink3' type=text id="clink3" value="<%=clink3%>" size="35">
		&nbsp;&nbsp;		</td>
    </tr>
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'>&nbsp;</td>
      <td>上传图片4 
        <input name='curl4' type=text id="curl4" value="<%=curl4%>" size="40"> 
        <input type="button" name="Button2" value="上传图片" onClick="window.open('../images/uploadpic.asp?size=curl4','','width=410,height=350,top=120,left=281')">
        (670 X 300) <img src='../images/helpmemo.gif' alt="图片预览：<br><img src='../<%=curl4%>'>"></td>
    </tr>
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'>&nbsp;</td>
      <td>链接地址4 
        <input name='clink4' type=text id="clink4" value="<%=clink4%>" size="35">
		&nbsp;&nbsp;		</td>
    </tr>
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'>&nbsp;</td>
      <td>上传图片5 
        <input name='curl5' type=text id="curl5" value="<%=curl5%>" size="40"> 
        <input type="button" name="Button2" value="上传图片" onClick="window.open('../images/uploadpic.asp?size=curl5','','width=410,height=350,top=120,left=281')">
        (670 X 300) <img src='../images/helpmemo.gif' alt="图片预览：<br><img src='../<%=curl5%>'>"></td>
    </tr>
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'>&nbsp;</td>
      <td>链接地址5 
        <input name='clink5' type=text id="clink5" value="<%=clink5%>"size="35">
		&nbsp;&nbsp;</td>
    </tr>
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';">
      <td height='23' align='right'>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
	
	
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'><b>首页头部横幅广告图</b></td>
      <td>上传图片 
        <input name='topnewspicurl' type=text id="topnewspicurl" value="<%=topnewspicurl%>" size="40"> 
        <input type="button" name="Button2" value="上传图片" onClick="window.open('../images/uploadpic.asp?size=topnewspicurl','','width=410,height=350,top=120,left=281')">
        (960 X 约55) <img src='../images/helpmemo.gif' alt="图片预览：<br><img src='../<%=topnewspicurl%>'>"></td>
    </tr>
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'>&nbsp;</td>
      <td>链接地址 
        <input name='topnewspiclink' type=text id="topnewspiclink" value="<%=topnewspiclink%>" size="40">
        <strong>可选，</strong>如果图片及链接留空，则这广告栏位将不显示</td>
    </tr>
	
	
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';">
      <td height='23' align='right'>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>

	<!--
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'><b>季节厨窗(热卖Bannar)</b></td>
      <td>上传图片 
        <input name='hotpicurl' type=text id="hotpicurl" value="<%=hotpicurl%>" size="40"> 
        <input type="button" name="Button2" value="上传图片" onclick="window.open('../images/uploadpic.asp?size=hotpicurl','','width=410,height=350,top=120,left=281')">
        (433 X 约85)</td>
    </tr>
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'>&nbsp;</td>
      <td>链接地址 
        <input name='hotpiclink' type=text id="hotpiclink" value="<%=hotpiclink%>" size="40"></td>
    </tr>
	-->
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'><b>首页底部横幅广告</b></td>
      <td>上传图片 
        <input name='newpicurl' type=text id="newpicurl" value="<%=newpicurl%>" size="40"> 
        <input type="button" name="Button2" value="上传图片" onClick="window.open('../images/uploadpic.asp?size=newpicurl','','width=410,height=350,top=120,left=281')">
        (960 X 约120) <img src='../images/helpmemo.gif' alt="图片预览：<br><img src='../<%=newpicurl%>'>"></td>
    </tr>
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'>&nbsp;</td>
      <td>链接地址 
        <input name='newpiclink' type=text id="newpiclink" value="<%=newpiclink%>" size="40">
        <strong>可选，</strong>如果图片及链接留空，则这广告栏位将不显示</td>
    </tr>
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <!--
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'><strong>Bannar图片</strong></td>
      <td>是否显示 
        <input type="radio" name="bannarflag" value="1" <%if bannarflag="1" then response.write "checked"%>>
        是　　 
        <input type="radio" name="bannarflag" value="0" <%if bannarflag="0" then response.write "checked"%>>
        否</td>
    </tr>
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'>&nbsp;</td>
      <td>上传图片 
        <input name='bannarurl' type=text id="bannarurl" value="<%=bannarurl%>" size="40"> 
        <input type="button" name="Button2" value="上传图片" onclick="window.open('../images/uploadpic.asp?size=bannarurl','','width=410,height=350,top=120,left=281')">
        (700 X 100)</td>
    </tr>
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'>&nbsp;</td>
      <td>链接地址 
        <input name='bannarlink' type=text id="bannarlink" value="<%=bannarlink%>" size="40"></td>
    </tr>
	-->
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='35' align='right'>&nbsp;</td>
      <td><input type="submit" name="Submit2" value=" 保存 " onClick="return confirm('确认要保存“广告图片设置”吗？')"> &nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="reset" name="Submit3" value=" 重置 "></td>
    </tr>
    <%				
		end if
rs4.close
set rs4=nothing		
%>
  </form>
</table>
<br><br><br>

</body>
</html>
<%
call connclose
%>

