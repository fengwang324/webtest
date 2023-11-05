<%
 

tupian1=curl1
imgtext1=""
tupian1url=clink1

tupian2=curl2
imgtext2=""
tupian2url=clink2

tupian3=curl3
imgtext3=""
tupian3url=clink3

tupian4=curl4
imgtext4=""
tupian4url=clink4

tupian5=curl5
imgtext5=""
tupian5url=clink5


 
  
%>
<script language="" type="text/javascript">
	<!--
imgUrl1="<%=tupian1%>";
imgtext1=""
imgLink1=escape("<%=tupian1url%>");
imgUrl2="<%=tupian2%>";
imgtext2=""
imgLink2=escape("<%=tupian2url%>");
imgUrl3="<%=tupian3%>";
imgtext3=""
imgLink3=escape("<%=tupian3url%>");
imgUrl4="<%=tupian4%>";
imgtext4=""
imgLink4=escape("<%=tupian4url%>");
imgUrl5="<%=tupian5%>";
imgtext5=""
imgLink5=escape("<%=tupian5url%>");

	var focus_width=724
	var focus_height=295
	var text_height=0
	var swf_height = focus_height+text_height
	
	 var pics=imgUrl1+"|"+imgUrl2+"|"+imgUrl3+"|"+imgUrl4+"|"+imgUrl5
 var links=imgLink1+"|"+imgLink2+"|"+imgLink3+"|"+imgLink4+"|"+imgLink5
 var texts=imgtext1+"|"+imgtext2+"|"+imgtext3+"|"+imgtext4+"|"+imgtext5
 	
	document.write('<object ID="focus_flash" classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ focus_width +'" height="'+ swf_height +'">');
	document.write('<param name="allowScriptAccess" value="sameDomain"><param name="movie" value="images/pixviewer.swf"><param name="quality" value="high"><param name="bgcolor" value="#AFDBFE">');
	document.write('<param name="menu" value="false"><param name=wmode value="opaque">');
	document.write('<param name="FlashVars" value="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'">');
	document.write('<embed ID="focus_flash" src="images/pixviewer.swf" wmode="opaque" FlashVars="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'" menu="false" bgcolor="#5A9FB5" quality="high" width="'+ focus_width +'" height="'+ swf_height +'" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />');		
	document.write('</object>');
	
	//-->
	
</script>