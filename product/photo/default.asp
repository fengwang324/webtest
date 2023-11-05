<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>




<script type="text/javascript" src="Jquery.js"></script>

</head>

<body>
<table width="600" border="1" cellpadding="0" cellspacing="0">  


	<tr style="display:" class="hback"><td width="81" height="25" align="right" >商品小图：</td>
		<td ><input type="text" name="sPictures" id="sPictures" value=""
		 style="width:280px;" align="absmiddle">
		  <label for="button"></label>
		<input type="button" name="button" id="button" value="选择图片.." onClick="OpenThenSetValue('Admin_UploadFile.asp?ChannelID=1&CurrentDir=',700,450,window,$('#sPictures')[0]);">
		</td>
	</tr>
	<script type="text/javascript">

	 function OpenThenSetValue(Url,Width,Height,WindowObj,SetObj){
		if (document.all){
		var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:no;help:no;scroll:no;status:0;help:0;scroll:0;');
			if (ReturnStr!='') 
			{
				//var str=new Array()
				//str=ReturnStr.split("/")
				if(typeof(ReturnStr)!=='undefined'){
				
				var str=ReplaceAll(ReturnStr,'/allpic/','')
				
				SetObj.value=str;
				return ReturnStr;}
		}
		}else{
		 obj=SetObj;
		 Width=Width+180;
		 Height=Height+80;
		 window.open(Url,'newWin','modal=yes,width='+Width+',height='+Height+',resizable=no,scrollbars=no');
		}
	 }
	function ReplaceAll(str, sptr, sptr1)
	{
	while (str.indexOf(sptr) >= 0)
	{
	   str = str.replace(sptr, sptr1);
	}
	return str;
	}
	</script>
	<tr style="display:" class="hback"><td width="81" height="25" align="right" >商品大图：</td>
		<td ><input type="text" name="BPictures" id="BPictures" value=""
		 style="width:280px;" align="absmiddle">
		  <label for="button2"></label>
		  <input type="button" name="button3" id="button2" value="选择图片.." onClick="OpenThenSetValue('Admin_UploadFile.asp?ChannelID=1&CurrentDir=',700,450,window,$('#BPictures')[0]);">
		  </td>
	</tr>


</table>

</body>
</html>
