<!-- #include file="conn.asp" -->
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
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>增加商品资料</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<script language=javascript src=../images/mouse.js></script>
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
  if (theForm.kind.value == "")
  {
    alert("请选择分类.");
    document.theForm.kind.focus();
    return false;
  }
  if (theForm.model.value == "")
  {
    alert("请填写货物代码.");
    document.theForm.model.focus();
    return false;
  }
  if (theForm.productname.value == "")
  {
    alert("请填写货物名称");
    document.theForm.productname.focus();
    return false;
  }
  
  if (theForm.price1.value == "")
  {
    alert("请填写市场价");
    document.theForm.price1.focus();
    return false;
  }
  if (theForm.price2.value == "")
  {
    alert("请填写普通会员价");
    document.theForm.price2.focus();
    return false;
  }
  if (theForm.price3.value == "")
  {
    alert("请填写铜级会员价");
    document.theForm.price3.focus();
    return false;
  }
  if (theForm.price4.value == "")
  {
    alert("请填写银级会员价");
    document.theForm.price4.focus();
    return false;
  }
  if (theForm.price5.value == "")
  {
    alert("请填写金级会员价");
    document.theForm.price5.focus();
    return false;
  }
  if (theForm.price6.value == "")
  {
    alert("请填写钻级会员价");
    document.theForm.price6.focus();
    return false;
  }

  if (isNaN(theForm.oneweight.value))
  {
    alert("单件重量要为数字。");
    document.theForm.price6.focus();
    return false;
  }
 
}
//-->
</script>

<script language="JavaScript">
<!--
function myprice()
{

  if (theForm.price2.value == "")
  {
    alert("请填写普通会员价");
    document.theForm.price2.focus();
    return false;
  }
  theForm.price3.value = theForm.price2.value;
  theForm.price4.value = theForm.price2.value;
  theForm.price5.value = theForm.price2.value;
  theForm.price6.value = theForm.price2.value;

}
//-->
</script>
<%


Dim htmlData

htmlData = Request.Form("disc")

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
			id : 'disc',
			urlType : 'absolute',
			imageUploadJson : '../../asp/upload_json.asp',
			fileManagerJson : '../../asp/file_manager_json.asp',
			allowFileManager : true
			 
		});
	</script>
<style>

body,td { font-size: 9pt; color: #000000;font-family: "宋体"}

.border {
	BORDER-BOTTOM: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid
}
input {
      	height=20px;BORDER-RIGHT: #999999 1px solid; BORDER-TOP: #999999 1px solid; FONT-SIZE: 9pt; BORDER-LEFT: #999999 1px solid; BORDER-BOTTOM: #999999 1px solid; 
		FONT-FAMILY: "宋体"; BACKGROUND-COLOR: #ffffff
}
.style1 {
	font-size: 14pt;
	font-weight: bold;
	color: #FF0000;
}
body {
	margin-top: 5px;
	margin-left: 10px;
}
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
	line-height: 24px;
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

#pictable{
	BORDER-COLLAPSE: collapse;
	border-top-width: 0px;
	border-left-width: 0px;
	border-top-style: solid;
	border-left-style: solid;
	border-top-color: #dddddd;
	border-left-color: #dddddd;
}
#pictable td {
	line-height: 24px;
	padding-left: 3px;
	padding-top: 2px;
	padding-bottom: 2px;
	color: #333333;
	border-right-width: 0px;
	border-bottom-width: 0px;
	border-right-style: solid;
	border-bottom-style: solid;
	border-right-color: #dddddd;
	border-bottom-color: #dddddd;
}
.pictxtwidth { width:225px;}

</style>
<script type="text/javascript" src="photo/Jquery.js"></script>
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


</head>
<body bgcolor="#fcfcfc">

<form name="theForm" method="post" action="prodadmin.asp"  OnSubmit="return Juge()">


  <table width="730" border="0" align="left" cellpadding="0" cellspacing="0" id="tabcss">
    <tr bgcolor="#BFDFFF"> 
      <td colspan="2"><b>商品资料新增</b></td>
    </tr>
    <tr> 
      <td >上市时间： <input name="addtime" type="text" id="addtime" value='<%=date()%>'></td>
      <td>种类：　　
        <select name="kind" style="width:150px; font-size:12px; ">
          <%
	  dim rs,sql,maxkind,kindnum,kindname,grade,i,gradestr

	response.write "<option value=''>##请选择分类##</option>"

		sql="select * from sh_kind order by kindnum"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
		if rs.bof or rs.eof then
		else
		do while not rs.eof
			kindnum=rs("kindnum")
			kindname=rs("kindname")
			grade=len(kindnum)/4
			'p=1
			gradestr=""
			'for p=1 to grade-1
				gradestr=gradestr & String(grade-1, "—")
			'next
			response.write "<option value='"&kindnum&"'>|—"&gradestr&kindname&"</option>"
		rs.movenext
		loop
		end if
		rs.close
		set rs=nothing
%>
        </select></td>
    </tr>
    <tr> 
      <td > 商品代码： 
        <input type="text" name="model"  size="32"> </td>
      <td> 商品名称： 
        <input type="text" name="productname" size="40"></td>
    </tr>
    <tr> 
      <td >市场价：　 
        <input name="price1" type="text" id="price1" value="<%=price1%>" size="13">
		单位：<select name="punit" style="width:70px; font-size:12px;"><option value='<%=punit%>'><%=punit%></option> <option value='套'>套</option><option value='件'>件</option><option value='只'>只</option><option value='个'>个</option><option value='台'>台</option><option value='支'>支</option><option value='盒'>盒</option><option value='袋'>袋</option></select>		</td>
      <td>会员价：　 
        <input name="price2" type="text" id="price2">(普通会员价)<input type="button" name="Buttonprice" value="全部会员价相同" title="方便录入其他会员价。点击这里后其他级别价格和普通会员价一样。" onClick="myprice()" style="background-color:#cccccc;width:90px;"></td>
    </tr>
    <tr> 
      <td >铜级价：　 
        <input name="price3" type="text" id="price3"></td>
      <td>银级价：　 
        <input name="price4" type="text" id="price4"> </td>
    </tr>
    <tr> 
      <td >金级价：　 
        <input name="price5" type="text" id="price5"></td>
      <td>钻级价：　 
        <input name="price6" type="text" id="price6"> </td>
    </tr>
    <tr> 
      <td align="left" height=30> <span style="position:absolute;"> 
        <table border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><font style="position:absolute;top:-11px; width: 152px; height: 0px; left: 0px; ">品牌：</font>
              <select name="selectmenu" style="position:absolute;top:-11px; width: 152px; height: 0px; left: 45px; clip: rect(0 180 110 133) "  onChange="theForm.pipai.value=this.options[this.selectedIndex].value">
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
              </select> <input name="pipai" type="text" value="<%=pipai%>" style="position:absolute;top:-11px; width: 138px; height:21px; left:45px;" onKeyDown="if(event.keyCode==13)event.keyCode=9">            </td>
          </tr>
        </table>
        </span> </td>
      <td>提示：如果不分会员级别，等级会员价填写与普通会员价一样即可。</td>
    </tr>

    <tr>
      <td>单件重量： 
		<input name="oneweight" type="text" size="25" maxlength="100" value="<%=oneweight%>">(克)
		<img src='../images/helpmemo.gif' alt="如果运费是按重量计算的，请填写商品的重量，<br>如果运费是按件计算的，填写重量或填写0。<br>
		具体请看“配送计费”-“选择配送计费方式”。">		</td>
      <td>库存数量： 
		<input name="kuchushu" type="text" size="25" maxlength="100" value="<%=kuchushu%>"></td>		
    </tr>	
	
	<%if s_colorsize="是" then%>
	<%end if%>	
	
	<!--
    <tr>
      <td height="24" >商品发货地： 
		<input name="fahoudi" type="text" size="23" maxlength="100" value="<%=fahoudi%>"></td>
      <td height="24" >库存说明：
		<input name="kuchumemo" type="text" size="25" maxlength="100" value="<%=kuchumemo%>"></td>
    </tr>
	-->

<tr><td colspan=2>
	<table width="100%" border=0  cellpadding="0" cellspacing="0" id="pictable">
    <tr> 
      <td width="65%"> 
		<input type="hidden" name="smallpicpathold" value="<%=smallpicpath%>">
		小图&nbsp;&nbsp;<input type="text" name="smallpicpath" id="smallpicpath" value="<%=smallpicpath%>" class="pictxtwidth">
        <input type="button" name="Button" style="background-color:#ddddd1;width:65px;" value="单个上传" onClick="MM_openBrWindow('upload.asp?size=smallpicpath','','width=370,height=290,top=124,left=470')">
        <input type="button" name="button" id="button" value="选图…" title="点击这里可以先批量上传图片，或者进行选择图片。" onClick="OpenThenSetValue('photo/Admin_UploadFile.asp?ChannelID=1&CurrentDir=',700,450,window,$('#smallpicpath')[0]);">
		<img src="../images/img.gif" width="20" height="20" border="0" align="absmiddle" style="cursor:hand;" onMouseOver="if(document.all('smallpicpath').value!=''){document.all('photo_2').src =document.all('smallpicpath').value; document.getElementById('showpic').href=document.all('smallpicpath').value;}">
		&nbsp;<img src='../images/helpmemo.gif' alt="单个上传：是指单个图片上传；<br>选图：可以<b>批量上传</b>本商品及其他商品的图片，然后在这里选择。<br>注意：放到虚拟主机空间服务器才支持“选图…”批量上传图片。<br><font color='#ff0000'>如果有在本地电脑通过红B运行的，只可单个上传，不支持图片批量上传。</font><br>在本地电脑上，直接将图片放到product\uploadimages目录下即可进行选图。<br><br>.jpg或.gif 宽度高度为“模板设置”-“显示格式”中设定，<br>系统默认为宽:135~145px，高170px，50KB内，名为英文或数字。<br>鼠标放在左边黄色小图，图片在右边预览。">		</td>
	    <td width="35%" rowspan="4"><a href='' id='showpic' target='_blank'><img name='photo_2' height="140"  width="140" border="0" src="" alt="鼠标放在左边黄色小图,图片在这预览,点击可查看原图" title="图片预览区域。点击可查看原图。"></a></td>
    </tr>
    <tr>
      <td height="22">
	  <input type="hidden" name="bigpicpathold" value="<%=bigpicpath%>">
	  	大图
	  	<input type="text" name="bigpicpath" id="bigpicpath" value="<%=bigpicpath%>" class="pictxtwidth">
        <input type="button" name="Button22" style="background-color:#ddddd1;width:65px;" value="单个上传" onClick="MM_openBrWindow('upload.asp?size=bigpicpath','','width=370,height=290,top=124,left=470')">
        <input type="button" name="button" id="button" value="选图…" title="点击这里可以先批量上传图片，或者进行选择图片。" onClick="OpenThenSetValue('photo/Admin_UploadFile.asp?ChannelID=1&CurrentDir=',700,450,window,$('#bigpicpath')[0]);">
		<img src="../images/img.gif" width="20" height="20" border="0" align="absmiddle" style="cursor:hand;" onMouseOver="if(document.all('bigpicpath').value!=''){document.all('photo_2').src =document.all('bigpicpath').value; document.getElementById('showpic').href=document.all('bigpicpath').value;}">
		&nbsp;<img src='../images/helpmemo.gif' alt="单个上传：是指单个图片上传；<br>选图：可以<b>批量上传</b>本商品及其他商品的图片，然后在这里选择。<br>注意：放到虚拟主机空间服务器才支持“选图…”批量上传图片。<br><font color='#ff0000'>如果有在本地电脑通过红B运行的，只可单个上传，不支持图片批量上传。</font><br>在本地电脑上，直接将图片放到product\uploadimages目录下即可进行选图。<br><br>.jpg或.gif 最佳约宽:630px，高:630~650px 150KB内 图片名为英文或数字。<br>鼠标放在左边黄色小图，图片在右边预览。">        </td>
    </tr>
	</table>
</td></tr>


    <tr> 
      <td valign="top" colspan="2">详细资料 
        <textarea id="disc" name="disc" cols="100" rows="8" style="width:725px;height:400px;visibility:hidden;"></textarea>
          </td>
    </tr>

     <tr>
		<td height="40" colspan="2"><div align="center"><input type="Submit" name="Submit" value="  新  增  "    style="background-color:#cccccc;width:90px; height:25px;"></div></td>
    </tr>
	<tr height="50"><td colspan="2"></td></tr>
  </table>
  

</form>
</body>
</html>

