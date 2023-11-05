<!--#include file="conn.asp"-->
<!--#include file="check2.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>网站后台,网站基本信息</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<script language=javascript src=../images/mouse.js></script>
<style>
<!--
td{font-size:9pt}
alt{font-size:9pt}
body {
	margin-top: 5px;margin-left: 2px;FONT: 12px arial,宋体; 
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
	line-height: 22px;
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
<%
dim sql4,rs4,id,num,pkid

s_qq=trim(request.form("s_qq"))


'检查组件是否被支持
'theInstalledObjects(9) = "Scripting.FileSystemObject"
Function IsObjInstalled(strClassString)
	'on error resume next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function

Function CopyMyFolder(FolderName,FolderPath) 
	sFolder=server.mappath(FolderName) '源文件夹
	tFolder=server.mappath(FolderPath) '目标文件夹
	set fso=server.createobject("scripting.filesystemobject") 
	if fso.folderexists(server.mappath(FolderName)) Then'检查原文件夹是否存在 
		if fso.folderexists(server.mappath(FolderPath)) Then'检查目标文件夹是否存在 
			fso.copyfolder sFolder,tFolder 
		Else '目标文件夹如果不存在就创建 
			CreateNewFolder = Server.Mappath(FolderPath) 
			fso.CreateFolder(CreateNewFolder) 
			fso.CopyFolder sFolder,tFolder 
		End if 
		CopyMyFolder="复制文件夹["&server.mappath(FolderName)&"]到["&server.mappath(FolderPath)&"]成功!" 
	Else 
		CopyMyFolder="错误,原文件夹["&sFolde&"]不存在!" 
	End If 
	set fso=nothing 
End Function


Submit2=trim(request.form("Submit2"))

if Submit2="保存" then

		sql4="select * from siteshow"
		set rs4=server.createobject("adodb.recordset")
		rs4.open sql4,conn,1,3
		if rs4.bof or rs4.eof then
		else
			rs4("s_col")=trim(request.form("s_col"))
			rs4("s_row")=trim(request.form("s_row"))	
			rs4("s_picwidth")=trim(request.form("s_picwidth"))
			rs4("s_picheight")=trim(request.form("s_picheight"))
			rs4("s_colorsize")=trim(request.form("s_colorsize"))
			
			rs4("s_model")=trim(request.form("s_model"))
			rs4("s_pipai")=trim(request.form("s_pipai"))
			rs4("s_kuchu")=trim(request.form("s_kuchu"))
			
			rs4("s_hot_line")=trim(request.form("s_hot_line"))
			rs4("s_cuxiao_line")=trim(request.form("s_cuxiao_line"))
			rs4("s_new_line")=trim(request.form("s_new_line"))
			
			rs4("s_news")=trim(request.form("s_news"))
			

			rs4("s_productkind")=4  'trim(request.form("s_productkind")) '此版本没有左边商品分类
			If not IsObjInstalled("Scripting.FileSystemObject") Then  '看是否支持FSO
			else
			rs4("s_head")="red"
			rs4("s_foot")="foot1.asp"
			end if
			rs4("s_hit")="是"  'trim(request.form("s_hit"))
			rs4("s_hitnum")=trim(request.form("s_hitnum"))
			rs4("s_diaocha")="是"  'trim(request.form("s_diaocha"))
			rs4("s_bbs")="是"  'trim(request.form("s_bbs"))
			rs4("s_mes")=trim(request.form("s_mes"))
			
			rs4("s_youqing")="否"
			rs4("s_zishu")="否"

			rs4("s_msn")=trim(request.form("s_msn"))
			rs4("s_wangwang")=trim(request.form("s_wangwang"))
			rs4("s_pinpai")=trim(request.form("s_pinpai"))
			rs4("s_pinpai_l_r")=trim(request.form("s_pinpai_l_r"))
			rs4("s_pinpaiw")=trim(request.form("s_pinpaiw"))
			rs4("s_pinpaih")=trim(request.form("s_pinpaih"))
			
			rs4.update
			
			If not IsObjInstalled("Scripting.FileSystemObject") Then  '看是否支持FSO
			else
				'--------------------换头部begin-----------------------------
				s_head=trim(request.form("s_head"))
				s_headold=trim(request.form("s_headold"))
				's_img=mid(s_head,5,1)		'看是哪个模板数
				's_imgold=mid(s_headold,5,1)	'看是哪个模板数
				
				if s_head<>s_headold and s_head<>"" and s_headold<>"" then
					FolderName="../muban/"&s_head&"" '原文件夹 
					FolderPath="../images" '目标文件夹 
					response.write""&CopyMyFolder(FolderName,FolderPath)&""
				end if
				'--------------------换头部end-------------------------------
				
				'--------------------换尾部begin-----------------------------
				s_foot=trim(request.form("s_foot"))
				s_footold=trim(request.form("s_footold"))
				if s_foot<>s_footold and s_foot<>"" and s_footold<>"" then
					set fso= server.CreateObject("scripting.fileSystemObject")
					p1="../foot.asp"
					p11="../muban/"&s_footold
					set f1=fso.getfile(server.mappath(""&p1&""))
					f1.copy(server.mappath(""&p11&""))
					
					p2="../muban/"&s_foot
					p21="../foot.asp"
					set f2=fso.getfile(server.mappath(""&p2&""))
					f2.copy(server.mappath(""&p21&""))
					set fso=nothing
					set f1=nothing
					set f2=nothing
				end if
				'--------------------换尾部end-------------------------------
			end if
		end if
		rs4.close
		set rs4=nothing
		response.write "<script language=JavaScript>" & "alert('保存成功!');"& "window.location.href='siteshow.asp';" & "</script>"
end if

%>
<script language="JavaScript">
<!--
function Juge()
{
  if (theForm.s_col.value == "")
  {
    alert("请您填写列数!");
    theForm.s_col.focus();
    return false;
  }
  if (isNaN(theForm.s_col.value))
  {
    alert("列数要为数字!");
    theForm.s_col.focus();
    return (false);
  }
  
  if (theForm.s_row.value == "")
  {
    alert("请您填写行数!");
    theForm.s_row.focus();
    return false;
  }
  if (isNaN(theForm.s_row.value))
  {
    alert("行数要为数字!");
    theForm.s_row.focus();
    return (false);
  }
  

  if (theForm.s_picwidth.value == "")
  {
    alert("请您填写宽度!");
    theForm.s_picwidth.focus();
    return false;
  }
  if (isNaN(theForm.s_picwidth.value))
  {
    alert("宽度要为数字!");
    theForm.s_picwidth.focus();
    return (false);
  }

  if (theForm.s_picheight.value == "")
  {
    alert("请您填写宽度!");
    theForm.s_picheight.focus();
    return false;
  }
  if (isNaN(theForm.s_picheight.value))
  {
    alert("高度要为数字!");
    theForm.s_picheight.focus();
    return (false);
  }



  if (theForm.s_hot_line.value == "")
  {
    alert("请您填写热卖商品行数!");
    theForm.s_hot_line.focus();
    return false;
  }
  if (isNaN(theForm.s_hot_line.value))
  {
    alert("热卖商品行数要为数字!");
    theForm.s_hot_line.focus();
    return (false);
  }  

  if (theForm.s_cuxiao_line.value == "")
  {
    alert("请您填写促销商品行数!");
    theForm.s_cuxiao_line.focus();
    return false;
  }
  if (isNaN(theForm.s_cuxiao_line.value))
  {
    alert("促销商品行数要为数字!");
    theForm.s_cuxiao_line.focus();
    return (false);
  }  

  if (theForm.s_new_line.value == "")
  {
    alert("请您填写最新商品行数!");
    theForm.s_new_line.focus();
    return false;
  }
  if (isNaN(theForm.s_new_line.value))
  {
    alert("最新商品行数要为数字!");
    theForm.s_new_line.focus();
    return (false);
  } 


  if (theForm.s_news.value == "")
  {
    alert("请您填写主页最新商讯显示条数!");
    theForm.s_news.focus();
    return false;
  }
  if (isNaN(theForm.s_news.value))
  {
    alert("主页最新商讯显示条数要为数字!");
    theForm.s_news.focus();
    return (false);
  }  


  if (theForm.s_mes.value == "")
  {
    alert("请您填写主页购物社区显示条数!");
    theForm.s_mes.focus();
    return false;
  }
  if (isNaN(theForm.s_mes.value))
  {
    alert("主页购物社区显示条数要为数字!");
    theForm.s_mes.focus();
    return (false);
  } 

  if (theForm.s_hitnum.value == "")
  {
    alert("请输入点击率排行条数!");
    theForm.s_hitnum.focus();
    return false;
  }
  if (isNaN(theForm.s_hitnum.value))
  {
    alert("点击率排行条数要为数字!");
    theForm.s_hitnum.focus();
    return (false);
  } 

   if(confirm("确认要这样设置吗?"))
      return true
   else
      return false;
}
//-->
</script>
</head>

<body bgcolor="#fcfcfc">          
<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0" id="tabcss">
  <tr bgcolor="#E1EFFF"> 
    <td height=30 colspan=2>&nbsp;<b>设置模板显示格式或内容：(在这里，选择头、尾部模板；进行设置，让页面布局显示得整齐；可按实际情况取舍一些功能等等)</b></td>
  </tr>
  <form name="theForm" method="post" action="siteshow.asp" onSubmit="return Juge()">
    <%
		
		sql4="select * from siteshow"
		set rs4=server.createobject("adodb.recordset")
		rs4.open sql4,conn,1,1
		if rs4.bof or rs4.eof then
			response.write "没有符合条件订单记录！"
		else
			s_col=rs4("s_col")
			s_row=rs4("s_row")			
			s_picwidth=rs4("s_picwidth")
			s_picheight=rs4("s_picheight")
			s_colorsize=rs4("s_colorsize")

			s_model=rs4("s_model")
			s_pipai=rs4("s_pipai")
			s_kuchu=rs4("s_kuchu")
			
			s_hot_line=rs4("s_hot_line")
			s_cuxiao_line=rs4("s_cuxiao_line")
			s_new_line=rs4("s_new_line")
			
			s_news=rs4("s_news")
			s_mes=rs4("s_mes")
			s_hit=rs4("s_hit")
			s_hitnum=rs4("s_hitnum")
			s_productkind=rs4("s_productkind")
			s_head=rs4("s_head")
			s_foot=rs4("s_foot")
			s_diaocha=rs4("s_diaocha")
			s_bbs=rs4("s_bbs")
			s_youqing=rs4("s_youqing")
			s_zishu=rs4("s_zishu")
			s_qq=rs4("s_qq")
			s_msn=rs4("s_msn")
			s_wangwang=rs4("s_wangwang")

			s_pinpai=rs4("s_pinpai")
			s_pinpai_l_r=rs4("s_pinpai_l_r")
			s_pinpaiw=rs4("s_pinpaiw")
			s_pinpaih=rs4("s_pinpaih")
%>
	
	<!--
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'>左边商品分类显示方式 
      </td>
      <td>
	  <input type="radio" name="s_productkind" value="1" <%if s_productkind="1" then response.write "checked"%>>商品分类自动换行&nbsp;&nbsp;&nbsp; 
	  <input type="radio" name="s_productkind" value="2" <%if s_productkind="2" then response.write "checked"%>>每行显示两个分类&nbsp;&nbsp;&nbsp; 
	  <input type="radio" name="s_productkind" value="3" <%if s_productkind="3" then response.write "checked"%>>每行显示三个分类&nbsp;&nbsp;&nbsp; 
	  <input type="radio" name="s_productkind" value="4" <%if s_productkind="4" then response.write "checked"%>>商品分类右拉方式
	  &nbsp;（请按您的商品分类多少及分类文字长短作选择合适的风格）</td>
    </tr>
	-->
	
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'>主页最新商讯显示条数</td>
      <td> <input type=text name='s_news' value="<%=s_news%>" size="4">      </td>
    </tr>
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' width="20%" align='right'>首页各商品的行数</td>
      <td width="80%">
	  热卖商品行数 <input type=text name='s_hot_line' value="<%=s_hot_line%>" size="5">&nbsp;&nbsp;
	  促销商品行数 <input type=text name='s_cuxiao_line' value="<%=s_cuxiao_line%>" size="5">&nbsp;&nbsp;
	  最新商品行数 <input type=text name='s_new_line' value="<%=s_new_line%>" size="5">
	  &nbsp;&nbsp;<br>
      (如果不要显示热卖商品、促销商品栏目，可以在“设定栏目”中将这些栏目的“是否显示”改为“否”)</td>
    </tr>


    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' width="20%" align='right'>前台商品列表</td>
      <td width="80%"> 每页列数 
        <input type=text name='s_col' value="<%=s_col%>" size="5">
        (通常4-5之间)&nbsp;&nbsp;&nbsp; 行数 
        <input type=text name='s_row' value="<%=s_row%>" size="5">
        (通常为4-6之间，请按您的实际情况修改) 
		<img src='../images/helpmemo.gif' alt="<font color='#ff3300'>前台商品列表小图设置建议：</font><br>由于各个行业的商品图规格差别较大（有的可能要高度大于宽度或宽度大于高度，用户要求商品小图高度/宽度不一），<br>所以在将开网店时建议计划好商品列表小图宽度、高度、列数，方便以后按这个尺寸做产品小图，可以先看系统自带的<br>商品图数据，设置下面参数，看下显示效果。如果下面设显示4列，小图宽度设为170，显示效果较佳。如果5列，小图<br>宽度为135，显示效果较佳。">		</td>
    </tr>
	
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' width="20%" align='right'>商品列表小图片</td>
      <td width="80%"> 宽度 
        <input type=text name='s_picwidth' value="<%=s_picwidth%>" size="5" style="BACKGROUND-COLOR: #fafafa;" >
        px&nbsp;(135-170px左右)&nbsp;&nbsp; 高度 
        <input type=text name='s_picheight' value="<%=s_picheight%>" size="5">
        px(高度通常为110-200px之间，请按您的实际情况修改) 
		<img src='../images/helpmemo.gif' alt="<font color='#ff3300'>前台商品列表小图设置建议：</font><br>由于各个行业的商品图规格差别较大（有的可能要高度大于宽度或宽度大于高度，用户要求商品小图高度/宽度不一），<br>所以在将开网店时建议计划好商品列表小图宽度、高度、列数，方便以后按这个尺寸做产品小图，可以先看系统自带的<br>商品图数据，设置下面参数，看下显示效果。如果下面设显示4列，小图宽度设为170，显示效果较佳。如果5列，小图<br>宽度为135，显示效果较佳。">		</td>
    </tr>
	
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'>商品列表中是否显示商品号</td>
      <td> <input type="radio" name="s_model" value="是" <%if s_model="是" then response.write "checked"%> >
        是&nbsp;&nbsp;&nbsp; <input type="radio" name="s_model" value="否" <%if s_model="否" then response.write "checked"%> >
        否 </td>
    </tr>
	
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'>商品详细中是否显示库存数</td>
      <td> <input type="radio" name="s_kuchu" value="是" <%if s_kuchu="是" then response.write "checked"%> >
        是&nbsp;&nbsp;&nbsp; <input type="radio" name="s_kuchu" value="否" <%if s_kuchu="否" then response.write "checked"%> >
        否 </td>
    </tr>



    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'>商品列表是否显示品牌图标</td>
      <td> <input type="radio" name="s_pinpai" value="是" <%if s_pinpai="是" then response.write "checked"%> >
        是&nbsp;&nbsp;&nbsp; <input type="radio" name="s_pinpai" value="否" <%if s_pinpai="否" then response.write "checked"%> >
        否&nbsp;&nbsp; <input type="hidden" name="s_pinpai_l_r" value="左"> &nbsp;&nbsp; 
        图标宽度 
        <input type=text name='s_pinpaiw' value="<%=s_pinpaiw%>" size="3">
        px&nbsp;&nbsp; 高度 
        <input type=text name='s_pinpaih' value="<%=s_pinpaih%>" size="3">
        px  <img src='../images/helpmemo.gif' alt="通常每行排两个，尺寸90px X 40px，按实际情况修改<br><img src='images/pinpai1.gif'><br>如果要每行只排一个，
		图片宽度为120至185，185px最佳<br><img src='images/pinpai2.gif'>"></td>
    </tr>
	
	
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'>主页点击率排行</td>
      <td> <!--<input type="radio" name="s_hit" value="是" <%if s_hit="是" then response.write "checked"%> >
        是&nbsp;&nbsp;&nbsp; <input type="radio" name="s_hit" value="否" <%if s_hit="否" then response.write "checked"%> >
        否&nbsp;&nbsp;&nbsp; -->显示前面 
        <input type=text name='s_hitnum' value="<%=s_hitnum%>" size="4">
        个商品 </td>
    </tr>
	<!--
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'>主页是否显示在线调查</td>
      <td> <input type="radio" name="s_diaocha" value="是" <%if s_diaocha="是" then response.write "checked"%> >
        是&nbsp;&nbsp;&nbsp; <input type="radio" name="s_diaocha" value="否" <%if s_diaocha="否" then response.write "checked"%> >
        否 </td>
    </tr>
	-->
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='23' align='right'>主页留言显示条数</td>
      <td> <!--<input type="radio" name="s_bbs" value="是" <%if s_bbs="是" then response.write "checked"%> >
        是&nbsp;&nbsp;&nbsp; <input type="radio" name="s_bbs" value="否" <%if s_bbs="否" then response.write "checked"%> >
        否&nbsp;&nbsp;&nbsp; -->显示最新 
        <input type=text name='s_mes' value="<%=s_mes%>" size="4">
        条</td>
    </tr>
    <tr bgcolor='#Ffffff' onMouseOver="this.style.background='#FFE4CA';" onMouseOut="this.style.background='#ffffff';"> 
      <td height='46' align='right'>&nbsp;</td>
      <td><input type="submit" name="Submit2" value=" 保存 "> &nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="reset" name="Submit3" value=" 重置 "> </td>
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

