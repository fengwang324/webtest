<!--#include file="conn.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="i.css" type=text/css rel=stylesheet>
<style type="text/css">
<!--

.style1 {font-size: 14px}
.style12 {color: #990000}
.style11 {
	font-size: 14px;
	color: #666666;
	font-weight: bold;
}

.colorcss{
BORDER: #666666 1px solid; BACKGROUND-COLOR: #eeeee1;width:45px; height:23px; padding:0px; margin-bottom: 5px; FONT-SIZE: 10pt;
}

.psizecss{
BORDER: #666666 1px solid; BACKGROUND-COLOR: #eeeee1;width:45px; height:23px; padding:0px; margin-bottom: 5px; FONT-SIZE: 10pt;
}
-->
</style>
<style type="text/css">
 .Numinput{padding:0 20px 0 0;position:relative;height:20px;}
 .Numinput input{font-size:12px;width:24px;height:15px;line-height:15px;}
 
 .Numinput .numadjust{position:absolute;width:18px;height:9px;right:1px;overflow:hidden;background-image:url(images/numadjust.gif);background-repeat:no-repeat;cursor:pointer;}
 
 
 .Numinput .numadjust.increase{background-position:0 0;top:0;}
 .Numinput .numadjust.increase.active{background-position:0 -20px;}
 .Numinput .numadjust.decrease{background-position:0 -10px;bottom:0;}
 .Numinput .numadjust.decrease.active{background-position:0 -30px;}

</style>

<script src="CJL.0.1.min.js"></script>
<script src="ImageZoom.js"></script>
<script>
ImageZoom._MODE = {
	//跟随
	"follow": {
		methods: {
			init: function() {
				this._stylesFollow = null;//备份样式
				this._repairFollowLeft = 0;//修正坐标left
				this._repairFollowTop = 0;//修正坐标top
			},
			load: function() {
				var viewer = this._viewer, style = viewer.style, styles;
				this._stylesFollow = {
					left: style.left, top: style.top, position: style.position
				};
				viewer.style.position = "absolute";
				//获取修正参数
				if ( !viewer.offsetWidth ) {//隐藏
					styles = { display: style.display, visibility: style.visibility };
					$$D.setStyle( viewer, { display: "block", visibility: "hidden" });
				}
				//修正中心位置
				this._repairFollowLeft = viewer.offsetWidth / 2;
				this._repairFollowTop = viewer.offsetHeight / 2;
				//修正offsetParent位置
				if ( !/BODY|HTML/.test( viewer.offsetParent.nodeName ) ) {
					var parent = viewer.offsetParent, rect = $$D.rect( parent );
					this._repairFollowLeft += rect.left + parent.clientLeft;
					this._repairFollowTop += rect.top + parent.clientTop;
				}
				if ( styles ) { $$D.setStyle( viewer, styles ); }
			},
			repair: function(e, pos) {
				var zoom = this._zoom,
					viewerWidth = this._viewerWidth,
					viewerHeight = this._viewerHeight;
				pos.left = ( viewerWidth / 2 - pos.left ) * ( viewerWidth / zoom.width - 1 );
				pos.top = ( viewerHeight / 2 - pos.top ) * ( viewerHeight / zoom.height - 1 );
			},
			move: function(e) {
				var style = this._viewer.style;
				style.left = e.pageX - this._repairFollowLeft + "px";
				style.top = e.pageY - this._repairFollowTop + "px";
			},
			dispose: function() {
				$$D.setStyle( this._viewer, this._stylesFollow );
			}
		}
	},
	//拖柄
	"handle": {
		options: {//默认值
			handle:		""//拖柄对象
    	},
		methods: {
			init: function() {
				var handle = $$( this.options.handle );
				if ( !handle ) {//没有定义的话用复制显示框代替
					var body = document.body;
					handle = body.insertBefore(this._viewer.cloneNode(false), body.childNodes[0]);
					handle.id = "";
					handle["_createbyhandle"] = true;//生成标识用于移除
				} else {
					var style = handle.style;
					this._stylesHandle = {
						left: style.left, top: style.top, position: style.position,
						display: style.display, visibility: style.visibility,
						padding: style.padding, margin: style.margin,
						width: style.width, height: style.height
					};
				}
				$$D.setStyle( handle, { padding: 0, margin: 0, display: "none" } );
				
				this._handle = handle;
				this._repairHandleLeft = 0;//修正坐标left
				this._repairHandleTop = 0;//修正坐标top
			},
			load: function() {
				var handle = this._handle, rect = this._rect;
				$$D.setStyle( handle, {
					position: "absolute",
					width: this._rangeWidth + "px",
					height: this._rangeHeight + "px",
					display: "block",
					visibility: "hidden"
				});
				//获取修正参数
				this._repairHandleLeft = rect.left + this._repairLeft - handle.clientLeft;
				this._repairHandleTop = rect.top + this._repairTop - handle.clientTop;
				//修正offsetParent位置
				if ( !/BODY|HTML/.test( handle.offsetParent.nodeName ) ) {
					var parent = handle.offsetParent, rect = $$D.rect( parent );
					this._repairHandleLeft -= rect.left + parent.clientLeft;
					this._repairHandleTop -= rect.top + parent.clientTop;
				}
				//隐藏
				$$D.setStyle( handle, { display: "none", visibility: "visible" });
			},
			start: function() {
				this._handle.style.display = "block";
			},
			move: function(e, x, y) {
				var style = this._handle.style, scale = this._scale;
				style.left = Math.ceil( this._repairHandleLeft - x / scale ) + "px";
				style.top = Math.ceil( this._repairHandleTop - y / scale )  + "px";
			},
			end: function() {
				this._handle.style.display = "none";
			},
			dispose: function() {
				if( "_createbyhandle" in this._handle ){
					document.body.removeChild( this._handle );
				} else {
					$$D.setStyle( this._handle, this._stylesHandle );
				}
				this._handle = null;
			}
		}
	},
	//切割
	"cropper": {
		options: {//默认值
			opacity:	.5//透明度
    	},
		methods: {
			init: function() {
				var body = document.body,
					cropper = body.insertBefore(document.createElement("img"), body.childNodes[0]);
				cropper.style.display = "none";
				
				this._cropper = cropper;
				this.opacity = this.options.opacity;
			},
			load: function() {
				var cropper = this._cropper, image = this._image, rect = this._rect;
				cropper.src = image.src;
				cropper.width = image.width;
				cropper.height = image.height;
				$$D.setStyle( cropper, {
					position: "absolute",
					left: rect.left + this._repairLeft + "px",
					top: rect.top + this._repairTop + "px"
				});
			},
			start: function() {
				this._cropper.style.display = "block";
				$$D.setStyle( this._image, "opacity", this.opacity );
			},
			move: function(e, x, y) {
				var w = this._rangeWidth, h = this._rangeHeight, scale = this._scale;
				x = Math.ceil( -x / scale ); y = Math.ceil( -y / scale );
				this._cropper.style.clip = "rect(" + y + "px " + (x + w) + "px " + (y + h) + "px " + x + "px)";
			},
			end: function() {
				$$D.setStyle( this._image, "opacity", 1 );
				this._cropper.style.display = "none";
			},
			dispose: function() {
				$$D.setStyle( this._image, "opacity", 1 );
				document.body.removeChild( this._cropper );
				this._cropper = null;
			}
		}
	}
}

ImageZoom.prototype._initialize = (function(){
	var init = ImageZoom.prototype._initialize,
		mode = ImageZoom._MODE,
		modes = {
			"follow": [ mode.follow ],
			"handle": [ mode.handle ],
			"cropper": [ mode.cropper ],
			"handle-cropper": [ mode.handle, mode.cropper ]
		};
	return function(){
		var options = arguments[2];
		if ( options && options.mode && modes[ options.mode ] ) {
			$$A.forEach( modes[ options.mode ], function( mode ){
				//扩展options
				$$.extend( options, mode.options, false );
				//扩展钩子
				$$A.forEach( mode.methods, function( method, name ){
					$$CE.addEvent( this, name, method );
				}, this );
			}, this );
		}
		init.apply( this, arguments );
	}
})();

</script>
<style type="text/css">
.list{ padding-right:4px;}
.list img{width:60px; cursor:pointer; padding:1px; border:1px solid #cdcdcd; margin-bottom:10px; display:block;}
.list img.onzoom, .list img.on{padding:1px; border:1px solid #336699;}

.container{ position:relative;}

.izImage2{ width:350px; border:1px solid #cccccc; z-index:80;}
.izViewer2{width:300px;height:280px;position:absolute;left:360px;top:0; display:none; border:1px solid #999;} /* 放大的框 */

.handle2{display:none;opacity:0.6;filter:alpha(opacity=60);background:#E6EAF3; cursor:crosshair;}
</style>	
</head>

<%pkid=request("pkid")
sql="select * from view_product where pkid = "&pkid

set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
		productname=rs("productname")
rs.close
%>

<title><%=productname%> - 美国德玉堂 SUNSNU.com </title>
<META  name="keywords" content="<%=productname%>">
<META  name=description content="美国德玉堂中草药肿瘤研究中心，美国德玉堂，孙德玉医师，美国灵草百消丹，抗肿瘤中草药，抗衰老中药，中医养生保健。">
<body>
<!-- #include file="head.asp" -->
<TABLE width="1020" border=0 align="center" cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
  <tr>
    <td> 



<table width="1020" height="352" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
    <td width="212" height="150" valign="top">
		<!-- #include file="inc_left_all.asp" -->
      </td>
    <td width="748" valign="top">
	<table width="99%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height=2></td>
      </tr>
    </table>
            <TABLE width="780" border=0 align="center" cellPadding=0 cellSpacing=2 class=a1>
              <TBODY>
          <TR>
            <TD height=30 colSpan=3><TABLE width="100%" border=0 cellPadding=0 cellSpacing=0 background="images/bg4.gif">
<TBODY>
                <TR>
                  <TD width=36 
                            height=30 background=images/hotnewpro.gif>　　　</TD>
                  <TD width="629"><span ><a href="index.asp">首頁</a> - <a href="productlist.asp">商品列表</a> - 商品詳細：</span></TD>
                  <TD width=71><a href="javascript:window.history.back()">&lt;&lt; 返回</a></TD>
                </TR>
              </TBODY>
            </TABLE></TD>
          </TR>
		  </tbody>
		  </table>
            <!--<TABLE width="780" 
border=0 align="center" cellPadding=0 cellSpacing=0>
              <TBODY>
                <TR> 
                  <TD width=547 
                            height=50 background=images/buystep0.jpg> 
<table width="538" border="0" cellspacing="0" cellpadding="0">
<tr> 
                        <td width="193" height="22">&nbsp;</td>
                        <td width="113"> 
                          <div align="center"><strong><font color="#FFFFFF">浏览商品</font></strong></div></td>
                        <td width="112"> 
                          <div align="center">购物车</div></td>
                        <td width="120"> 
                          <div align="center">结算</div></td>
                      </tr>
                  </table></TD>
                  <TD>&nbsp;</TD>
                  <TD width=59>&nbsp;</TD>
                </TR>
              </TBODY>
            </TABLE>-->
            <%
dim pkid,model,productname,smallpicpath,smallpicpath2,smallpicpath3,smallpicpath4,smallpicpath5,price1,price2,pipai

pkid=request("pkid")
sql="select * from view_product where pkid = "&pkid

set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof or rs.eof then
	response.write "没有商品记录！"
else

		pkid=rs("pkid")
		model=rs("model")
		productname=rs("productname")
		smallpicpath=rs("smallpicpath")
		bigpicpath=rs("bigpicpath")
		bigpicpath2=rs("bigpicpath2")
		bigpicpath3=rs("bigpicpath3")
		bigpicpath4=rs("bigpicpath4")
		bigpicpath5=rs("bigpicpath5")
	
		price1=rs("price1")
		price2=rs("price"&session("customkind"))
		pipai=rs("pipai")
		kindname=rs("kindname")
		disc=rs("disc")
		color=rs("color")
		psize=rs("psize")
		hit=rs("hit")
		saleshu=rs("saleshu")
		addtime=rs("addtime")
		kuchushu=rs("kuchushu")
		oneweight=rs("oneweight")
		punit=rs("punit")
		selt1=rs("selt1")
		
		allpic=bigpicpath&"|"&bigpicpath2&"|"&bigpicpath3&"|"&bigpicpath4&"|"&bigpicpath5
		
		if disc<>"" then
			'disc = replace(disc,vbcrlf,"<br>")
			'disc = replace(disc,"shop/","")    '在这里设置过滤文件夹
			'disc = replace(disc," ","&nbsp;")
			'disc = replace(disc,"IMG&nbsp;src=","IMG src=")
			'disc = replace(disc,"&nbsp;border="," border=")
		end if
		'---------更新点击次数begin---------
		conn.execute("update e_product set hit=hit+1 where pkid="&pkid)
		'---------更新点击次数end-----------
		
		'---------记录浏览过的商品begin-----
		haveshow=request.cookies("haveshow")
		haveshow="," & haveshow & ","
		haveshow = replace(haveshow,","& pkid &",","")
		haveshow = pkid &","& haveshow
		haveshow = replace(haveshow,",,",",")
		haveshow = replace(haveshow,",,",",")
		response.cookies("haveshow")=left(haveshow,150)
		response.cookies("haveshow").Expires="2020-12-31"
		'response.write haveshow
		'---------记录浏览过的商品end--------
%>

		    <TABLE width=780 border=0 align="center" cellPadding=2 cellSpacing=0><TBODY>
              <TR> 
                <TD colSpan=2 height="8"></TD>
              </TR>
              <TR>
                <TD height="50" colspan="2" align=center><font style="font-size:24px; font-weight:bold; color:#F60"><%=productname%></font></TD>
                </TR>
              <TR> 
                <TD width=420 height="30" align=left vAlign=top><table>
                  <tr>
		<td valign="top"><div id="idList" class="list"> </div></td>
		<td valign="top">
			<div class="container" style="z-index:6"> 
		        <img id="idImage2" class="izImage2" />
				<div id="idViewer2" class="izViewer2"></div> <!-- 放大的框 -->
			</div>
		</td>
	</tr>
</table>
<div id="idHandle3" class="handle2"  style="z-index:99"></div> <!-- 放大的图 -->




				  
				  </TD>
                <TD width="320" align=left vAlign=top style="padding-top:5px;padding-left:3px;"> 
				<TABLE width=315 border=0 cellPadding=0 cellSpacing=0 style="BORDER-RIGHT: #cccccc 0px solid; BORDER-TOP: #cccccc 0px solid; BORDER-LEFT: #cccccc 0px solid; BORDER-BOTTOM: #cccccc 0px solid">
					<TBODY>
                      <TR align=left> 
                        <TD height="95" valign="top" style="PADDING-LEFT: 8px;PADDING-top: 22px; line-height:180%;"> 
                          商品類別：<B><%=kindname%></B><br>
						  商品名稱：<font style="font-size:14px;color:#CC6600"><%=productname%></font><br>
                          商品編號：<b><%=model%></b> <br>

						  商品單位：<b><%=punit%></b><BR>
						  <%if cstr(oneweight)>"0" then%>
                          商品重量：<B><%=oneweight%></B> 克<BR>
						  <%end if%>
						  <%if s_kuchu="是" then%>
                          庫存數量：<B><%=kuchushu%></B><BR>
						  <%end if%>
						  瀏覽次數：<B><%=hit%></B><BR>
						  銷售數量：<B><%=saleshu%></B><BR>
						  上市時間：<%=addtime%><BR>
                          </TD>
                      </TR>
                      <TR align=middle> 
                        <TD height="13"><HR 
                  style="BORDER-RIGHT: #999999 1px dotted; BORDER-TOP: #999999 1px dotted; BORDER-LEFT: #999999 1px dotted; BORDER-BOTTOM: #999999 1px dotted" width=300 noShade SIZE=1> 
                        </TD>
                      </TR>
                      <TR align=left> 
                        <TD height="50" vAlign=center style="PADDING-LEFT: 8px; COLOR: #666666; "> 
                          <font style="width:80px;">市場價：</font><FONT class='oldprice'><del>＄<%=price1%></del></FONT> <BR>
						  <FONT color=#ff0000 style="width:80px;"><%=session("customkindname")%>價：</font><font class='nowprice'>＄<%=price2%></font></TD>
                      </TR>
                      <TR align=middle> 
                        <TD height="13"><HR 
                  style="BORDER-RIGHT: #999999 1px dotted; BORDER-TOP: #999999 1px dotted; BORDER-LEFT: #999999 1px dotted; BORDER-BOTTOM: #999999 1px dotted" width=300 noShade SIZE=1> 
                        </TD>
                      </TR>
					  </TBODY>
				    </table>
				<TABLE width=315 border=0 cellPadding=0 cellSpacing=0 style="padding:4px;BORDER-RIGHT: #dddddd 1px solid; BORDER-TOP: #dddddd 1px solid; BORDER-LEFT: #dddddd 1px solid; BORDER-BOTTOM: #dddddd 1px solid;BACKGROUND-COLOR: #fcfcfc;">
					<TBODY>
                    <form name="formshow" action="?tk=shop7z" method="post" onSubmit="return Juge()">
					<%
					if s_colorsize="是" then
						if color<>"" then
					%>
					<%
					  	end if
						if psize<>"" then
					%>
					<%
						end if
						if color<>"" or  psize<>"" then
					%>
					<%
						end if
					end if
					%>
                      <TR align=left> 
                        <TD width="100%" height="41" vAlign=center style="PADDING-LEFT: 8px; COLOR: #666666">
						<table border=0 cellPadding=0 cellSpacing=0><tr><td>
                        <div class="Numinput" style="z-index:3">
						購買數量：
						<input type="text" name="num" id=num size="7" value=1 style="WIDTH: 40px;height:21px;font-weight:bold; font-size:12pt; padding-top:2px;padding-left:3px;background-color:#FFFFFF;FONT-FAMILY: 'Century Gothic';"/>
						<span  class="numadjust increase" onClick="document.all.num.value=parseFloat(document.all.num.value)+1"></span>
						<span  class="numadjust decrease" onClick="if(parseFloat(document.all.num.value)>1){document.all.num.value=parseFloat(document.all.num.value)-1}"></span>						</div>
						</td>
						<td>&nbsp;&nbsp;
						<INPUT type="hidden" value="<%=request("pkid")%>" name="pkid"> 
                        <input type="image" name="Submit" value="Submit" src="images/gouwu.gif" style="height:27px;width:108x;BORDER:0px; vertical-align:top;; ">						</td></tr></table>                        </TD>
                      </TR>
                    </form>
				    </table><!-- Baidu Button BEGIN -->
<div id="bdshare" class="bdshare_t bds_tools_32 get-codes-bdshare">
<a class="bds_qzone"></a>
<a class="bds_tsina"></a>
<a class="bds_tqq"></a>
<a class="bds_renren"></a>
<a class="bds_t163"></a>
<span class="bds_more"></span>
<a class="shareCount"></a>
</div>
<script type="text/javascript" id="bdshare_js" data="type=tools&amp;uid=0" ></script>
<script type="text/javascript" id="bdshell_js"></script>
<script type="text/javascript">
document.getElementById("bdshell_js").src = "http://bdimg.share.baidu.com/static/js/shell_v2.js?cdnversion=" + Math.ceil(new Date()/3600000)
</script>
<!-- Baidu Button END -->
				<TABLE width=315 border=0 cellPadding=0 cellSpacing=0 style="BORDER-RIGHT: #cccccc 0px solid; BORDER-TOP: #cccccc 0px solid; BORDER-LEFT: #cccccc 0px solid; BORDER-BOTTOM: #cccccc 0px solid">
					<div style="width:310px;text-align:left;"><a href='zoom.asp?pkid=<%=pkid%>&allpic=<%=Server.URLEncode(allpic)%>&productname=<%=Server.URLEncode(productname)%>' target='_blank'><IMG src="images/zoom.gif" border=0 align="absmiddle"></a>&nbsp;&nbsp <a href="javascript:window.external.AddFavorite('http://#<%=request.ServerVariables("SCRIPT_NAME")%>?<%=request.serverVariables("Query_string")%>', '<%=application("sitename")%>-<%=productname%>')"><img src="images/button-shoucang.gif" border="0"></a></div>
                  </TABLE></TD>
              </TR>
              <TR> 
                <TD colSpan=2 height="3"></TD>
              </TR>
              <TR> 
                <TD style="BACKGROUND-REPEAT: repeat-x" background=images/talbe_2_back.jpg colSpan=2 height=1></TD>
              </TR>
			  </TBODY>
			  </table>

<script language="JavaScript">
<!--
function Juge()
{
<%if s_colorsize="是" and color<>"" then%>
	if (formshow.color.value == "")
	{
		alert("請選擇顏色!");
		document.formshow.color.focus();
		return false;
	}
<%end if
if s_colorsize="是" and psize<>"" then%>
	if (formshow.psize.value == "")
	{
		alert("請選擇尺碼!");
		document.formshow.psize.focus();
		return false;
	}
<%end if%>
}
//-->
</script>
<%  action=request("tk")
if action="shop7z" then 
pkid=request("pkid")
num=request("num")
if isnumeric(num) and not isnull(num) then
	num=abs(num) 
else
	num="1"
end if
if request.Cookies("notemp")="" then
	mytime=now()
	myyear=right(year(mytime),2)
	mymonth=month(mytime)
	myday=day(mytime)
	myhour=hour(mytime)
	myminute=minute(mytime)
	mysec=second(mytime)
	if cint(mymonth)<10 then
		mymonth="0"&mymonth
	end if
	if cint(myday)<10 then
		myday="0"&myday
	end if
	dim num1
	dim rndnum
	Randomize
	Do While Len(rndnum)<4
		num1=CStr(Chr((57-48)*rnd+48))
		rndnum=rndnum&num1
	loop
	billno=myyear&mymonth&myday&myhour&myminute&mysec&"-"&rndnum
	response.Cookies("notemp")=billno
end if
if pkid<>"" then
	sql="insert into x_order(notemp,pkid,num) values('"&request.Cookies("notemp")&"','"&pkid&"','"&num&"')"
	conn.execute(sql)
end if
conn.close
set conn=nothing
response.Redirect("order_all.asp")
End if 
%>
	
<!--以下详细信息及评论-->
			<br>
		    <TABLE width=780 border=0 align="center" cellPadding=1 cellSpacing=0>
              <TR> 
                <TD colSpan=2 height="5"></TD>
              </TR>
              <TR> 
                <TD colSpan=2 height=20> 
				
                  <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0 style="BORDER: #cccccc 1px solid;">
                      <TR> 
                        <TD  valign="top">
							 <TABLE cellSpacing=0 cellPadding=0 width="740" border=0 bgcolor="#eeeeee">
								  <TR style="text-align:center;"> 
									<TD height=27 width=100 bgcolor="#ffffff" id=but1  style="border-right:1px solid #cccccc"><A HREF="javascript:void(0)" onClick="OnClickColor('1');"><b>商品詳細</b></a></TD>
									<TD width=100 id=but2><A HREF="javascript:void(0)" onClick="OnClickColor('2');"><b>商品評論</b></a></TD>
									<TD width=100 id=but3><A HREF="javascript:void(0)" onClick="OnClickColor('3');"><b>發表評論</b></a></TD>
									<TD width=440>&nbsp;</TD>
								  </TR>
							  </TABLE>
								<SCRIPT LANGUAGE="JavaScript">
								var previous = "1"; 
								
								function OnClickColor(eleName) 
								{  
								  if(previous  != "" && document.getElementById("but"+previous) != null){  
										document.getElementById("but"+previous).style.background = "#eeeeee";  
										document.getElementById("tab"+previous).style.display="none"; 
									} 
								  document.getElementById("but"+eleName).style.background = "#ffffff"; 
								  document.getElementById("but"+eleName).style.borderRight="1px solid #cccccc";
								  document.getElementById("tab"+eleName).style.display="block";
								  
								  previous  = eleName; 
								} 
								
								</SCRIPT>


							 <!-- 详细begin -->
							 <TABLE cellSpacing=0 cellPadding=10 width="740" height="240" border=0 align="center"  id=tab1>
								  <TR> 
									<TD align=left style="WIDTH: 700; WORD-WRAP:break-word;" valign="top"><br>
<%=disc%> <p>&nbsp;</p></TD>
								  </TR>
							  </TABLE>
							<!-- 详细end -->


							<!-- 评论begin -->
							<table width="100%" border="0" cellpadding="4" cellspacing="1"  id=tab2 style='display:none'>
								<tr>
									<td height="230" bgcolor="#FFFFFF"  valign="top"> 
									<table width="100%" border="0" cellspacing="0" cellpadding="1">
									<tr>
											<td><strong>商品評論：</strong></td>
									</tr>
									</table>
									<%
									dim mr
										sql="select * from z_pijia where productid="&request.querystring("pkid")&" and check='1' order by id desc"
										set rsp=server.CreateObject("adodb.recordset")
										rsp.open sql,conn,1,1
										if rsp.bof or rsp.eof then
											response.write "<div align=left>暫時沒有此產品的評價記錄！</div>"
										else
											mr=1
											do while not rsp.eof
												set yourname=rsp("yourname")
												set tel=rsp("tel")
												set memo=rsp("memo")
												set rememo=rsp("rememo")
												set addtime=rsp("addtime")
												if memo<>"" then
												memo = replace(memo,vbcrlf,"<br>")
												memo = replace(memo," ","&nbsp;")
												end if
												if rememo<>"" then
												rememo = replace(rememo,vbcrlf,"<br>")
												rememo = replace(rememo," ","&nbsp;")
												end if
									%>
								  
										<table width="100%" border="0" cellspacing="2" cellpadding="0" class="border" >
										  <tr> 
											<td width="17%" height="23" bgcolor="#f7f7f7"> <div align="right"> 姓名：</div></td>
											<td width="26%"><%=yourname%>&nbsp; </td>
											<td width="17%" bgcolor="#f7f7f7"> <div align="right">聯繫方式：</div></td>
											<td width="40%"><%=tel%></td>
										  </tr>
										  <tr> 
											<td height="23" bgcolor="#f7f7f7"> <div align="right">評論內容：</div></td>
											<td colspan="3" style="WIDTH: 450; WORD-WRAP: break-word"><%=memo%>&nbsp;[<%=addtime%>]</td>
										  </tr>
										  <tr> 
											<td height="25" bgcolor="#f7f7f7"> <div align="right"><font color="#CC0033">【管理員回復】：</font></div></td>
											<td colspan="3"  style="WIDTH: 450; WORD-WRAP: break-word"><%=rememo%>&nbsp;</td>
										  </tr>
										</table>
										<table width="100%" border="0" cellspacing="0">
										  <tr> 
											<td height="5"></td>
										  </tr>
										</table>
						  		<%
									mr=mr+1
									rsp.movenext
									loop
								end if
								rsp.close
								set rsp=nothing
								session("pijia")="1"
						  		%>  
				
								
								</td>
								</tr>
							</table>
							<!-- 评论end -->

							<!-- 发表评论begin -->
							<table width="100%" border="0" cellpadding="4" cellspacing="1"  id=tab3 style='display:none'>
							  <form name="Form2" method="post" action="pijia.asp" onSubmit="return Jugepijia()">
								<tr> 
									<td height="23"><strong>發表評論：</strong></td>
								</tr>
								<tr> 
								<td width="17%" height="28"> <div align="right"> 
								<input type="hidden" name="productid" value="<%=pkid%>">姓名：</div></td>
								  <td width="83%"> <input type="text" name="yourname"></td>
								</tr>
								<tr> 
									<td height="28"> <div align="right">聯繫方式：</div></td>
								  <td><input name="tel" type="text" size="30">(可以是電話、email、qq等)</td>
								</tr>
								<tr> 
								<td><div align="right">評論內容：</div></td>
								  <td><textarea name="memo" cols="50" rows="5"></textarea></td>
								</tr>
								<tr> 
								<td height="30">&nbsp;</td>
								  <td><input type="submit" name="Submit2" value="提交評論" class='input3'></td>
								</tr>
							  </form>
							</table>
							<script language="JavaScript">
							<!--
							function Jugepijia()
							{
							  if (Form2.memo.value == "")
							  {
								alert("請填寫評論內容!");
								document.Form2.memo.focus();
								return false;
							  }
							}
							//-->
							</script>
						  <!-- 发表评论end -->
						
						</TD>
                      </TR>
                  </TABLE>
				  
				  
                </TD>
              </TR>
            </TABLE>
		 <%	
		end if
		rs.close
		set rs=nothing

		 %>


   </td>
  </tr>
</table>


    </td>
  </tr>

<TR><TD>

</TD></TR>
</table>

<script language="JavaScript">
(function(){
var iz = new ImageZoom( "idImage2", "idViewer2", {
	mode: "handle", handle: "idHandle3", scale: 1.714, delay: 0   //scale: 1.6放大倍数 上传源图600 大图350    600/350=1.714
});

var arrPic = [], list = $$("idList"), image = $$("idImage2"); 

arrPic.push({ smallPic: "product/<%=bigpicpath%>", originPic: "product/<%=bigpicpath%>", zoomPic: "product/<%=bigpicpath%>" }); 
<%if bigpicpath2<>"" then%>
arrPic.push({ smallPic: "product/<%=bigpicpath2%>", originPic: "product/<%=bigpicpath2%>", zoomPic: "product/<%=bigpicpath2%>" }); 
<%end if
if bigpicpath3<>"" then%>
arrPic.push({ smallPic: "product/<%=bigpicpath3%>", originPic: "product/<%=bigpicpath3%>", zoomPic: "product/<%=bigpicpath3%>" }); 
<%end if
if bigpicpath4<>"" then%>
arrPic.push({ smallPic: "product/<%=bigpicpath4%>", originPic: "product/<%=bigpicpath4%>", zoomPic: "product/<%=bigpicpath4%>" }); 
<%end if
if bigpicpath5<>"" then%>
arrPic.push({ smallPic: "product/<%=bigpicpath5%>", originPic: "product/<%=bigpicpath5%>", zoomPic: "product/<%=bigpicpath5%>" }); 
<%end if%>

$$A.forEach(arrPic, function(o, i){
	var img = list.appendChild(document.createElement("img"));
	img.src = o.smallPic;
	img.onclick = function(){
		iz.reset({ originPic: o.originPic, zoomPic: o.zoomPic });
		$$A.forEach(list.getElementsByTagName("img"), function(img){  img.className = ""; });
		img.className = "onzoom";
	}
	
	var temp;
	img.onmouseover = function(){ if( !this.className ){ this.className = "on"; temp = image.src; image.src = o.originPic; } }
	img.onmouseout = function(){ if( this.className == "on" ){ this.className = ""; image.src = temp; } }
	
	if(!i){ img.onclick(); }
})
})()
</script>
<center>
  <a href="show.asp?pkid=4850" target="_blank"><img src="images/ad1.jpg" alt="美國靈草百消丹" width="1020" vspace="8" border="0" align="middle"></a>
</center>

<!-- #include file="foot.asp" -->
<%
conn.close
set conn=nothing
%>
</body>
</html>

