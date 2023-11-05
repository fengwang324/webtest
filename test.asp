<!--#include file="conn.asp"-->
<%
'request.ServerVariables("HTTP_ACCEPT")),"text/vnd.wap.wml")>0 then
'response.redirect  "" '如果是手机访问则跳转WAP页面
'response.end
'else
'response.redirect  "index.asp" '如果电脑访问跳转到首页
'response.end 
'end if
%>
<!--#include file="inc_product_list.asp"-->
<%
dim rs,sql
dim c_id,c_title,detail
dim sql_new,rs_new,pkid,model,productname,smallpicpath,price1,price2,pipai
dim sitekeyword
%>
<%
if s_head="head4.asp" or s_productkind="4" then
	response.write "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">"
else
	response.write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">"
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%
dim sql4,rs4,meta_all
dim sitename,sitedisc,sitetel
dim hotflag,cxflag
sql4="select sitename,sitedisc,sitetel,sitekeyword from siteinfo"
set rs4=server.createobject("adodb.recordset")
rs4.open sql4,conn,1,1
if rs4.bof or rs4.eof then
	
else
	sitename=rs4("sitename")
	sitedisc=rs4("sitedisc")
	sitetel=rs4("sitetel")
	sitekeyword = rs4("sitekeyword")&"&nbsp;"
	application("sitename")=sitename
	application("sitedisc")=sitedisc
	application("sitetel")=sitetel
	response.cookies("sitekeyword")= sitekeyword
	response.cookies("sitekeyword").Expires="2027-12-30"
end if
rs4.close
set rs4=nothing



sql4="select meta_a11 from meta"
set rs4=server.createobject("adodb.recordset")
rs4.open sql4,conn,1,1
if rs4.bof or rs4.eof then
	response.write "我的网店"
else
		meta_all=rs4("meta_a11")
		response.write meta_all
end if
rs4.close
set rs4=nothing


sql4="select l_id,showflag from e_left where l_id=25 or l_id=39 "
set rs4=server.createobject("adodb.recordset")
rs4.open sql4,conn,1,1
if rs4.bof or rs4.eof then
	
else
	do while not rs4.eof

		l_id=rs4("l_id")
		if l_id=25 then
			hotflag=rs4("showflag")
		else
			cxflag=rs4("showflag")
		end if
	rs4.movenext
	loop
end if
rs4.close
set rs4=nothing
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="i.css" type=text/css rel=stylesheet>
<link href="index_ad.css" type=text/css rel=stylesheet>

<script language="JavaScript">
<!--
function a(f)
{
  var r = f.rad;
  for(var ii=0; ii<r.length; ii++)
    if(r[ii].checked)
      return true;
 alert("请您选择一个选项。");
 return false;
}
//-->
</script>

<script language="javascript">
<!--
function SetBgColor(Menu)
{
		Menu.style.background="#F3F3F5";
}
function RestoreBgColor(Menu)
{
		Menu.style.background="#ffffff";
}
-->
</script>
<SCRIPT LANGUAGE="JavaScript">
var previous = "1"; 

function OnClickColor(eleName) 
{  
  if(previous  != "" && document.getElementById("but"+previous) != null){ 
        document.getElementById("but"+previous).style.background = "url(images/soft_tab_current.jpg)"; 
		document.getElementById("tab"+previous).style.display="none"; 
    } 
  document.getElementById("but"+eleName).style.background = "url(images/tab_bg.jpg)"; 
  document.getElementById("tab"+eleName).style.display="block";
  
  previous  = eleName; 
} 

</SCRIPT>

</head>

<body >
<!-- #include file="head.asp" -->
<%
dim biaoflag,lefturl,leftlink,righturl,rightlink,bannarflag,bannarurl,bannarlink
dim changeflag,curl1,curl2,curl3,curl4,curl5,clink1,clink2,clink3,clink4,clink5
sql4="select * from ad"
set rs4=server.createobject("adodb.recordset")
rs4.open sql4,conn,1,1
if rs4.bof or rs4.eof then

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

	ctext1=rs4("ctext1")
	ctext2=rs4("ctext2")
	ctext3=rs4("ctext3")
	ctext4=rs4("ctext4")
	ctext5=rs4("ctext5")

	clink1=rs4("clink1")
	clink2=rs4("clink2")
	clink3=rs4("clink3")
	clink4=rs4("clink4")
	clink5=rs4("clink5")
	
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
	
end if 
rs4.close
set rs4=nothing
%>

<script>
        window.g_hb_monitor_st = +new Date();
                window.g_config = {appId: 4, assetsHost: "http://a.tbcdn.cn/", toolbar: false, pageType: "tmall",shopId: "105124646",sellerId: "1744773663",wtId:"2079482315",showShopQrcode:true};
                window.shop_config = {
            "shopId": "105124646",
            "siteId": "2",
            "userId": "1744773663",
            "user_nick": "%E5%86%B3%E8%83%9C%E6%96%B0%E6%95%99%E8%82%B2%E9%9B%86%E5%9B%A2%E6%97%97%E8%88%B0%E5%BA%97",
                            "async_mod": {
                    "url": "/index.asp",
                    "hps": 800
                },
                        "template": {
                "name": "",
                "id": "400791854",
                "design_nick": "美国德玉堂"
            }
        };
        window.shop_config.isView = true;
                window.shop_config.isvStat = {
            nickName: '%E5%86%B3%E8%83%9C%E6%96%B0%E6%95%99%E8%82%B2%E9%9B%86%E5%9B%A2%E6%97%97%E8%88%B0%E5%BA%97',
            userId: '1744773663',
            shopId: '105124646',
            siteId: '2',
            siteCategoryId: '3',
            itemId: '',
            shopStats: '1',
            validatorUrl: 'http://store.taobao.com/tadget/shop_stats.htm',
            templateId: '400791854',
            templateName: ''
        };
                window._poc = window._poc || [];
                window._poc.push(["_trackCustom", "tpl", "new_shop"]);
    </script>
<script src="http://g.tbcdn.cn/??kissy/k/1.3.0/seed-min.js,mui/seed/1.1.11/seed.js,mui/seed-g/1.0.9/seed.js,mui/globalmodule/1.3.19/global-module.js,mui/global/1.2.14/global.js,tm/detail/1.5.5/app.js,tm/detail/1.5.5/page/shop.js"></script>
<link rel="stylesheet" href="images/modules-other-tmall-default-min.css?t=20130725.css"/>
<link rel="stylesheet" href="images/T17pdLFxleXXX7spjX.css"/>
<script src="images/init-async-min.js?t=20140401.js"></script>
<style>
        #content .design-page {
            position: relative;
            z-index: 1000;
            width: 950px;
            margin: 0 auto;
        }
                        #page #content #bd {
                width: 100%
            }
                </style>
            <!-- page_style -->
    <style>
          #content{
background:url(http://img01.taobaocdn.com/bao/uploaded/i3/T1S.WAFv8cXXb1upjX) repeat left 0 ;
}
#hd{
background:url(http://img01.taobaocdn.com/bao/uploaded/i1/T1Z9GFFtRdXXb1upjX.jpg) no-repeat center 0 #FFF;
}

     </style>
<div id="page">
<script>
            TB.use("mod~global", function () {
                if (TB.Global && TB.Global.init) {
                    TB.Global.init();
                }
            });
        </script>
<script>
        if(window.TShop&&window.TShop.ModUtil){
            KISSY.ready(function(){
                TShop.use('mod~ww-hover',function(TS){
                    new TS.TshopPsmShopWwHover({mod:KISSY.DOM.get('.tshop-psm-shop-ww-hover')});
                });
            });
        }
    </script>
<div id="content" class="eshop head-expand tb-shop">
          <div id="bd" >
                          
            <div class="layout grid-m0 J_TLayout" data-widgetid="5976798424" data-componentid="231" data-context="bd" data-prototypeid="231" data-id="5976798424" data-sitecategory="">
	
    	<div class="col-main">
		<div class="main-wrap J_TRegion" data-modules="main" style="overflow:visible;" data-width="b950">
			<div class="J_TModule" data-widgetid="5976798425"  id="shop5976798425"  data-componentid="7206126"  data-spm='110.0.7206126-5976798425'  microscope-data='7206126-5976798425' data-title="★F02全屏海报"  >



<div class="tb-module tshop-um tshop-um-haibao">
<div class="quanping_all">
	
<div class="haibao_l" style="height:480px;">
	<div class="haibao_con" style="height:500px;width:1920px;margin-left:-960px;">
		<div class="lunbo_950 J_TWidget dd" style="height:500px;" data-widget-config="{'nextBtnCls':'next','duration':1.0,'activeTriggerCls':'ks_1','easing':'easeOutStrong','effect':'scrollx','interval':5,'navCls':'ks-switchable-nav','steps':1,'contentCls':'ks-switchable-content','prevBtnCls':'prev','autoplay':true}" data-widget-type="Carousel">
<div class="p_n">
	
		<span class="prev">
			<div class="n_d"></div>
		</span>
		<span class="next">
			<div class="n_d"></div>
		</span>
</div>

			<div class="neirong" style="height:500px;width:1920px;">
				<div class="ks-switchable-content">
				<a style="width:1920px;height:500px;background:url(images/T2ZJYzX.jpg) center center no-repeat;" href="show_foot.asp?c_id=315" target="_blank"></a><a style="width:1920px;height:500px;background:url(images/T2EkMCXD.jpg) center center no-repeat;" href="show.asp?pkid=4850" target="_blank"></a><a style="width:1920px;height:500px;background:url(images/T2ZJYzX.jpg) center center no-repeat;" href="show_foot.asp?c_id=315" target="_blank"></a><a style="width:1920px;height:500px;background:url(images/T2EkMCXD.jpg) center center no-repeat;" href="show.asp?pkid=4850" target="_blank"></a>
						  </div>
								
						</div>
					</div>
				</div>
			</div>
				</div>
	
	
	</div>
</div>
		</div>
	</div>
    </div>

    </div>
</div>
</div>

<!------------中间三个图片广告begin--------------->
<%if mid_flag="1" then%>
<table width="1020" border="0" cellspacing="0" cellpadding="0"  align="center" style="margin-top:10px;">
	<tr> 
	  <td width="33%"  align=left> 
		<%
		if rightpicurl1<>"" then
			response.write "<a href='"&rightpiclink1&"'><IMG  src='"&rightpicurl1&"' width=322 border=0></a> "
		end if
		%>
	  </td>
	  <td width="34%" align=center> 
		<%
		if rightpicurl2<>"" then
			response.write "<a href='"&rightpiclink2&"'><IMG  src='"&rightpicurl2&"' width=322 border=0></a> "
		end if
		%>
	  </td>
	  <td width="33%" align=right> 
		<%
		if rightpicurl3<>"" then
			response.write "<a href='"&rightpiclink3&"'><IMG  src='"&rightpicurl3&"' width=322 border=0></a> "
		end if
		%>
	  </td>
	</tr>
</table>
<%end if%>
<!-------------中间三个图片广告end------------------><!--------------按分类显示begin-------------->
<%
sql="select kindnum,kindname from sh_kind where len(kindnum)=4 and indexshow='1' order by kindnum asc"
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof or rs.eof then
	response.write "<table width=960><tr><td height=22><div align=center>没有记录!</div></td></tr></table>"
else
	do while not rs.eof
	kind_num=rs("kindnum")
	kind_name=rs("kindname")
%>

	<table width="1020" border="0" align="center" cellpadding="2" cellspacing="0" class="kindbg" style="margin-top:10px;">
	  <tr>
		<td width="337" height="30"><font class="kindtext"><%=kind_name%></font></td>
		<td width="675" align="right"><%if kind_num=0002 then%><a href='productlist.asp?kind=00020001'>美容养颜</a>　<a href='productlist.asp?kind=00020002'>护肤保养</a>　<a href='productlist.asp?kind=00020006'>美白产品</a> &nbsp; &nbsp; <%elseif kind_num=0001 then%>　<a href='productlist.asp?kind=00010007'>抗肿瘤</a>　<a href='productlist.asp?kind=00010008'>抗衰老</a>　<a href='productlist.asp?kind=00010009'>强身健体</a> &nbsp; &nbsp; <%end if%></td>
	  </tr>
	</table>
	<table width="1020" border="0" align="center" cellpadding="0" cellspacing="0" class="kindtab">
	  <tr>
		<td height="2" class="kindlinebg"></td>
	  </tr>
	  <tr>
		<td align="right"><a href="productlist.asp?kind=<%=kind_num%>"><img src="images/kindpro_more.jpg"  border="0"></a></td>
	  </tr>
	  <tr>
		<td height="5"></td>
	  </tr>
	  <tr>
		<td>
				<TABLE cellSpacing=0 cellPadding=0 border=0 width="100%">
					<TBODY>
					  <TR> 
						<%'按分类显示
						sql_new="select top "&s_col*s_new_line&" pkid,model,productname,smallpicpath,price1,price"&session("customkind")&",pipai,addtime from view_product where kind like '"&kind_num&"%' and updown='1' order by pkid desc "  
						'response.write sql_new
						'response.end
						set rs_new=server.createobject("adodb.recordset")
						rs_new.open sql_new,conn,1,1
						if rs_new.bof or rs_new.eof then
							response.write "<td>此栏暂时没有商品记录！</td>"
						else
							i=1
							do while not rs_new.eof
							pkid=rs_new("pkid")
							model=rs_new("model")
							productname=rs_new("productname")
							smallpicpath=rs_new("smallpicpath")
							price1=rs_new("price1")
							price2=rs_new("price"&session("customkind"))
							pipai=rs_new("pipai")
							addtime=rs_new("addtime")
						
							call product_list(1)
							
							if i mod s_col =0 then
								response.write "</tr><tr><td height=8>&nbsp;</td></tr><tr>"
								i=1
							else
								i=i+1
							end if
							rs_new.movenext
							loop
						end if
						rs_new.close
						set rs_new=nothing
						%>
					  </TR>
				  </TBODY>
		  </TABLE>
		</td>
	  </tr>
</table>
<%
	rs.movenext
	loop
end if
rs.close
set rs=nothing
%>
<!-------------按分类显示end--------------->

<!-------------下面横广告begin------------->
<%if newpicurl<>"" then%>

<table width="1020" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-top:10px;">
	<TR> 
	  <TD vAlign=center align=middle >
	  <div align="center"><a href='<%=newpiclink%>'><IMG src="<%=newpicurl%>" border="0" width="1020"></a></div>
	  </TD>
	</TR>
</table>
<%end if%>
<!-------------下面横广告end--------------->



 

<%if biaoflag="1" then%>
<!--#include file="ad.asp"-->
<%end if%>


<center>
  <a href="show.asp?pkid=4850" target="_blank"><img src="images/ad1.jpg" alt="美国灵草百消丹" width="1020" vspace="8" border="0" align="middle"></a>
</center>
<table width="1020" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top"><table width="98%" border="0" cellpadding="2" cellspacing="0" class="kindbg" >
      <tr>
        <td width="101"><font class="kindtext">公司介绍</font></td>
        <td width="911"><strong class="nowprice">美国德玉堂肿瘤研究中心</strong> &nbsp; <a href="zj.asp">专家介绍</a> &nbsp; <a href="dbsh.asp">德被四海</a> &nbsp; <a href="jg.asp">保健品加工</a> &nbsp; <a href="jz.asp?c_id=282">客户见证</a> &nbsp; <a href="show_foot.asp?c_id=289">联系我们</a></td>
      </tr>
      <tr class="kindtab">
        <td height="2" colspan="2" class="kindlinebg"></td>
      </tr>
    </table>
      <br>
      <table width="98%" border="0" cellspacing="0" cellpadding="0" >
        <tr>
          <td valign="top"><p><span class="large"> &nbsp; &nbsp; 美国德玉堂肿瘤研究中心是一家集研发、加工、生产和销售为一体的跨国保健品股份有限公司，公司由数十位享誉世界具有40多年丰富临床经验的中草药研究专家创办而成，经美国加利福尼亚州政府批准于2002年注册成立，公司前身是孙氏保健中心。</span><span class="text"><br />
            </span><span class="large"> &nbsp; &nbsp; 孙德玉教授是四代祖传中医，具有40多年肿瘤临床经验，她是应用中草药防癌抗癌的先驱。早在15年前就成功研制出"癌健消"，即现在的"美国灵草百消丹"，因疗效显著，于1996年荣获加拿大保健品金杯奖 。经过十几年的临床使用和改良，应用现代高科技技术，采用纯天然名贵优质中药材精制而成。如今的"美国灵草百消丹"的品种已多达28种系列，针对乳腺、子宫颈、肺、肠、淋巴、鼻咽、胰腺、脑等，不同的病症使用不同的编号，对症选用。一般使用三个疗程左右（90天）症状有显著的改善，"美国灵草百消丹"已遍布美国，远销加拿大、英国、法国、德国、新加坡、澳大利亚、马来西亚、韩国、朝鲜、印尼、菲律宾、泰国、墨西哥、日本、中国大陆、台湾、香港、澳门等国家和地区，所到之处深得消费者的好评和当地多家权威媒体的青睐。</span><a href="zj.asp" class="img"><font color="#FF0000">更多&gt;&gt;&gt;</font></a><br />
          </p></td>
        </tr>
    </table></td>
    <td valign="top" style="padding-top:5px;"><table width="282" border="0" cellspacing="0" cellpadding="0" align=right>
      <tr>
        <td height=1 background="images/newstop.jpg"></td>
      </tr>
      <tr>
        <td height="268" valign="top" background="images/newsmid.jpg"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="24" background="images/tab_bg.jpg" id=but1 style="padding-top: 5px;" onClick="OnClickColor('1');"><div align="center" class="oldprice"><font color="#666666">信息中心</font></div></td>
          </tr>
          <tr>
            <td colspan=3 height=1 bgcolor='#cccccc'></td>
          </tr>
        </table>
          <table width="94%" border="0" align="center" cellpadding="0" cellspacing="0" id=tab1>
            <tr>
              <td height="70"><table width="265" height="47" border=0 cellpadding=1 cellspacing=1 align=center>
                <tbody>
                  <tr>
                    <td><DIV class="line_orange">
                      <UL class=num>
                        <% '商讯
							sql="select top "&s_news&" c_id,c_title,detail from e_contect where c_parent2=30 order by c_num desc,c_addtime desc,c_id desc"
							set rs=server.CreateObject("adodb.recordset")
							rs.open sql,conn,1,1
							if rs.bof or rs.eof then
								response.write "<div align=center>没有记录!</div>"
							else
							
								do while not rs.eof
									set c_id=rs("c_id")
									set c_title=rs("c_title")
									'if len(c_title)>20 then 
									'c_title2=left(c_title,19)&"…"
									'else
									c_title2=c_title
									'end if
									set detail=rs("detail")
								  if detail="1" then
										response.write "<li><NOBR><a href='show_all.asp?c_id="&c_id &"' title='"&c_title&"'>"&c_title2&"</a></NOBR></li>"
								  else
										response.write "<li><NOBR><a href='news.asp?l_id=30' title='"&c_title&"'>"&c_title2&"</a></NOBR></li>"
								  end if
								rs.movenext
								loop
								response.write "<TR>" 
									response.write "<TD height='20' align=right><a href='news.asp?l_id=30' class=txt>更多>></a></TD>"
								response.write "</TR>"
							end if
							rs.close
							set rs=nothing
							%>
                      </UL>
                    </DIV></td>
                  </tr>
                </tbody>
              </table></td>
            </tr>
          </table>
          <table width="94%" border="0" align="center" cellpadding="0" cellspacing="0" id=tab2 style='display:none'>
            <tr>
              <td ></td>
            </tr>
          </table></td>
      </tr>
      <tr>
        <td height=11 background="images/newsbottom.jpg"></td>
      </tr>
    </table></td>
  </tr>
</table>
<br>
<table width="1020" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#d6eda8">
  <tbody>
    <tr>
      <td height="5" align="center" valign="bottom"></td>
    </tr>
    <tr>
      <td height="134" align="center" valign="bottom"><table border="0" cellspacing="0">
        <tbody>
          <tr>
            <td align="center" valign="bottom"><table border="0" cellspacing="4" bgcolor="#FFFFFF" style="BORDER-LEFT: #808e65 1px solid;BORDER-right: #808e65 1px solid;BORDER-top: #808e65 1px solid; BORDER-BOTTOM: #808e65 1px solid">
              <tbody>
                <tr>
                  <td height="105" align="center" valign="top" class="p9"><a href="zj.asp" title=""><img src="images/uploadimages/20103904512588.jpg" alt="" height="136" border="0"><br>
                    生活周刊</a></td>
                </tr>
              </tbody>
            </table>
              <img src="images/bg_photoList.gif" alt="" width="90" height="5"></td>
            <td align="center" valign="bottom"><table border="0" cellspacing="4" bgcolor="#FFFFFF" style="BORDER-LEFT: #808e65 1px solid;BORDER-right: #808e65 1px solid;BORDER-top: #808e65 1px solid; BORDER-BOTTOM: #808e65 1px solid">
              <tbody>
                <tr>
                  <td height="105" align="center" valign="top" class="p9"><a href="dbsh.asp" title=""><img src="images/uploadimages/20103904557826.jpg" alt="德被四海 玉汝于成" height="136" border="0"><br>
                    德被四海</a></td>
                </tr>
              </tbody>
            </table>
              <img src="images/bg_photoList.gif" alt="" width="90" height="5"></td>
            <td height="134" align="center" valign="bottom"><table border="0" cellspacing="4" bgcolor="#FFFFFF" style="BORDER-LEFT: #808e65 1px solid;BORDER-right: #808e65 1px solid;BORDER-top: #808e65 1px solid; BORDER-BOTTOM: #808e65 1px solid">
              <tbody><tr>
                <td height="105" align="center" valign="top" class="p9"><a href="zj.asp" title=""><img src="images/uploadimages/201132113463335.jpg" alt="" width="170" height="136" border="0"><br>
                  2013年获奖</a></td>
              </tr>
            </tbody></table>
                  <img src="images/bg_photoList.gif" width="170" height="5"></td>
            <td align="center" valign="bottom"><table border="0" cellspacing="4" bgcolor="#FFFFFF" style="BORDER-LEFT: #808e65 1px solid;BORDER-right: #808e65 1px solid;BORDER-top: #808e65 1px solid; BORDER-BOTTOM: #808e65 1px solid">
              <tbody><tr>
                <td height="105" align="center" valign="top" class="p9"><a href="zj.asp" title=""><img src="images/uploadimages/001.jpg" alt="" width="190" height="136" border="0"><br>
                  2013年获奖</a></td>
                </tr>
                </tbody></table>
              <img src="images/bg_photoList.gif" width="198" height="5"></td>
            <td align="center" valign="bottom"><table border="0" cellspacing="4" bgcolor="#FFFFFF" style="BORDER-LEFT: #808e65 1px solid;BORDER-right: #808e65 1px solid;BORDER-top: #808e65 1px solid; BORDER-BOTTOM: #808e65 1px solid">
              <tbody>
                <tr>
                  <td height="105" align="center" valign="top" class="p9"><a href="zj.asp" title=""><img src="images/uploadimages/20083719019173.jpg" alt="" height="136" border="0"><br>
                    美国德玉堂</a></td>
                </tr>
              </tbody>
            </table>
              <img src="images/bg_photoList.gif" alt="" width="130" height="5"></td>
            <td align="center" valign="bottom"><table border="0" cellspacing="4" bgcolor="#FFFFFF" style="BORDER-LEFT: #808e65 1px solid;BORDER-right: #808e65 1px solid;BORDER-top: #808e65 1px solid; BORDER-BOTTOM: #808e65 1px solid">
              <tbody><tr>
                <td height="105" align="center" valign="top" class="p9"><a href="zj.asp" title=""><img src="images/uploadimages/20103904641787.jpg" alt="" height="136" border="0"><br>
                  世界周刊</a></td>
                </tr>
                </tbody></table>
              <img src="images/bg_photoList.gif" width="100" height="5"></td>
            <td align="center" valign="bottom"><table border="0" cellspacing="4" bgcolor="#FFFFFF" style="BORDER-LEFT: #808e65 1px solid;BORDER-right: #808e65 1px solid;BORDER-top: #808e65 1px solid; BORDER-BOTTOM: #808e65 1px solid">
              <tbody><tr>
                <td height="105" align="center" valign="top" class="p9"><a href="zj.asp" title=""><img src="images/uploadimages/200837185230470.jpg" alt="" height="136" border="0"><br>
                  获奖无数</a></td>
                </tr>
                </tbody></table>
              <img src="images/bg_photoList.gif" width="90" height="5"></td>
            </tr>
        </tbody>
      </table></td>
    </tr>
  </tbody>
</table>
<br>
<table width="1020" border="0" align="center" cellpadding="4" cellspacing="4" bgcolor="#FFD9FF">
  <tbody>
    <tr>
      <td height="40"><span class="nowprice">友情链接：</span><a href="http://www.usa.gov/" title="美國政府" target="_blank" class="large">美國政府</a>
<a target="_blank" title="克利夫蘭醫學中心" href="http://www.clevelandclinic.org/ " class="large">克利夫蘭醫學中心</a>
<a target="_blank" title="波士頓兒童醫院" href="http://www.childrenshospital.org/ " class="large">波士頓兒童醫院</a>
<a target="_blank" title="美國國立衛生研究院" href="http://www.nih.gov/ " class="large">美國國立衛生研究院</a>
<a target="_blank" title="MyMedLab" href="https://www.mymedlab.com/ " class="large">MyMedLab</a>
<a target="_blank" title="布萊根婦女醫院" href="http://www.brighamandwomens.org/ " class="large">布萊根婦女醫院</a>
<a target="_blank" title="google" href="http://www.google.com/ " class="large">Google</a>
<a target="_blank" title="百度" href="http://www.baidu.com/ " class="large">百度</a>
<a target="_blank" title="美国保健品、食品、药品协会" href="http://www.anaha.org/ " class="large">美国保健品食品药品协会</a></td>
    </tr>
  </tbody>
</table>
<!-- #include file="foot.asp" -->

<%
conn.close
set conn=nothing
%>
<script language="javascript"> 
function showsrc()
{
	imgs = document.getElementsByTagName("img");
	imgsnum = imgs.length;
	for(imgi = 0 ;imgi < imgsnum;imgi++){
		 if ((typeof(imgs[imgi].src) == 'undefined' || imgs[imgi].src =='') && imgs[imgi].getAttribute('thissrc') != null)
		 imgs[imgi].src = imgs[imgi].getAttribute('thissrc');
	}
}
window.setTimeout("showsrc();",400);
</script>

</body>
</html>




