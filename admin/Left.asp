<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>管理页面</title>

<script src="images/prototype.lite.js" type="text/javascript"></script>
<script src="images/moo.fx.js" type="text/javascript"></script>
<script src="images/moo.fx.pack.js" type="text/javascript"></script>
<style>
body {
	font:12px Arial, Helvetica, sans-serif;
	color: #000;
	background-color: #EEF2FB;
	margin: 0px;
}
#container {
	width: 182px;
}
H1 {
	font-size: 12px;
	margin: 0px;
	width: 182px;
	cursor: pointer;
	height: 30px;
	line-height: 20px;	
}
H1 a {
	display: block;
	width: 182px;
	color: #000;
	height: 30px;
	text-decoration: none;
	moz-outline-style: none;
	background-image: url(images/menu_bgS.gif);
	background-repeat: no-repeat;
	line-height: 30px;
	text-align: center;
	margin: 0px;
	padding: 0px;
}
.content{
	width: 182px;
	height: 26px;
	
}
.MM ul {
	list-style-type: none;
	margin: 0px;
	padding: 0px;
	display: block;
}
.MM li {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	line-height: 26px;
	color: #333333;
	list-style-type: none;
	display: block;
	text-decoration: none;
	height: 26px;
	width: 182px;
	padding-left: 0px;
}
.MM {
	width: 182px;
	margin: 0px;
	padding: 0px;
	left: 0px;
	top: 0px;
	right: 0px;
	bottom: 0px;
	clip: rect(0px,0px,0px,0px);
}
.MM a:link {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	line-height: 26px;
	color: #333333;
	background-image: url(images/menu_bg1.gif);
	background-repeat: no-repeat;
	height: 26px;
	width: 182px;
	display: block;
	text-align: center;
	margin: 0px;
	padding: 0px;
	overflow: hidden;
	text-decoration: none;
}
.MM a:visited {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	line-height: 26px;
	color: #333333;
	background-image: url(images/menu_bg1.gif);
	background-repeat: no-repeat;
	display: block;
	text-align: center;
	margin: 0px;
	padding: 0px;
	height: 26px;
	width: 182px;
	text-decoration: none;
}
.MM a:active {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	line-height: 26px;
	color: #333333;
	background-image: url(images/menu_bg1.gif);
	background-repeat: no-repeat;
	height: 26px;
	width: 182px;
	display: block;
	text-align: center;
	margin: 0px;
	padding: 0px;
	overflow: hidden;
	text-decoration: none;
}
.MM a:hover {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	line-height: 26px;
	font-weight: bold;
	color: #006600;
	background-image: url(images/menu_bg2.gif);
	background-repeat: no-repeat;
	text-align: center;
	display: block;
	margin: 0px;
	padding: 0px;
	height: 26px;
	width: 182px;
	text-decoration: none;
}
</style>
</head>

<body>
<table width="100%" height="280" border="0" cellpadding="0" cellspacing="0" bgcolor="#EEF2FB">
  <tr>
    <td width="182" valign="top"><div id="container">
      <h1 class="type"><a href="javascript:void(0)">网站常规管理</a></h1>
      <div class="content">
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><img src="images/menu_topline.gif" width="182" height="5" /></td>
          </tr>
        </table>
        <ul class="MM">
          <li><a href="siteinfo.asp" target="main">基本设置</a></li>
          <li><a href="sitemeta.asp" target="main">页头信息</a></li>
          <!--<li><a href="siteright.asp" target="main">页尾信息</a></li>-->
          <li><a href="sitead.asp" target="main">广告图片</a></li>
          <li><a href="siteshow.asp" target="main">其他设置</a></li>
        
        </ul>
      </div>
      <h1 class="type"><a href="javascript:void(0)">商品信息管理</a></h1>
      <div class="content">
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><img src="images/menu_topline.gif" width="182" height="5" /></td>
          </tr>
        </table>
        <ul class="MM">
          <li><a href="../product/b_kind.asp" target="main">分类管理</a></li>
          <li><a href="../product/prod.asp" target="main">添加商品</a></li>
          <li><a href="../product/productlist.asp" target="main">修改商品</a></li>
          <li><a target="main" href="pl_modify.asp">批量修改</a></li>
           
          <li><a href="pijiaadmin.asp" target="main">商品评论</a></li>
          
        </ul>
      </div>
      <h1 class="type"><a href="javascript:void(0)">订单信息管理</a></h1>
      <div class="content">
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><img src="images/menu_topline.gif" width="182" height="5" /></td>
          </tr>
        </table>
        <ul class="MM">
		  <li><a href="dingdan.asp" target="main">订单管理</a></li>
          
          <li><a href="wuliugs.asp" target="main">物流管理</a></li>
          <li><a href="report.asp" target="main">商品销售报表</a></li>
   
        </ul>
      </div>
      <h1 class="type"><a href="javascript:void(0)">新闻内容管理</a></h1>
      <div class="content">
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><img src="images/menu_topline.gif" width="182" height="5" /></td>
          </tr>
        </table>
        <ul class="MM">
          <li><a href="conadd.asp" target="main">添加网站新闻</a></li>
          <li><a href="conlist.asp?l_id=30" target="main">管理网站新闻</a></li>
		  <li><a href="conlist.asp?l_id=48" target="main">管理购物指南</a></li>
		  <li><a href="conlist.asp?l_id=45" target="main">管理送货说明</a></li>
		  <li><a href="conlist.asp?l_id=46" target="main">管理支付方式</a></li>
		  <li><a href="conlist.asp?l_id=44" target="main">管理服务政策</a></li>
		  <li><a href="conlist.asp?l_id=43" target="main">管理关于我们</a></li>
		  <li><a href="khly.asp" target="main">网站留言管理</a></li>
          </ul>
      </div>
    </div>
        <h1 class="type"><a href="javascript:void(0)">支付方式设置</a></h1>
        <div class="content">
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td><img src="images/menu_topline.gif" width="182" height="5" /></td>
            </tr>
          </table>
        <ul class="MM">
            <li><a href="pay1.asp" target="main">支付参数</a></li>
            <li><a href="pay2.asp" target="main">银行转帐</a></li>
            <li><a href="pay3.asp" target="main">货到付款</a></li>
            <li><a href="pay4.asp" target="main">邮局汇款</a></li>
            <li><a href="pay5.asp" target="main">帐户余额</a></li>
        </ul>
      </div>
	          <h1 class="type"><a href="javascript:void(0)">会员信息管理</a></h1>
              <div class="content">
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td><img src="images/menu_topline.gif" width="182" height="5" /></td>
            </tr>
          </table>
        <ul class="MM">
            <li><a href="huiyuan.asp" target="main">会员管理</a></li>
            <li><a href="user.asp" target="main">管理员管理</a></li>
			<li><a href="passwordmodify.asp" target="main">修改密码</a></li>
            <li><a href="chongzhi.asp" target="main">会员充值</a></li>
            <li><a href="pay5.asp" target="main">帐户余额</a></li>
        </ul>
      </div>
	  
	   <h1 class="type"><a href="javascript:void(0)">配送计费管理</a></h1>
              <div class="content">
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td><img src="images/menu_topline.gif" width="182" height="5" /></td>
            </tr>
          </table>
        <ul class="MM">
            <li><a href="wuliuflag.asp" target="main">配送方式</a></li>
            <li><a href="sendkind.asp" target="main">计费设置</a></li>
          
        </ul>
      </div>
	     <h1 class="type"><a href="javascript:void(0)">积分礼品设置</a></h1>
              <div class="content">
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td><img src="images/menu_topline.gif" width="182" height="5" /></td>
            </tr>
          </table>
        <ul class="MM">
            <li><a href="zhfbili.asp" target="main">积分比率</a></li>
            <li><a href="lipin.asp" target="main">礼品设置</a></li>
           <li><a href="huihuan.asp" target="main">礼品兑换</a></li>
		   
        </ul>
      </div>
	  
	  
	  
	   <h1 class="type"><a href="javascript:void(0)">数据备份管理</a></h1>
              <div class="content">
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td><img src="images/menu_topline.gif" width="182" height="5" /></td>
            </tr>
          </table>
        <ul class="MM">
            <li><a href="mgdata.asp?action=backup" target="main">数据库备份</a></li>
            <li><a href="mgdata.asp?action=backupList" target="main">数据库恢复</a></li>
           <li><a href="mgdata.asp?action=compact" target="main">数据库压缩</a></li>
		  <li><a href="../CheckSpace.asp" target="main">空间占用量</a></li>
		   <li><a href="check.asp" target="main">系统环境</a></li>
        </ul>
      </div>
	  
      </div>
        <script type="text/javascript">
		var contents = document.getElementsByClassName('content');
		var toggles = document.getElementsByClassName('type');
	
		var myAccordion = new fx.Accordion(
			toggles, contents, {opacity: true, duration: 400}
		);
		myAccordion.showThisHideOpen(contents[0]);
	</script>
        </td>
  </tr>
</table>
</body>
</html>
