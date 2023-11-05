<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN"
	"http://www.w3.org/TR/html4/strict.dtd">
<html lang="en" xml:lang="en">
<% DBPaths="../" %>
<!--#include file="../CONN.asp"-->
<!--#include file="../../common/Function.asp"-->
<!--#include file="../System_Function.asp"-->
<% 
dim typeFilter,path
path=trim(s("path"))
select case path
case "video"
typeFilter="Video (*.mid,*.mp3,*.wmv,*.asf,*.avi,*.mpg,*.ram,*.rm,*.ra)':'*.mid;*.mp3;*.wmv;*.asf;*.avi;*.mpg;*.ram;*.rm;*.ra"
case "flash"
typeFilter="Flash (*.swf, *.flv)':'*.swf;*.flv;"
case "other"
typeFilter="file (*.rar,*.doc,*.zip)':'*.rar;*.doc;*.zip"
case else 
typeFilter=typeFiler(Setting(7))
end select

function typeFiler(str)
if str="" then
 typeFiler="Images (*.jpg, *.jpeg, *.gif, *.png)': '*.jpg; *.jpeg; *.gif; *.png"
 else
 strarr=split(str,"|")
	 for i=0 to ubound(strarr)
	 	if i=ubound(strarr) then
		strarr1=strarr1&"*."&strarr(i)
		strarr2=strarr2&"*."&strarr(i)
		else
	 	strarr1=strarr1&"*."&strarr(i)&","
		strarr2=strarr2&"*."&strarr(i)&";"
		end if
	 next
 typeFiler="图片类型 ("&strarr1&")':'"&strarr2&""""
 end if
end function
 %>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<title>(可批量)选择文件 -> 开始上传</title>

	<meta name="author" content="Harald Kirschner, digitarald.de" />
	<meta name="copyright" content="Copyright 2009 Harald Kirschner" />

	<!-- Framework CSS -->
	<link rel="stylesheet" href="source/screen.css" type="text/css" media="screen, projection">
	<link rel="stylesheet" href="source/print.css" type="text/css" media="print">
	<script type="text/javascript" src="source/mootools.js"></script>

	<script type="text/javascript" src="source/Swiff.Uploader.js"></script>

	<script type="text/javascript" src="source/Fx.ProgressBar.js"></script>

	<script type="text/javascript" src="source/FancyUpload2.js"></script>


	<!-- See script.js -->
	<script type="text/javascript">
		//<![CDATA[

		/**
 * FancyUpload Showcase
 *
 * @license		MIT License
 * @author		Harald Kirschner <mail [at] digitarald [dot] de>
 * @copyright	Authors
 */

window.addEvent('domready', function() { // 等待内容

	// our uploader instance 
	
	var up = new FancyUpload2($('demo-status'), $('demo-list'), { // 选择对象
		// we console.log infos, remove that in production!!
		verbose: true,
		
		// 网址是读取的形式
		url: $('form-demo').action,
		
		// SWF文件路径
		path: 'source/Swiff.Uploader.swf',
		
		// 删除该行选择所有文件，或修改，添加更多类型
		typeFilter: {
			'<%= typeFilter %>'
		},
		
		// 这是浏览按钮, *target* 是与覆盖的Flash电影
		target: 'demo-browse',
		
		// graceful degradation, onLoad is only called if all went well with Flash
		onLoad: function() {
			$('demo-status').removeClass('hide'); // we show the actual UI
			$('demo-fallback').destroy(); // ... and hide the plain form
			
			// We relay the interactions with the overlayed flash to the link
			this.target.addEvents({
				click: function() {
					return false;
				},
				mouseenter: function() {
					this.addClass('hover');
				},
				mouseleave: function() {
					this.removeClass('hover');
					this.blur();
				},
				mousedown: function() {
					this.focus();
				}
			});

			// Interactions for the 2 other buttons
			
			$('demo-clear').addEvent('click', function() {
				up.remove(); // 移除所有文件
				return false;
			});

			$('demo-upload').addEvent('click', function() {
				up.start(); // 开始上传
				return false;
			});
		},
		
		// Edit the following lines, it is your custom event handling
		
		/**
		 * Is called when files were not added, "files" is an array of invalid File classes.
		 * 
		 * This example creates a list of error elements directly in the file list, which
		 * hide on click.
		 */ 
		onSelectFail: function(files) {
			files.each(function(file) {
				new Element('li', {
					'class': 'validation-error',
					html: file.validationErrorMessage || file.validationError,
					title: MooTools.lang.get('FancyUpload', 'removeTitle'),
					events: {
						click: function() {
							this.destroy();
						}
					}
				}).inject(this.list, 'top');
			}, this);
		},
				
		/**
		 * This one was directly in FancyUpload2 before, the event makes it
		 * easier for you, to add your own response handling (you probably want
		 * to send something else than JSON or different items).
		 */
		onFileSuccess: function(file, response) {
			var json = new Hash(JSON.decode(response, true) || {});
			if (response == '1') {
				file.element.addClass('file-success');
			} else {
				file.element.addClass('file-failed');
				file.info.set('html', '<strong>An error occured:</strong> ' + (json.get('error') ? (json.get('error') + ' #' + json.get('code')) : response));
			}
		},
		onAllComplete:function() {
        alert("上传完成!")
    },

		
		/**
		 * onFail is called when the Flash movie got bashed by some browser plugin
		 * like Adblock or Flashblock.
		 */
		onFail: function(error) {
			switch (error) {
				case 'hidden': // works after enabling the movie and clicking refresh
					alert('为了使嵌入式上传，在浏览器解除封锁，并刷新（见Adblock的）');
					break;
				case 'blocked': // This no *full* fail, it works after the user clicks the button
					alert('To enable the embedded uploader, enable the blocked Flash movie (see Flashblock).');
					break;
				case 'empty': // Oh oh, wrong path
					alert('A required file was not found, please be patient and we fix this.');
					break;
				case 'flash': // no flash 9+ :(
					alert('To enable the embedded uploader, install the latest Adobe Flash plugin.')
			}
		}
		
	});
	
});
		//]]>
	</script>



	<!-- See style.css -->
	<style type="text/css">
		/**
 * FancyUpload Showcase
 *
 * @license		MIT License
 * @author		Harald Kirschner <mail [at] digitarald [dot] de>
 * @copyright	Authors
 */

/* CSS vs. Adblock tabs */
.swiff-uploader-box a {
	display: none !important;
}

/* .hover simulates the flash interactions */
a:hover, a.hover {
	color: red;
}

#demo-status {
	padding: 10px 15px;
	width: 420px;
	border: 1px solid #eee;
	float:left
}

#demo-status .progress {
	background: url(assets/progress-bar/progress.gif) no-repeat;
	background-position: +50% 0;
	margin-right: 0.5em;
	vertical-align: middle;
}

#demo-status .progress-text {
	font-size: 0.9em;
	font-weight: bold;
}

#demo-list {
	list-style: none;
	width: 450px;
	margin: 0;
	height:300px;
	overflow:auto
}

#demo-list li.validation-error {
	padding-left: 44px;
	display: block;
	clear: left;
	line-height: 40px;
	color: #8a1f11;
	cursor: pointer;
	border-bottom: 1px solid #fbc2c4;
	background: #fbe3e4 url(assets/failed.png) no-repeat 4px 4px;
}

#demo-list li.file {
	border-bottom: 1px solid #eee;
	background: url(assets/file.png) no-repeat 4px 4px;
	overflow: auto;
}
#demo-list li.file.file-uploading {
	background-image: url(assets/uploading.png);
	background-color: #D9DDE9;
}
#demo-list li.file.file-success {
	background-image: url(assets/success.png);
}
#demo-list li.file.file-failed {
	background-image: url(assets/failed.png);
}

#demo-list li.file .file-name {
	font-size: 13px;
	margin-left: 44px;
	display: block;
	clear: left;
	line-height: 40px;
	height: 40px;
	font-weight: bold;
}
#demo-list li.file .file-size {
	font-size: 12px;
	line-height: 18px;
	float: right;
	margin-top: 2px;
	margin-right: 6px;
}
#demo-list li.file .file-info {
	display: block;
	margin-left: 44px;
	font-size: 13px;
	line-height: 20px;
	clear
}
#demo-list li.file .file-remove {
	font-size:12px;
	clear: right;
	float: right;
	line-height: 18px;
	margin-right: 6px;
}
#after{
content:"";
display:block;
height:0;
line-height:0;
clear:both;
visibility:hidden;
}
</style>


</head>
<body>

	<div class="container">
		<div>
			<form action="../uploadTester.asp?Path=<%= Trim(Request.QueryString("path")) %>" method="post"  id="form-demo">

	<fieldset id="demo-fallback">
		<legend>文件上传</legend>
		<label for="demo-photoupload">
			上传图片:
			<input type="file" name="Filedata" />
            <input name="Path" type="hidden" value="<%= Trim(Request.QueryString("path")) %>">
		</label>
	</fieldset>

	<div id="demo-status" class="hide"><!---->

		<div style="float:right; width:115px; position:absolute; margin-left: 310px; margin-top: 0px; height: 100px; padding-top:1px;"><table border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="18%" height="25" align="right"><img src="../images/1.gif" width="16" height="16"></td>
    <td width="82%" align="left"><a href="#" id="demo-browse">选择文件</a></td>
  </tr>
   <tr>
    <td height="22" align="right"><img src="../images/2.gif" width="16" height="16"></td>
    <td align="left"><a href="#" id="demo-clear">清空列表</a></td>
  </tr>
   <tr>
    <td height="25" align="right"><img src="../images/3.gif" width="16" height="16"></td>
    <td align="left"><a href="#" onClick="document.all.child1.style.display =''" id="demo-upload">开始上传</a></td>
  </tr>
   <tr id="child1" style="display:none">
    <td height="22" align="right"><img src="../images/yes.gif" width="16" height="16"></td>
    <td align="left"><a href="#" onClick="window.close()">完成并关闭窗口</a></td>
  </tr>
</table>
        </div><div style="width:296px;">
			<strong class="overall-title"></strong><br />
			<img src="assets/progress-bar/bar.gif" class="progress overall-progress" />
		</div>
        
		<div style="width:296px;">
	    <strong class="current-title"></strong><br />
			<img src="assets/progress-bar/bar.gif" class="progress current-progress" />
		</div>
		<div class="current-text"></div>
	</div>
<div id="after" class="demo"></div>
	<ul id="demo-list"></ul>

</form>		</div>
</div>
</body>
</html>
