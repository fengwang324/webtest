var curPage = top.document.location.href.toLowerCase();

var $ = function(el){ return (typeof el=='object')?el:document.getElementById(el); };

function setMainPage(obj,URL){obj.style.behavior='url(#default#homepage)';obj.setHomePage(URL);};

function addMYFavorite(URL,Title){ window.external.addFavorite(URL,Title); };

var isChar     = /^[a-zA-Z0-9_]{6,20}/;
var isEmail    = /^([a-zA-Z0-9_-])+@([a-zA-Z0-9_-])+(.[a-zA-Z0-9_-])+/; 
var isNumber   = /^[0-9]/;  
var isPostCode = /^[0-9]{6}/;

function Flash(ur,w,h){
    document.write('<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" \
    codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0" \
	width="'+w+'" height="'+h+'" VIEWASTEXT> \
	<param name="movie" value="' + ur + '"> \
	<param name="quality" value="high"> \
	<param name="wmode" value="transparent"> \
	<param name="menu" value="false"> \
	<embed src="' + ur + '" quality="high" \
	 pluginspage="http://www.macromedia.com/go/getflashplayer" \
	 type="application/x-shockwave-flash" wmode="transparent" \
	width="'+w+'" height="'+h+'" ></embed></object> ');
}
