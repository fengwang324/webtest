<SCRIPT language=javascript>
function myleftload(){
myleft.style.top=document.body.scrollTop+document.body.offsetHeight-280;
myleft.style.left=1;
myleftmove();}
function myleftmove(){
 myleft.style.top=document.body.scrollTop+document.body.offsetHeight-280;
 myleft.style.left=1;     
 setTimeout("myleftmove();",100)}

function myrightload(){
myright.style.top=document.body.scrollTop+document.body.offsetHeight-280;
myright.style.right=10;
myrightmove();}
function myrightmove(){
myright.style.top=document.body.scrollTop+document.body.offsetHeight-280;
myright.style.right=10; 
setTimeout("myrightmove();",100)}

function close_right(){
	myright.style.visibility='hidden';
}
function close_left(){
	myleft.style.visibility='hidden';
}

//要修改的话,请修改以下两个代码中<a href='点击后链接地址' target='_blank'><img src='images/ad1.jpg' width=108 height=图片高度 border=0 name=chan></a>
//点击后链接地址 和 图片高度
//注册要改对应的位置
document.write("<div id=myleft style='position: absolute;width:110;top:20;visibility: visible;z-index: 1'><TABLE><TR><TD align=left>  <a href='javascript:close_left();void(0);' target=_self><img src='images/qqclose.gif' border='0' width='27' height='14'></a></TD></TR><TR><TD><a href='<%=leftlink%>' target='_blank'><img src='<%=lefturl%>' width=110 height=140 border=0 name=chan></a></TD></TR></TABLE></div>");
document.write('<div id=myright style="position: absolute;width:100;top:20;visibility: visible;z-index: 1"><TABLE><TR><TD align=right><a href="javascript:close_right();void(0);" target=_self><img src="images/qqclose.gif" border="0" width="27" height="14"></a></TD></TR><TR><TD><A HREF="<%=rightlink%>" target="_blank"><IMG SRC="<%=righturl%>" WIDTH=110 HEIGHT=140 BORDER=0 ></a></TD></TR></TABLE></div>');

myrightload();
myleftload();

</SCRIPT>
