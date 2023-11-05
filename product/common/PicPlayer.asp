<STYLE>
#au {
	FILTER: progid:DXImageTransform.Microsoft.Fade ( duration=0.5,overlap=1.0 ); FLOAT: left; WIDTH: 216px; HEIGHT: 146px
}
#au IMG {
	WIDTH: 216px; BORDER-TOP-STYLE: none; BORDER-RIGHT-STYLE: none; BORDER-LEFT-STYLE: none; HEIGHT: 146px; BORDER-BOTTOM-STYLE: none
}
#conau {
	FLOAT: left; WIDTH: 216px; LINE-HEIGHT: 14px; PADDING-TOP: 5px; HEIGHT: 15px; TEXT-ALIGN: center
}
#conau A {
	COLOR: #505050; LINE-HEIGHT: 18px; TEXT-DECORATION: none; font-size:12px;
}
#conau A:hover {
	COLOR: #236a00; TEXT-DECORATION: underline
}
.bbg0 {
	FONT-SIZE: 10px; BACKGROUND: url(fp00.gif); CURSOR: hand; COLOR: white; LINE-HEIGHT: 10px; FONT-FAMILY: Arial
}
.bbg0 A {
	COLOR: white; TEXT-DECORATION: none
}
.bbg1 {
	FONT-SIZE: 10px; BACKGROUND: url(fp01.gif); CURSOR: hand; COLOR: white; LINE-HEIGHT: 10px; FONT-FAMILY: Arial
}
.bbg1 A {
	COLOR: white; TEXT-DECORATION: none
}
#No {
	MARGIN-TOP: -6px; Z-INDEX: 1; LEFT: -6px; WIDTH: 95px; POSITION: relative; TOP: 146px
}
</STYLE>

<TABLE class=marb6 cellSpacing=0 cellPadding=0 width=216 border=0>
<TBODY>
<TR><TD align=left width=216>
<SCRIPT> 
var n=0;
function Mea(value){ n=value;setBg(value);	plays(value);conaus(n);	}

function setBg(value)
{
	for(var i=0;i<5;i++){document.getElementById("t"+i+"").className="bbg0";}
	document.getElementById("t"+value+"").className="bbg1";
} 

function plays(value)
{
	try
	{
		with (au)
		{
			filters[0].Apply();
			for(i=0;i<5;i++)i==value?children[i].style.display="block":children[i].style.display="none"; 
			filters[0].play(); 		
		}
	}
	catch(e)
	{
		var d = document.getElementById("au").getElementsByTagName("div");
		for(i=0;i<5;i++)i==value?d[i].style.display="block":d[i].style.display="none"; 
	}
}

function conaus(value)
{
	try
	{
		with (conau)
		{
		    for(i=0;i<5;i++)i==value?children[i].style.display="block":children[i].style.display="none"; 
				
		}
	}
	catch(e)
	{
		var d = document.getElementById("conau").getElementsByTagName("div");
		for(i=0;i<5;i++)i==value?d[i].style.display="block":d[i].style.display="none"; 
	}
}

function clearAuto(){clearInterval(autoStart);}

function setAuto(){autoStart=setInterval("auto(n)", 2000)}

function auto(){n++;if(n>4)n=0;Mea(n);conaus(n);} 

setAuto(); 

</SCRIPT>
<TABLE cellSpacing=0 cellPadding=0 width=216 align=center border=0>
<TBODY>
<TR><TD vAlign=bottom align=right>
    <DIV id=No>
    <TABLE cellSpacing=0 cellPadding=0 width=93 border=0>
    <TBODY>
    <TR><TD class=bbg1 id=t0 onmouseover=Mea(0);clearAuto();onmouseout=setAuto() align=middle width=13 
            height=12><A href="#" target=_blank>1</A></TD>
        <TD width=7></TD>
        <TD class=bbg0 id=t1 onmouseover=Mea(1);clearAuto();onmouseout=setAuto() align=middle width=13>
		<A href="#" target=_blank>2</A></TD>
        <TD width=7></TD>
        <TD class=bbg0 id=t2 onmouseover=Mea(2);clearAuto();onmouseout=setAuto() align=middle width=13>
		<A href="#" target=_blank>3</A></TD>
        <TD width=7></TD>
        <TD class=bbg0 id=t3 onmouseover=Mea(3);clearAuto(); onmouseout=setAuto() align=middle width=13>
		<A href="#" target=_blank>4</A></TD>
        <TD width=7></TD>
        <TD class=bbg0 id=t4 onmouseover=Mea(4);clearAuto();onmouseout=setAuto() align=middle width=13>
		<A href="#" target=_blank>5</A></TD></TR></TBODY></TABLE></DIV>
    <DIV id=au>
		<DIV style="DISPLAY: block"><A href="#" target=_blank><IMG src="images/1.gif"></A></DIV>
		<DIV style="DISPLAY: none"><A href="#" target=_blank><IMG src="images/2.gif"></A></DIV>
		<DIV style="DISPLAY: none"><A href="#" target=_blank><IMG src="images/3.gif"></A></DIV>
		<DIV style="DISPLAY: none"><A href="#" target=_blank><IMG src="images/4.gif"></A></DIV>
		<DIV style="DISPLAY: none"><A href="#" target=_blank><IMG src="images/3.gif"></A></DIV>
	</DIV></TD></TR>
    <TR><TD align=middle background=images/a24.jpg height=21>
        <DIV id=conau>
		<DIV style="DISPLAY: block"></DIV>
		<DIV style="DISPLAY: none"></DIV>
		<DIV style="DISPLAY: none"></DIV>
		<DIV style="DISPLAY: none"></DIV>
		<DIV style="DISPLAY: none"></DIV>
		</DIV></TD>
	</TR></TBODY></TABLE></TD>
</TR></TBODY></TABLE>