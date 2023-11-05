function showSubCategory(){
	$("#AllSort dt").mouseover(function(){
			var newDiv=document.getElementById("EFF_div_"+this.id.substr(7));
			this.className="curr";
			if (newDiv){
				newDiv.style.display="block";
				return;
			}else{
				var CLASS_NAME=($.browser.msie?($.browser.version>"7.0")?"class":"className":"class");
				var newDiv_wrap=document.createElement("div");
				newDiv_wrap.setAttribute(CLASS_NAME,"pop_wrap");
				newDiv_wrap.setAttribute("id","EFF_div_"+this.id.substr(7))
				var newDiv=document.createElement("div");
				newDiv.setAttribute(CLASS_NAME,"pop");
				newDiv_wrap.appendChild(newDiv);
				newDiv.innerHTML=this.nextSibling.innerHTML;
				this.parentNode.insertBefore(newDiv_wrap,this);	
				newDiv_wrap.style.display="block";
			}
		//test/		
		$(".pop_wrap").mouseover(function(){
			$(this).css({"display":"block"});
			this.nextSibling.className="curr";
		}).bind("mouseleave",function(){
			$(this).css({"display":"none"});
			this.nextSibling.className="";
		})
		//test end/
	}).bind("mouseleave",function(){
		this.className=(this.nextSibling.className=="Dis")?"curr":"";
		$(".pop_wrap").css({"display":"none"});
	});	
}

