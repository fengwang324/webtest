<link href="i.css" type=text/css rel=stylesheet>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<style type="text/css">
<!--
.STYLE1 {
	color: #000000;
	font-weight: bold;
}
-->
</style>
<BODY>
<table class="tableBorder" width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
  <tr>
    <td colspan="4" align="center" bgcolor="#E1EFFF"><span class="STYLE1">空间的使用情况</span>
	<% 	
 	Sub ShowSpaceInfo(drvpath)
 		dim fso,d,size,showsize
 		set fso=server.createobject("scripting.filesystemobject") 		
 		drvpath=server.mappath(drvpath) 		 		
 		set d=fso.getfolder(drvpath) 		
 		size=d.size
 		showsize=size & "&nbsp;Byte" 
 		if size>1024 then
 		   size=(size\1024)
 		   showsize=size & "&nbsp;KB"
 		end if
 		if size>1024 then
 		   size=(size/1024)
 		   showsize=size & "&nbsp;MB"		
 		end if
 		if size>1024 then
 		   size=(size/1024)
 		   showsize=size & "&nbsp;GB"	   
 		end if   
 		response.write "<font face=verdana>" & showsize & "</font>"
 	End Sub	
 	Sub Showspecialspaceinfo(method)
 		dim fso,d,fc,f1,size,showsize,drvpath 		
 		set fso=server.createobject("scripting.filesystemobject")
 		drvpath=server.mappath("pic")
 		drvpath=left(drvpath,(instrrev(drvpath,"\")-1))
 		set d=fso.getfolder(drvpath) 		
 		if method="All" then 		
 			size=d.size
 		elseif method="Program" then
 			set fc=d.Files
 			for each f1 in fc
 				size=size+f1.size
 			next	
 		end if	
 		showsize=size & "&nbsp;Byte" 
 		if size>1024 then
 		   size=(size\1024)
 		   showsize=size & "&nbsp;KB"
 		end if
 		if size>1024 then
 		   size=(size/1024)
 		   showsize=size & "&nbsp;MB"		
 		end if
 		if size>1024 then
 		   size=(size/1024)
 		   showsize=size & "&nbsp;GB"	   
 		end if   
 		response.write "<font face=verdana>" & showsize & "</font>"
 	end sub 	 	 	
 	Function Drawbar(drvpath)
 		dim fso,drvpathroot,d,size,totalsize,barsize
 		set fso=server.createobject("scripting.filesystemobject")
 		drvpathroot=server.mappath("pic")
 		drvpathroot=left(drvpathroot,(instrrev(drvpathroot,"\")-1))
 		set d=fso.getfolder(drvpathroot)
 		totalsize=d.size
 		drvpath=server.mappath(drvpath) 		
 		set d=fso.getfolder(drvpath)
 		size=d.size
 		barsize=cint((size/totalsize)*400)
 		Drawbar=barsize
 	End Function 	
 	Function Drawspecialbar()
 		dim fso,drvpathroot,d,fc,f1,size,totalsize,barsize
 		set fso=server.createobject("scripting.filesystemobject")
 		drvpathroot=server.mappath("pic")
 		drvpathroot=left(drvpathroot,(instrrev(drvpathroot,"\")-1))
 		set d=fso.getfolder(drvpathroot)
 		totalsize=d.size
 		set fc=d.files
 		for each f1 in fc
 			size=size+f1.size
 		next	
 		barsize=cint((size/totalsize)*400)
 		Drawspecialbar=barsize
 	End Function 	
	%>	</td>
  </tr>
  <tr>
    <td valign=top bgcolor=#fbf4f4><table width="100%" align="center" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td><%
 			fsorank=1
 			if fsorank=1 then
 			%>
            <%Function GetPP
			dim s
			s=Request.ServerVariables("path_translated")
			GetPP=left(s,instrrev(s,"\",len(s)))
 			End function
 			if sPP="" then sPP=GetPP
 			if right(sPP,1)<>"\" then sPP=sPP&"\"
 			set fso=server.createobject("scripting.filesystemobject")
 			Set f = fso.GetFolder(sPP)
 			Set fc = f.SubFolders
 			i=1
			i2=1
 			For Each f in fc%>
            <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#ffffff">
              <tr bgcolor="#ffffff">
                <td width="20%" bgcolor="#fbf4f4">&nbsp; <%=f.name%>文件夹</td>
                <td width="80%" bgcolor="#fbf4f4">&nbsp;<img src="images/bar.gif" width=<%=drawbar(""&f.name&"")%> height=10>&nbsp;
                    <%showSpaceinfo(""&f.name&"")%></td>
              </tr>
            </table>
          <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td height="2"></td>
              </tr>
            </table>
          <%i=i+1
			if i2<10 then
			i2=i2+1
			else
			i2=1
			end if
			Next%>
            <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#ffffff">
              <tr bgcolor="#ffffff">
                <td width="20%" bgcolor="#fbf4f4">&nbsp;程序文件占用空间</td>
                <td width="80%" bgcolor="#fbf4f4">&nbsp;<img src="images/bar.gif" width=<%=drawspecialbar%> height=10>&nbsp;
                    <%showSpecialSpaceinfo("Program")%></td>
              </tr>
            </table>
          <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td height="2"></td>
              </tr>
            </table>
          <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#ffffff">
              <tr bgcolor="#ffffff">
                <td width="20%" bgcolor="#fbf4f4">&nbsp;系统占用空间总计</td>
                <td width="80%" bgcolor="#fbf4f4">&nbsp;<img src="images/bar.gif" width=400 height=10>
                    <%showspecialspaceinfo("All")%>
                </td>
              </tr>
            </table>
          <%
 			else
			response.write "<br><li>本功能已经被关闭"
 			end if
 			%>
        </td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
