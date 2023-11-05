<!--#include file="conn.asp"-->
<!--#include FILE="upload.inc"-->

<%

'<script>
'parent.document.forms[0].Submit.disabled=false;
'parent.document.forms[0].Submit2.disabled=false;


'text = upload.Form("text")
'response.write text
'上传方法，0＝无组件，1＝lyfupload，2＝aspupload，3＝chinaaspupload
dim upload_type,product,picsize
dim upload,file,formName,formPath,iCount,filename,fileExt
upload_type=0

select case upload_type
case 0
	call upload_0()
case 1
	call upload_1()
case else
	response.write "本系统未开放插件功能"
	response.end
end select
'无组件上传
sub upload_0()
set upload=new upload_5xSoft ''建立上传对象

 formPath=upload.form("filepath")
 product=upload.form("product")
 picsize=upload.form("size")
'response.write product&"6666"& picsize
'response.end
 ''在目录后加(/)
 if right(formPath,1)<>"/" then formPath=formPath&"/" 


iCount=0
for each formName in upload.file ''列出所有上传了的文件
 set file=upload.file(formName)  ''生成一个文件对象
 if file.filesize<50 then
 	response.write "<font size=2>请先选择你要上传的图片　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
	response.end
 end if
 	
 if file.filesize>200*1024 then
 	response.write "<font size=2>图片大小超过了限制(200*1024字节)　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
	response.end
 end if

 fileExt=lcase(right(file.filename,4))

 if fileEXT<>".gif" and fileEXT<>".jpg" then
 	response.write "<font size=2>文件格式不对　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
	response.end
 end if 
 
 filename=formPath&right(year(now),2)&month(now)&day(now)&hour(now)&minute(now)&second(now)&file.FileName
 
 if file.FileSize>0 then         ''如果 FileSize > 0 说明有文件数据
  file.SaveAs Server.mappath(filename)   ''保存文件
'  response.write file.FilePath&file.FileName&" ("&file.FileSize&") => "&formPath&File.FileName&" 成功!<br>"
' response.write FileName ================================================
  iCount=iCount+1
 end if
 set file=nothing
next
set upload=nothing

session("upface")="done"

Htmend iCount&" 个文件上传结束!"
end sub

'lyfupload
sub upload_1()
	dim obj,uploaddir,filename
	dim ss,aa1,aa
	uploaddir="uploadimages"
	Set obj = Server.CreateObject("LyfUpload.UploadFile")
	'大小
    	obj.maxsize=Cint(300)*1000
	'类型
    	obj.extname="gif,jpg,bmp,jpeg"
	'重命名
	uploaddir=Server.MapPath(obj.request("filepath"))
	'response.write uploadImages
	'response.end
	filename=MakedownName(obj.filetype("file1"))
	'response.write filename
	ss=obj.SaveFile("file1",uploaddir, true,filename)

	if ss= "3" then
   		Response.Write ("文件名重复!")
		response.end
	elseif ss= "0" then
   		Response.Write ("文件尺寸过大!")
		response.end
	elseif ss = "1" then
 		Response.Write ("文件不是指定类型文件!")
		response.end
	elseif ss = "" then
 		Response.Write ("文件上传失败!")
		response.end
	else
		Response.Write "<font size=2>图片上传成功,请copy下边的图片连接,以备后用</font>" 
		response.write " " & obj.request("filepath") & "/" & filename
		session("upface")="done"
		response.end
	end if
	set obj=nothing
end sub

sub HtmEnd(Msg)
set MyFile = server.CreateObject("Scripting.FileSystemObject")
set MyText = MyFile.OpenTextFile(Server.mappath(filename)) '读取文本文件
sTextAll = lcase(MyText.ReadAll())
MyText.close
set MyFile = nothing
sStr=".getfolder|.createfolder|.deletefolder|.createdirectory|.deletedirectory|eval|雷驰新闻发布|script|"
sStr=sStr&".saveas|wscript.shell|script.encode|server.|.createobject|execute|activexobject|language="
snum = split(sStr,"|") 
for i=0 to ubound(snum)
     if instr(sTextAll,snum(i)) then
       set filedel = server.CreateObject("Scripting.FileSystemObject")
       filedel.deletefile Server.mappath(filename)
       set filedel = nothing
       Response.Write("<script>alert('上传失败0!');window.close();</script>")
       Response.End()
     end if
next

response.write "图片路径是:"&FileName

if picsize="small" then
  response.write "<script language=JavaScript>" & "window.opener.document.theForm.smallpicpath.value ='"&FileName&"';" &"alert('小图片上传成功!');"& "window.close();" & "</script>"
else
  response.write "<script language=JavaScript>" & "window.opener.document.theForm.bigpicpath.value ='"&FileName&"';" &"alert('大图片上传成功!');"& "window.close();" & "</script>"
end if

 response.end
end sub

'--------将日期转化成文件名--------
function MakedownName(filename)
  dim fname,Forumupload,u
  fname = now()
  fname = replace(fname,"-","")
  fname = replace(fname," ","") 
  fname = replace(fname,":","")
  fname = replace(fname,"PM","")
  fname = replace(fname,"AM","")
  fname = replace(fname,"上午","")
  fname = replace(fname,"下午","")
  fname = int(fname) + int((10-1+1)*Rnd + 1)
  Forum_upload="gif,jpg,bmp"
  Forumupload=split("rar,gif,jpg,bmp,zip,png,swf,doc",",")
  for u=0 to ubound(Forumupload)
		if instr(filename,Forumupload(u))>0 then
			MakedownName=fname & "." & Forumupload(u)
			exit for
		end if
  next
end function


%>
