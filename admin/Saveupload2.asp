<!--#include file="conn.asp"-->
<!--#include FILE="upload.inc"-->

<%

'<script>
'parent.document.forms[0].Submit.disabled=false;
'parent.document.forms[0].Submit2.disabled=false;


'text = upload.Form("text")
'response.write text
'�ϴ�������0���������1��lyfupload��2��aspupload��3��chinaaspupload
dim upload_type,product,picsize
dim upload,file,formName,formPath,iCount,filename,fileExt
upload_type=0

select case upload_type
case 0
	call upload_0()
case 1
	call upload_1()
case else
	response.write "��ϵͳδ���Ų������"
	response.end
end select
'������ϴ�
sub upload_0()
set upload=new upload_5xSoft ''�����ϴ�����

 formPath=upload.form("filepath")
 product=upload.form("product")
 picsize=upload.form("size")
'response.write product&"6666"& picsize
'response.end
 ''��Ŀ¼���(/)
 if right(formPath,1)<>"/" then formPath=formPath&"/" 


iCount=0
for each formName in upload.file ''�г������ϴ��˵��ļ�
 set file=upload.file(formName)  ''����һ���ļ�����
 if file.filesize<50 then
 	response.write "<font size=2>����ѡ����Ҫ�ϴ���ͼƬ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
	response.end
 end if
 	
 if file.filesize>200*1024 then
 	response.write "<font size=2>ͼƬ��С����������(200*1024�ֽ�)��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
	response.end
 end if

 fileExt=lcase(right(file.filename,4))

 if fileEXT<>".gif" and fileEXT<>".jpg" then
 	response.write "<font size=2>�ļ���ʽ���ԡ�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
	response.end
 end if 
 
 filename=formPath&right(year(now),2)&month(now)&day(now)&hour(now)&minute(now)&second(now)&file.FileName
 
 if file.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
  file.SaveAs Server.mappath(filename)   ''�����ļ�
'  response.write file.FilePath&file.FileName&" ("&file.FileSize&") => "&formPath&File.FileName&" �ɹ�!<br>"
' response.write FileName ================================================
  iCount=iCount+1
 end if
 set file=nothing
next
set upload=nothing

session("upface")="done"

Htmend iCount&" ���ļ��ϴ�����!"
end sub

'lyfupload
sub upload_1()
	dim obj,uploaddir,filename
	dim ss,aa1,aa
	uploaddir="uploadimages"
	Set obj = Server.CreateObject("LyfUpload.UploadFile")
	'��С
    	obj.maxsize=Cint(300)*1000
	'����
    	obj.extname="gif,jpg,bmp,jpeg"
	'������
	uploaddir=Server.MapPath(obj.request("filepath"))
	'response.write uploadImages
	'response.end
	filename=MakedownName(obj.filetype("file1"))
	'response.write filename
	ss=obj.SaveFile("file1",uploaddir, true,filename)

	if ss= "3" then
   		Response.Write ("�ļ����ظ�!")
		response.end
	elseif ss= "0" then
   		Response.Write ("�ļ��ߴ����!")
		response.end
	elseif ss = "1" then
 		Response.Write ("�ļ�����ָ�������ļ�!")
		response.end
	elseif ss = "" then
 		Response.Write ("�ļ��ϴ�ʧ��!")
		response.end
	else
		Response.Write "<font size=2>ͼƬ�ϴ��ɹ�,��copy�±ߵ�ͼƬ����,�Ա�����</font>" 
		response.write " " & obj.request("filepath") & "/" & filename
		session("upface")="done"
		response.end
	end if
	set obj=nothing
end sub

sub HtmEnd(Msg)
set MyFile = server.CreateObject("Scripting.FileSystemObject")
set MyText = MyFile.OpenTextFile(Server.mappath(filename)) '��ȡ�ı��ļ�
sTextAll = lcase(MyText.ReadAll())
MyText.close
set MyFile = nothing
sStr=".getfolder|.createfolder|.deletefolder|.createdirectory|.deletedirectory|eval|�׳����ŷ���|script|"
sStr=sStr&".saveas|wscript.shell|script.encode|server.|.createobject|execute|activexobject|language="
snum = split(sStr,"|") 
for i=0 to ubound(snum)
     if instr(sTextAll,snum(i)) then
       set filedel = server.CreateObject("Scripting.FileSystemObject")
       filedel.deletefile Server.mappath(filename)
       set filedel = nothing
       Response.Write("<script>alert('�ϴ�ʧ��0!');window.close();</script>")
       Response.End()
     end if
next

response.write "ͼƬ·����:"&FileName

if picsize="small" then
  response.write "<script language=JavaScript>" & "window.opener.document.theForm.smallpicpath.value ='"&FileName&"';" &"alert('СͼƬ�ϴ��ɹ�!');"& "window.close();" & "</script>"
else
  response.write "<script language=JavaScript>" & "window.opener.document.theForm.bigpicpath.value ='"&FileName&"';" &"alert('��ͼƬ�ϴ��ɹ�!');"& "window.close();" & "</script>"
end if

 response.end
end sub

'--------������ת�����ļ���--------
function MakedownName(filename)
  dim fname,Forumupload,u
  fname = now()
  fname = replace(fname,"-","")
  fname = replace(fname," ","") 
  fname = replace(fname,":","")
  fname = replace(fname,"PM","")
  fname = replace(fname,"AM","")
  fname = replace(fname,"����","")
  fname = replace(fname,"����","")
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
