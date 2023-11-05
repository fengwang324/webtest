<!--#include file="conn.asp"-->
<!--#include file="check2.asp"-->
<%
dbfolder="data/"   '数据库路径，如果你更改了数据库的路径，这里也请进行修改
dbname="#ioaskl089235.mdb"    
dbfolder="../"&dbfolder
%>
<meta http-equiv=Content-Type content=text/html; charset=utf-8>
<link href="../inc/i.css" type=text/css rel=stylesheet>
<style type="text/css">
<!--
.STYLE1 {
	color: #000000;
	font-weight: bold;
}
-->
</style>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
      <td> 
        <table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#6a7f9a">
          <tr> 
            
          <td width="15%" align="center" bgcolor="#E1EFFF"><span class="STYLE1">数据库管理操作</span></td>
          </tr>
         
          <tr > 
            
          <td height="30" bgcolor="#FFFFFF" > 

<%
SET FSO=Server.CreateObject("Scripting.FileSystemObject")
 IF Err<>0 THEN
 Err.Clear
 Response.Write("<p>您的服务器关闭了FSO，无法进行数据库操作")
 response.end
end if
action=request("action")
IF action<>"" THEN
strfileName=Request("FileName")

Function goaler_isFloderExist(strFolderName)
IF FSO.FolderExists(strFolderName) THEN
 goaler_isFloderExist=1
ELSE
 goaler_isFloderExist=-1
END IF
End Function

Function goaler_isfileExist(strfileName)
IF FSO.FileExists(strfileName) THEN
 goaler_isfileExist=1
ELSE
 goaler_isfileExist=-1
END IF
End Function

Function goaler_DeleteFile(strfileName)
 If FSO.FileExists(strfileName) Then
  FSO.DeleteFile strfileName,True
  goaler_DeleteFile=1
 Else
  goaler_DeleteFile=-1
 End If
End Function 
Function CopyFiles(TempSource,TempEnd)
 IF FSO.FileExists(TempEnd) THEN
  Response.Write "目标文件 <b><font color=red>" & TempEnd & "</font></b> 已存在，请先删除!"
  SET FSO=Nothing
  Exit Function
 End IF
 If Not FSO.FileExists(TempSource) Then
  Response.Write "源文件 <b><font color=red>"&TempSource&"</font></b> 不存在!"
  Set FSO=Nothing
  Exit Function
 End If
 FSO.CopyFile TempSource,TempEnd
 Response.Write "已经成功复制文件 <b><font color=red>"&TempSource&"</font></b> 到 <b><font color=red>"&TempEnd&"</font></b>"
End Function


 IF action="compact" THEN
 
  Response.write "<br><div align=center> 压缩数据库可以去除冗余数据，为数据库减肥，同时也能提高速度，此功能无须经常使用！</div><div align=center><a href=?action=compact2>点此压缩数据库！</a></div>"
 response.end
  
   ELSEIF action="compact2" THEN
  
  conn.close
  set conn=nothing
		
  Response.Write("<p>")
  IF goaler_isFileExist(Server.Mappath(dbfolder&dbname))=1 THEN
   Response.Write "压缩数据库开始...<br>"
   Set Engine=CreateObject("JRO.JetEngine")
   Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(dbfolder&dbname), "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath(dbfolder&"temp.mdb")
   SET FSO=Server.CreateObject("Scripting.FileSystemObject")
   FSO.CopyFile Server.Mappath(dbfolder&"temp.mdb"),Server.Mappath(dbfolder&dbname)
   FSO.DeleteFile(Server.Mappath(dbfolder&"temp.mdb"))
   Set Engine=Nothing
   Response.write "数据库：<font color=red>"&Server.MapPath(dbfolder&dbname)&"</font>压缩完成！"
  ELSE
   Response.Write("数据库源文件：<b><font color=red>"&Server.Mappath(dbfolder&dbname)&"</font></b>不存在！")
  END IF
  response.end

 
 ELSEIF action="backup" THEN
 Function GetTotalSize(GetLocal,GetType)
Set FSO=Server.CreateObject("Scripting.FileSystemObject")
If Err<>0 Then
   Err.Clear
   GetTotalSize="服务器不支持FSO，查看空间失败"
Else
   Dim SiteFolder
   If GetType="Folder" Then
      Set SiteFolder=FSO.GetFolder(GetLocal) 
   Else
      Set SiteFolder=FSO.GetFile(GetLocal) 
   End If
   GetTotalSize=SiteFolder.Size
   If GetTotalSize>1024*1024 Then
      GetTotalSize=GetTotalSize/1024/1024
      If inStr(GetTotalSize,".") Then GetTotalSize=Left(GetTotalSize,inStr(GetTotalSize,".")+2)
         GetTotalSize=GetTotalSize&" MB"
      Else
         GetTotalSize=Fix(GetTotalSize/1024)&" KB"
      End If
         Set SiteFolder=Nothing
   End If
Set FSO=Nothing
End Function
	  Response.write "<br><div align=center>数据库备份可以防止以后网站出现问题后及时得到恢复，建议经常备份数据库并下载保存</div>"
	   response.write " <div align=center>当前数据库大小为:<strong>"&GetTotalSize(Server.Mappath(dbfolder&dbname),File)&"</strong> <a href=?action=backup2>点此备份开始！</a></div>"
	 
	
 response.end
  
   ELSEIF action="backup2" THEN
  BackUpTime=right(year(now()),4)&month(now())&day(now())&hour(now())&minute(now())&second(now())
  Response.Write("备份数据库开始...<br>")
  shop7zdb="sunsnu"
  CopyFiles Server.Mappath(dbfolder&dbname),Server.Mappath(dbfolder&shop7zdb & "_" & BackUpTime &".bak")
  Response.write "<br>数据库备份完成！<br>"
  Response.write "注：多次备份数据会占用较多的空间，请及时对备份的数据库进行下载备份或删除。 "
  response.end
  
 ELSEIF action="restore" THEN
  RestoreFileName=strFileName
  Response.Write("<p>")
  IF goaler_isFileExist(Server.Mappath(dbfolder&RestoreFileName))=1 THEN
   Response.Write "开始恢复数据库...<br>"
   FSO.CopyFile Server.Mappath(dbfolder&RestoreFileName),Server.Mappath(dbfolder&dbname)
   Response.write "数据库恢复完成，从备份数据库：<b><font color=red>"&Server.MapPath(dbfolder&RestoreFileName)&"</font></b> 到：<b><font color=red>"&Server.MapPath(dbfolder&dbname)&"</font></b>！"
  ELSE
   Response.Write("备份文件：<b><font color=red>"&Server.Mappath(dbfolder&RestoreFileName)&"</font></b>不存在,或者含有特殊字符！！")
  END IF
  Response.Write("<br><a href="""&Request.ServerVariables("HTTP_REFERER")&""">请点击返回</a>")
  response.end
 
   ELSEIF(action="delfile")THEN
  Response.Write("<p>")
  IF goaler_isFileExist(Server.Mappath(dbfolder&strFileName))=1 THEN
   IF goaler_DeleteFile(Server.Mappath(dbfolder&strFileName))=1 THEN
    Response.Write("已成功删除备份文件：<b><font color=red>"&Server.Mappath(dbfolder&strFileName)&"</font></b>！")
   ELSE
    Response.Write("删除备份文件：<b><font color=red>"&Server.Mappath(dbfolder&strFileName)&"</font></b> 失败！")
   END IF
  ELSE
    Response.Write("备份文件：<b><font color=red>"&Server.Mappath(dbfolder&strFileName)&"</font></b> 不存在,或者含有特殊字符！")
  END IF 
  Response.Write("<br><a href="""&Request.ServerVariables("HTTP_REFERER")&""">请点击返回</a>")
  response.end
 ELSEIF(action="backupList")THEN
  Response.Write("<p>")
  dcheck=0
  Set DataFolder=FSO.GetFolder(Server.MapPath(dbfolder))
  Set DataFileList=DataFolder.Files
   For Each DataFile in DataFileList
    If Instr(DataFile,"bak") Then
     DataFileName=DataFile.Name
     Response.Write "<font color=#FF0000>"&DataFileName&"</font>&nbsp;&nbsp;|&nbsp;&nbsp;"
     Response.Write "备份时间："&DataFile.DateCreated&"&nbsp;&nbsp;|&nbsp;&nbsp;"
     Response.Write "<a href="&dbfolder&DataFileName&">下载</a>&nbsp;|&nbsp;"
     Response.Write "<a href=mgdata.asp?action=delfile&FileName="&DataFileName&" onclick=""return confirm('确认要删除该备份文件吗？')"">删除</a>&nbsp;|&nbsp;"
     Response.Write "<a href=mgdata.asp?action=restore&FileName="&DataFileName&" onclick=""return confirm('确认要从该备份文件还原数据吗？')"">从此备份文件还原数据</a><br>"
     dcheck=1
    End If
   Next
   if dcheck=0 then Response.Write "数据库尚未备份，或备份数据已被删除!(请备份后再执行恢复)"
 END IF
END IF 
SET FSO=Nothing
%></td>
          </tr>
        </table>
				
      </td>

</tr>
</table>

