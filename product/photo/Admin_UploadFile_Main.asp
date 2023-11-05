<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="CONN.asp"-->
<!--#include file="../common/Function1.asp"-->
<!--#include file="System_Function.asp"-->
<!--#include file="CommonCls.asp"-->
<%Call adminIsLogin()%>
<%
Dim TruePath, theFolder, theSubFolder, theFile, thisfile, FileCount, TotalSize, TotalSize_Page
Dim TotalUnit, strTotalUnit, PageUnit, strPageUnit
Dim StrFileType, strFiles,FoundErr,ErrMsg
Dim strDirName, tUploadDir, ShowFileStyle
Dim strPath, strPath2, strPath3

Set T=New Thumb
'获取频道相关数据
MaxPerPage=15
CurrentPage=chkclng(s("page"))
ChannelID=chkclng(s("ChannelID"))
Action=s("Action")
types=s("types")
ParentDir=s("ParentDir")
Dim req, sortBy, priorSort, curFiles, currentSlot, fileItem, reverse
Dim fname, fext, fsize, ftype, fcreate, fmod, faccess
Dim kind, minmax, minmaxSlot, temp, i, mark, j
Dim theFiles, SearchKeyword

'读取查看方式
ShowFileStyle = GetUploadFileStyle()

SearchKeyword = Trim(Request("SearchKeyword"))

ParentDir = Replace(Replace(Replace(Replace(Trim(Request("ParentDir")), "../", ""), "..\", ""), "\", "/"),"|","/")
If Left(ParentDir, 1) = "/" Then ParentDir = Right(ParentDir, Len(ParentDir) - 1)
CurrentDir = Replace(Replace(Replace(Trim(Request("CurrentDir")), "/", ""), "\", ""), "..", "")
Dim rs, sql
strFiles = LCase("gif|jpg|jpeg|jpe|bmp|png|rar|zip")
if tUploadDir<>"uploadimages" then UploadDir = "uploadimages"
'Setting(3)="/allpic/"
RootDir = pic_path & UploadDir
'RootDir = "/allpic/" & UploadDir      '2011-8-1
if tUploadDir<>"" then RootDir=RootDir&"/"&tUploadDir
strPath = replace(RootDir,"//","/")
RootDir = replace(RootDir,"//","/")
'strPath2 = RootDir
strPath2 = replace(RootDir,pic_path,"")
strPath3 = ""
If ParentDir <> "" Then
    If InStr(ParentDir, ChannelDir & "/" & UploadDir) > 0 Then
        ParentDir = Replace(ParentDir, ChannelDir & "/" & UploadDir & "/", "")
        ParentDir = Replace(ParentDir, ChannelDir & "/" & UploadDir, "")
    End If
    strPath = strPath & "/" & ParentDir
    strPath2 = strPath2 & "/" & ParentDir
    strPath3 = ParentDir
End If
If CurrentDir <> "" Then
    strPath = strPath & "/" & CurrentDir
    strPath2 = strPath2 & "/" & CurrentDir
    If ParentDir <> "" Then
        strPath3 = strPath3 & "/" & CurrentDir & "/"
    Else
        strPath3 = CurrentDir & "/"
    End If
End If
strPath = Replace(strPath, "//", "/")
strPath2 = Replace(strPath2, "//", "/")
TruePath = Server.MapPath(strPath)
strFileName = "Admin_UploadFile_Main.asp?ChannelID=" & ChannelID & "&UploadDir=" & TUploadDir

Response.Write "<html><head><title>上传文件管理</title><base target=""_self"" /><meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
Response.Write "<link href=""images/skin/Css_"&Session("Admin_Style_Num")&"/"&Session("Admin_Style_Num")&".css"" rel=""stylesheet"" type=""text/css""></head>"
Response.Write "<script type=""text/javascript"" src=""../common/Common.js""></script>"
Response.Write "<script src=""js/jquery.js""></script>"
Response.Write "<script src=""js/cenbel.box.js""></script>"
Response.Write "<SCRIPT language='javascript'>" & vbCrLf
Response.Write "function unselectall(){" & vbCrLf
Response.Write "    if(document.myform.chkAll.checked){" & vbCrLf
Response.Write " document.myform.chkAll.checked = document.myform.chkAll.checked&0;" & vbCrLf
Response.Write "    }" & vbCrLf
Response.Write "}" & vbCrLf

Response.Write "function CheckAll(form){" & vbCrLf
Response.Write "  for (var i=0;i<form.elements.length;i++)" & vbCrLf
Response.Write "    {" & vbCrLf
Response.Write "    var e = form.elements[i];" & vbCrLf
Response.Write "    if (e.Name != 'chkAll')" & vbCrLf
Response.Write "       e.checked = form.chkAll.checked;" & vbCrLf
Response.Write "    }" & vbCrLf
Response.Write "}" & vbCrLf
Response.Write "function preloadImg(src) {" & vbCrLf
Response.Write "  var img=new Image();" & vbCrLf
Response.Write "  img.src=src" & vbCrLf
Response.Write "}" & vbCrLf
Response.Write "preloadImg('Images/admin_upload_open.gif');" & vbCrLf
Response.Write "" & vbCrLf
Response.Write "var displayBar=false;" & vbCrLf
Response.Write "function switchBar(obj) {" & vbCrLf
Response.Write "  if (displayBar) {" & vbCrLf
Response.Write "    parent.frame.cols='0,*';" & vbCrLf
Response.Write "    displayBar=false;" & vbCrLf
Response.Write "    obj.src='Images/admin_upload_open.gif';" & vbCrLf
Response.Write "    obj.title='打开左边文件目录树型导航';" & vbCrLf
Response.Write "  } else {" & vbCrLf
Response.Write "    parent.frame.cols='160,*';" & vbCrLf
Response.Write "    displayBar=true;" & vbCrLf
Response.Write "    obj.src='Images/admin_upload_close.gif';" & vbCrLf
Response.Write "    obj.title='关闭左边文件目录树型导航';" & vbCrLf
Response.Write "  }" & vbCrLf
Response.Write "}" & vbCrLf
Response.Write "</script>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0' onload='switchBar(this)'>" & vbCrLf
If SearchKeyword <> "" Then
    Response.Write "    <form name='myform' method='Post' action='" & strFileName & "&ParentDir=" & ParentDir & "&CurrentDir=" & CurrentDir & "&SearchKeyword=" & SearchKeyword & "'>" & vbCrLf
Else
    Response.Write "    <form name='myform' method='Post' action='" & strFileName & "&ParentDir=" & ParentDir & "&CurrentDir=" & CurrentDir & "'>" & vbCrLf
End If
Response.Write "<table width='98%' border='0' align='center' cellpadding='3' cellspacing='1' class='table'>" & vbCrLf

Response.Write "  <tr class='hback'> "
Response.Write "    <td width='0' height='30'>"
Response.Write "    <table width='100%' border='0' align='center' cellpadding='3' cellspacing='1'><tr><td height='22'>"
Response.Write "      &nbsp;<img onclick='switchBar(this)' src='Images/admin_upload_open.gif' title='打开左边文件目录树型导航' style='cursor:hand'></td><td height='22'>" & vbCrLf
Response.Write "</td><td>"
If ShowFileStyle = 1 Then
    Response.Write "<a href='Admin_UploadFile_Style.asp?ShowFileStyle=2'>缩略图方式 </a>" & vbCrLf
Else
    Response.Write "<a href='Admin_UploadFile_Style.asp?ShowFileStyle=1'>列表方式</a>" & vbCrLf
End If
 'Response.Write " | <a href='Admin_UploadFile_Clear.asp'>清理图片</a>" & vbCrLf
Response.Write "    </td></tr></table>"
Response.Write "    </td>"

Response.Write "<td height='30'>"
Response.Write "    <table width='100%' border='0' align='center' cellpadding='1' cellspacing='1'><tr><td height='22' align='right'>"
Response.Write "&nbsp; 搜索当前目录文件：</td><td height='22'><input type='text' name='SearchKeyword' id='SearchKeyword' size='18' value=''>&nbsp;<input type='submit' name='submit1' value=' 搜索 '>&nbsp;&nbsp;&nbsp;&nbsp;"
Response.Write "    <button name="" type=""utton"" onclick='UpFile();' style='width:160px;font-weight:bold;' />点这批量文件上传</button></td></tr></table></td>"
Response.Write "  </tr>"
Response.Write "</table>" & vbCrLf
If IsObjInstalled("Scripting.FileSystemObject") = False Then
    Response.Write "<b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能</font></b>"
    Response.End
else
	Set fso = Server.CreateObject("Scripting.FileSystemObject")	
End If
Select Case Action
Case "Del"
   Call DelFiles
Case "DelThisFolder"
    Call DelThisFolder
Case "DelCurrentDir"
    Call DelCurrentDir
Case "DelAll"
    Call DelAll
Case "DoAddWatermark"
    Call DoAddWatermark
Case "DoAddWatermark_CurrentDir"
    Call DoAddWatermark_CurrentDir
Case Else
    Call main
End Select
If FoundErr = True Then
    Call ShowError(ErrMsg)
End If

If currentSlot > -1 And FoundErr = False Then
    Response.Write "<Script Language=""JavaScript"">" & vbCrLf
    Response.Write "setTimeout('Change()',1000);"
    Response.Write "function Change(){"
    Response.Write "var Sort=document.getElementById(""Sort" & sortBy & """);" & vbCrLf
    If reverse Then
        Response.Write "    Sort.src=""Images/Calendar_Down.gif"";" & vbCrLf
    Else
        Response.Write "    Sort.src=""Images/Calendar_Up.gif"";" & vbCrLf
    End If
    Response.Write "    Sort.style.display="""";    " & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</Script>" & vbCrLf
End If
Response.Write "</body></html>"
Call CloseConn

Sub main()
    Dim Add2Array
	 If fso.FolderExists(TruePath) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到文件夹！请上传文件后再进行管理！</li>"
        Exit Sub
    End If

    Response.Write "<Script Language=""JavaScript"">" & vbCrLf
    Response.Write "function reSort(which)" & vbCrLf
    Response.Write "{" & vbCrLf
    If SearchKeyword <> "" Then
        Response.Write "document.myform.SearchKeyword.value = '" & SearchKeyword & "';" & vbCrLf
    End If
	Response.Write "document.myform.Action.value = which;" & vbCrLf
    Response.Write "document.myform.sortby.value = which;" & vbCrLf
    Response.Write "document.myform.submit();" & vbCrLf
    Response.Write "}" & vbCrLf
	Response.Write "function UpFile()"
	Response.Write "{"
	Response.Write "  OpenWindow('fupload/upload.asp?Path="&CurrentDir&"&ChannelID="&ChannelID&"',370,355,window);"
	'Response.Write "PopupCenterIframe('多文件上传','fupload/upload.html',468,430,'no');"
	Response.Write("window.location.href=window.location.href;")
	Response.Write("window.location.reload;")
	Response.Write "}"
    Response.Write "</Script>" & vbCrLf
    'Dim FolderCount

    req = Trim(Request("sortBy"))
    If Len(req) < 1 Or req = "-1" Then
        sortBy = 5
    Else
        sortBy = CInt(req)
    End If
    req = Request("priorSort")
    If Len(req) < 1 Or req = "-1" Then
        priorSort = 5
    Else
        priorSort = CInt(req)
    End If
    '设置倒序
    If sortBy = priorSort Then
        reverse = True
        priorSort = -1
    Else
        reverse = False
        priorSort = sortBy
    End If
    Set theFolder = fso.GetFolder(TruePath)
    Set curFiles = theFolder.Files
    

    ReDim theFiles(500)
    currentSlot = -1
    
    For Each fileItem In curFiles
        Add2Array = False
        fname = fileItem.name
        If SearchKeyword <> "" Then
            If InStr(LCase(fname), LCase(SearchKeyword)) > 0 Then
                Add2Array = True
            End If
        Else
            Add2Array = True
        End If
        If Add2Array = True Then
            fext = InStrRev(fname, ".")
            If fext < 1 Then fext = "" Else fext = Mid(fname, fext + 1)
            ftype = fileItem.Type
            fsize = fileItem.size
            fcreate = fileItem.DateCreated
            fmod = fileItem.DateLastModified
            faccess = fileItem.DateLastAccessed
            currentSlot = currentSlot + 1
            If currentSlot > UBound(theFiles) Then
                ReDim Preserve theFiles(currentSlot + 99)
            End If

            theFiles(currentSlot) = Array(fname, fext, fsize, ftype, fcreate, fmod, faccess)
        End If
    Next

    If currentSlot > -1 Then
        FileCount = currentSlot ' 文件数量
        ReDim Preserve theFiles(currentSlot)


        If VarType(theFiles(0)(sortBy)) = 8 Then
            If reverse Then kind = 1 Else kind = 2
        Else
            If reverse Then kind = 3 Else kind = 4
        End If
        For i = FileCount To 0 Step -1
            minmax = theFiles(0)(sortBy)
            minmaxSlot = 0
            For j = 1 To i
                Select Case kind
                    Case 1
                    mark = (StrComp(theFiles(j)(sortBy), minmax, vbTextCompare) < 0)
                    Case 2
                    mark = (StrComp(theFiles(j)(sortBy), minmax, vbTextCompare) > 0)
                    Case 3
                    mark = (theFiles(j)(sortBy) < minmax)
                    Case 4
                    mark = (theFiles(j)(sortBy) > minmax)
                End Select
                If mark Then
                    minmax = theFiles(j)(sortBy)
                    minmaxSlot = j
                End If
            Next
            If minmaxSlot <> i Then
                temp = theFiles(minmaxSlot)
                theFiles(minmaxSlot) = theFiles(i)
                theFiles(i) = temp
            End If
        Next
    Else
        FileCount = 0
    End If
    

    If ShowFileStyle = 1 Then
        Call ShowFileDetail
    Else
        Call ShowFileThumb
    End If
End Sub

Sub ShowFileThumb()
    If SearchKeyword = "" Then
        Response.Write "<table width='98%' cellpadding='3' cellspacing='1' class='table' align='center'><tr height='22' class='hback'><td align='left' >当前目录：" & RootDir
        If ParentDir <> "" Then
            Response.Write "/" & ParentDir
        End If
        If CurrentDir <> "" Then
            Response.Write "/" & CurrentDir
        End If
        'Response.Write "</td>" & vbCrLf
        'Response.Write "    <td align='right'>" & vbCrLf
        If CurrentDir <> "" Then
            If ParentDir <> "" Then
                If InStrRev(ParentDir, "/") > 0 Then
                    Response.Write " <a href='" & strFileName & "&ParentDir=" & Left(ParentDir, InStrRev(ParentDir, "/") - 1)
                    Response.Write "&CurrentDir=" & Mid(ParentDir, InStrRev(ParentDir, "/") + 1)
                Else
                    Response.Write " <a href='" & strFileName & "&ParentDir=&CurrentDir=" & ParentDir
                End If
            Else
                Response.Write " <a href='" & strFileName
            End If
            Response.Write "'> ↑返回上级目录</a>"
        End If
        Response.Write "</td></tr></table>" & vbCrLf
        Response.Write "<table width='98%' cellpadding='2' cellspacing='1' class='table' align='center'><tr class='back' height='22'><td colspan='20' class='xingmu'>子目录导航：" & vbCrLf
        Response.Write "</td></tr><tr class='hback'>"
        Dim FolderCount
        Set theFolder = fso.GetFolder(TruePath)
        For Each theSubFolder In theFolder.SubFolders
            If ParentDir <> "" Then
                Response.Write "<td><a href='" & strFileName & "&ParentDir=" & ParentDir & "/" & CurrentDir & "&CurrentDir=" & theSubFolder.name & "'>" & theSubFolder.name & "</a></td>"
            Else
                Response.Write "<td><a href='" & strFileName & "&ParentDir=" & CurrentDir & "&CurrentDir=" & theSubFolder.name & "'>" & theSubFolder.name & "</a></td>"
            End If
            FolderCount = FolderCount + 1
            If FolderCount Mod 10 = 0 Then Response.Write "</td><tr class='tdbg'>"
        Next
        Response.Write "</tr></table>" & vbCrLf
    Else
		Response.Write "<table width='98%' cellpadding='3' cellspacing='1' class='table' align='center'><tr height='22' class='hback'><td align='left' >"
        Response.Write "&gt;&gt;&nbsp;当前目录文件名中含有的 <font color='red'>" & SearchKeyword & "</font> 文件"
		Response.Write "</td></tr></table>" & vbCrLf
    End If
    Response.Write "    <table width='98%' border='0' cellspacing='1' cellpadding='3' class='table' align='center'>" & vbCrLf
    Response.Write "    <tr class='back'>" & vbCrLf
    If currentSlot > -1 Then
        Response.Write "    <td height='18' class='xingmu'>排序方式：&nbsp;&nbsp;<a href=""javascript:reSort(0);"">文件名&nbsp;<img src='Images/Calendar_Down.gif' border='0' style='display:none' id='Sort0'></a>" & vbCrLf
        'Response.Write "   <a href=""javascript:reSort(1);"">扩展名</a>" & vbCrLf
        Response.Write "    &nbsp;&nbsp;<a href=""javascript:reSort(2);"">大小&nbsp;<img src='Images/Calendar_Down.gif' border='0' style='display:none' id='Sort2'></a>" & vbCrLf
        Response.Write "    &nbsp;&nbsp;<a href=""javascript:reSort(3);"">类型&nbsp;<img src='Images/Calendar_Down.gif' border='0' style='display:none' id='Sort3'></a>" & vbCrLf
        'Response.Write "   <a href=""javascript:reSort(4);"">建立时间</a>" & vbCrLf
        Response.Write "    &nbsp;&nbsp;<a href=""javascript:reSort(5);"">上次修改时间&nbsp;<img src='Images/Calendar_Down.gif' border='0' style='display:none' id='Sort5'></a></td>" & vbCrLf
    Else
        Response.Write "    <td height='18' class='xingmu'></td>" & vbCrLf
    End If
    Response.Write "    <td align='right' class='xingmu'>" & vbCrLf
    Response.Write "</td></tr>" & vbCrLf
    If currentSlot = -1 Then
        Response.Write "<tr class='hback'><td align='center' colspan='2'><br><br>当前目录下没有任何文件！<br><br></td>"
        Response.Write " </tr>"
        Response.Write "</table>" & vbCrLf
    Else
        strFileName = strFileName & "&ParentDir=" & ParentDir & "&CurrentDir=" & CurrentDir
        TotalSize = 0
        TotalUnit = 1
        For Each theFile In theFolder.Files
            If TotalUnit = 1 Then
                TotalSize = TotalSize + theFile.size / 1024
            ElseIf TotalUnit = 2 Then
                TotalSize = TotalSize + theFile.size / 1024 / 1024
            ElseIf TotalUnit = 3 Then
                TotalSize = TotalSize + theFile.size / 1024 / 1024 / 1024
            End If
            If TotalSize > 1024 Then
                TotalSize = TotalSize / 1024
                TotalUnit = TotalUnit + 1
            End If
            If TotalUnit = 1 Then
                strTotalUnit = "KB"
            ElseIf TotalUnit = 2 Then
                strTotalUnit = "MB"
            ElseIf TotalUnit = 3 Then
                strTotalUnit = "GB"
            End If
        Next
        TotalSize = Round(TotalSize, 2)
        totalPut = FileCount + 1
        If CurrentPage < 1 Then
            CurrentPage = 1
        End If
        If (CurrentPage - 1) * MaxPerPage > totalPut Then
            If (totalPut Mod MaxPerPage) = 0 Then
                CurrentPage = totalPut \ MaxPerPage
            Else
                CurrentPage = totalPut \ MaxPerPage + 1
            End If
        End If
        If CurrentPage > 1 Then
            If (CurrentPage - 1) * MaxPerPage >= totalPut Then
                CurrentPage = 1
            End If
        End If
        Dim c
        Dim theFileName, tUsed, FileNum
        FileNum = 0
        TotalSize_Page = 0
        PageUnit = 1
        Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0' align='center'>"
        Response.Write "  <tr>"
        Response.Write "     <td colspan='2'><table width='98%' border='0' cellpadding='3' cellspacing='1' class='table' align='center'>"
        Response.Write "  <tr>" & vbCrLf
        For i = 0 To FileCount
            c = c + 1
            If FileNum >= MaxPerPage Then
                Exit For
            ElseIf c > MaxPerPage * (CurrentPage - 1) Then
                Response.Write "    <td>"
                Response.Write "      <table width='100%' border='0' cellpadding='3' cellspacing='1' class='table' align='center'>"
                Response.Write "        <tr class='hback'>"
                Response.Write "          <td colspan='2' align='center'>"
                theFileName = strPath & "/" & theFiles(i)(0)
				If types="select" Then 
				Response.Write "<a href='" & theFileName & "'>"
				else
                Response.Write "<a href='#' onClick=""window.returnValue='" & strPath2 & "/" & theFiles(i)(0) & "';window.close();"">"
				end if
                'StrFileType = LCase(Mid(theFile.Name, InStrRev(theFile.Name, ".") + 1))
                'Select Case StrFileType
                Select Case LCase(theFiles(i)(1))
                Case "gif", "jpg", "jpeg", "jpe", "bmp", "png"
                    Response.Write "<img src='" & theFileName & "'"
                Case "swf"
                    Response.Write "<img src='images/filetype_flash.gif'"
                Case "wmv", "avi", "asf", "mpg"
                    Response.Write "<img src='images/filetype_media.gif'"
                Case "rm", "ra", "ram"
                    Response.Write "<img src='images/filetype_rm.gif'"
                Case "rar", "zip", "exe"
                    Response.Write "<img src='images/filetype_" & theFiles(i)(1) & ".gif'"
                Case Else
                    Response.Write "<img src='images/filetype_other.gif'"
                End Select
                Response.Write " width='80' height='80' border='0'"
                If InStr(strFiles, LCase(theFiles(i)(3))) > 0 Then
                    tUsed = True
                Else
                    tUsed = False
                End If
                If tUsed = True Then
                    Response.Write " border='0' alt='文 件 名：" & theFiles(i)(0) & vbCrLf & "文件大小：" & Round(theFiles(i)(2) / 1024) & " KB" & vbCrLf & "文件类型：" & theFiles(i)(3) & vbCrLf & "修改时间：" & theFiles(i)(5) & "'>"
                Else
                    Response.Write " border='2' alt='文 件 名：" & theFiles(i)(0) & vbCrLf & "文件大小：" & Round(theFiles(i)(2) / 1024) & " KB" & vbCrLf & "文件类型：" & theFiles(i)(3) & vbCrLf & "修改时间：" & theFiles(i)(5) & "'>"
                End If

                Response.Write "</a>"
                Response.Write "          </td>"
                Response.Write "        </tr>" & vbCrLf
                Response.Write "        <tr class='hback'>"
                'Response.Write "          <td align='right'>文件名：</td>"
                Response.Write "          <td align='center'>"
                If tUsed = True Then
                    Response.Write "<input name='FileName' type='checkbox' class='checkbox' id='FileName' value='" & theFiles(i)(0) & "' onclick='unselectall()'"
                If tUsed = False Then Response.Write " checked"
                Response.Write "> 选中<a href='" & theFileName & "' target='_blank'>" & CutStr(theFiles(i)(0)) & "</a>"
                Else
                    Response.Write "<input name='FileName' type='checkbox' class='checkbox' id='FileName' value='" & theFiles(i)(0) & "' onclick='unselectall()'"
                'If tUsed = False Then Response.Write " checked"
                Response.Write "><a href='" & theFileName & "' target='_blank' title='看原图'><font color=red>" & CutStr(theFiles(i)(0)) & "</font></a>"
                End If

                Response.Write "       </td>"
                Response.Write "        </tr>" & vbCrLf
                Response.Write "      </table>"
                Response.Write "    </td>" & vbCrLf
                FileNum = FileNum + 1
                If FileNum Mod 5 = 0 Then Response.Write "</td><tr class='tdbg'>"
                If PageUnit = 1 Then
                    TotalSize_Page = TotalSize_Page + theFiles(i)(2) / 1024
                ElseIf PageUnit = 2 Then
                    TotalSize_Page = TotalSize_Page + theFiles(i)(2) / 1024 / 1024
                ElseIf PageUnit = 3 Then
                    TotalSize_Page = TotalSize_Page + theFiles(i)(2) / 1024 / 1024 / 1024
                End If
                If TotalSize_Page > 1024 Then
                    TotalSize_Page = TotalSize_Page / 1024
                    PageUnit = PageUnit + 1
                End If
                If PageUnit = 1 Then
                    strPageUnit = "KB"
                ElseIf PageUnit = 2 Then
                    strPageUnit = "MB"
                ElseIf PageUnit = 3 Then
                    strPageUnit = "GB"
                End If
            End If
        Next
        TotalSize_Page = Round(TotalSize_Page, 2)

        Response.Write "  </tr>"
        Response.Write "</table>"
        Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
        Response.Write "  <tr>"
        Response.Write "    <td width='200' height='30'>&nbsp;&nbsp;&nbsp;&nbsp;<input name='chkAll' type='checkbox' class='checkbox' id='chkAll' onclick='CheckAll(this.form);' value='checkbox'> 选中本页所有文件</td><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
        Response.Write "    <td><input name='Action' type='hidden' id='Action' value=''><input name='UploadDir' type='hidden' value='" & TUploadDir & "'><input name='CurrentDir' type='hidden' value='" & CurrentDir & "'>"
        Response.Write "&nbsp;&nbsp;<button type='submit' name='Submit3' onClick=""document.myform.Action.value='Del';return confirm('你真的要删除此文件吗!');"" title='删除' >删除</button>"
        'If IsObjInstalled("Persits.Jpeg") = True Then
        '    Response.Write "&nbsp;&nbsp;<button type='submit' name='Submit3' onClick=""document.myform.Action.value='DoAddWatermark'"" title='给选中的图片添加水印' >给选中的图片添加水印</button>&nbsp;&nbsp;<button type='submit' name='Submit4' onClick=""document.myform.Action.value='DoAddWatermark_CurrentDir'""  title='给当前目录添加水印' >给当前目录添加水印</button>"
        'End If
       
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "</table>" & vbCrLf

        Response.Write "<input type='hidden' name='priorsort' value='" & priorSort & "'>" & vbCrLf
        Response.Write "<input type='hidden' name='sortby' value='-1'>" & vbCrLf


        Response.Write "</td></form></tr></table>" & vbCrLf
        Response.Write showpage2(strFileName, totalPut, MaxPerPage)
        Response.Write "<div align='center'>本页共显示 <b>" & FileNum & "</b> 个文件，占用 <b>" & TotalSize_Page & "</b> " & strPageUnit & "</div>"

    End If

End Sub

Sub ShowFileDetail()
    If SearchKeyword <> "" Then
		Response.Write "<table  width='98%' border='0' cellpadding='3' cellspacing='1' align='center' class='table'><tr class='hback'><td>"
        Response.Write "<br>&gt;&gt;&nbsp;当前目录文件名中含有的 <font color='red'>" & SearchKeyword & "</font> 文件"
		 Response.Write "</td></tr></table>" & vbCrLf
    End If
    Response.Write "<table  width='98%' border='0' cellpadding='3' cellspacing='1' align='center' class='table'><tr class='hback'><td style=""font-size:14px;FONT-FAMILY: '宋体'; color:#0000ff; "">当前目录：" 
	showdir = mid(RootDir,2)
	showdirshu = instr(showdir,"/")+1
	response.write mid(showdir,showdirshu)
    If ParentDir <> "" Then
        Response.Write "/" & ParentDir
    End If
    If CurrentDir <> "" Then
        Response.Write "/" & CurrentDir & " "
        If ParentDir <> "" Then
            If InStrRev(ParentDir, "/") > 0 Then
                Response.Write "&nbsp; <a href='" & strFileName & "&ParentDir=" & replace(Left(ParentDir, InStrRev(ParentDir, "/") - 1),"/","|")
                Response.Write "&CurrentDir=" & Mid(ParentDir, InStrRev(ParentDir, "/") + 1)
            Else
                Response.Write "&nbsp; <a href='" & strFileName & "&ParentDir=&CurrentDir=" & ParentDir
            End If
        Else
            Response.Write "&nbsp; <a href='" & strFileName
        End If
        Response.Write "'>↑返回上级目录</a>"
	else
		Response.Write "&nbsp;&nbsp;&nbsp;(已是顶级目录)"
    End If
    Response.Write "</td></tr></table>" & vbCrLf

    Response.Write "    <table width='98%' border='0' cellpadding='3' cellspacing='1' align='center' class='table'>" & vbCrLf

    Response.Write "    <tr class='xingmu'>" & vbCrLf
    Response.Write "    <td height='18' class='title0' onmouseout=""this.className='title0'"" onmouseover=""this.className='tdbgmouseover1'"">&nbsp;&nbsp;<a href=""javascript:reSort(0);"">文件名&nbsp;&nbsp;<img src='Images/Calendar_Down.gif' border='0' style='display:none' id='Sort0'></a></td>" & vbCrLf
    'Response.Write "   <a href=""javascript:reSort(1);"">扩展名</a>" & vbCrLf
    Response.Write "    <td width='80' align=""right"" class='title0' onmouseout=""this.className='title0'"" onmouseover=""this.className='tdbgmouseover1'""><a href=""javascript:reSort(2);"">大小&nbsp;<img src='Images/Calendar_Down.gif' border='0' style='display:none' id='Sort2'></a>&nbsp;</td>" & vbCrLf
    Response.Write "    <td width='180' class='title0' onmouseout=""this.className='title0'"" onmouseover=""this.className='tdbgmouseover1'"">&nbsp;<a href=""javascript:reSort(3);"">类型&nbsp;&nbsp;<img src='Images/Calendar_Down.gif' border='0' style='display:none' id='Sort3'></a></td>" & vbCrLf
    'Response.Write "   <a href=""javascript:reSort(4);"">建立时间</a>" & vbCrLf
    Response.Write "    <td width='140' class='title0' onmouseout=""this.className='title0'"" onmouseover=""this.className='tdbgmouseover1'""><a href=""javascript:reSort(5);"">上次修改时间&nbsp;&nbsp;<img src='Images/Calendar_Down.gif' border='0' style='display:none' id='Sort5'></a></td>" & vbCrLf
    Response.Write "    <td width='30' align='center' class='title0'>操作&nbsp;</td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    If sortBy <> 0 Then
        Call ShowFileDetail_Fil
        If SearchKeyword = "" Then
            'Response.Write "<tr class='hback'><td height=1></td></tr>" & vbCrLf
            'Call ShowFileDetail_fol
        End If
    Else
        If SearchKeyword = "" Then
           ' Call ShowFileDetail_fol
            'Response.Write "<tr class='hback'><td height=1></td></tr>" & vbCrLf
        End If
        Call ShowFileDetail_Fil
    End If
    Response.Write "    </table>" & vbCrLf
    Response.Write "<input type='hidden' name='priorsort' value='" & priorSort & "'>" & vbCrLf
    Response.Write "<input type='hidden' name='sortby' value='-1'>" & vbCrLf
    If currentSlot > -1 Then
        Call ShowJS_Tooltip
        Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
        Response.Write "  <tr>"
        Response.Write "    <td width='200' height='30'>&nbsp;&nbsp;&nbsp;&nbsp;<input name='chkAll' type='checkbox' class='checkbox' id='chkAll' onclick='CheckAll(this.form);' value='checkbox'> 选中本页所有文件</td><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
        Response.Write "    <td><input name='Action' type='hidden' id='Action' value=''><input name='UploadDir' type='hidden' value='" & TUploadDir & "'><input name='CurrentDir' type='hidden' value='" & CurrentDir & "'>"
        Response.Write "&nbsp;&nbsp;<button type='submit' name='Submit3' onClick=""document.myform.Action.value='Del';return confirm('你真的要删除此文件吗!');"" title='删除' >删除</button>"
        'If IsObjInstalled("Persits.Jpeg") = True Then
        '    Response.Write "&nbsp;&nbsp;<button type='submit' name='Submit3' onClick=""document.myform.Action.value='DoAddWatermark'"" title='给选中的图片添加水印' >给选中的图片添加水印</button>&nbsp;&nbsp;<button type='submit' name='Submit4' onClick=""document.myform.Action.value='DoAddWatermark_CurrentDir'""  title='给当前目录添加水印' >给当前目录添加水印</button>"
        'End If
       
        Response.Write "    </td>"
        Response.Write "  </tr>"
        Response.Write "</table>" & vbCrLf
    End If

    
    Response.Write "    </form>" & vbCrLf

End Sub

Function GetUploadFileStyle()
    ShowFileStyle = Request.Cookies("ShowFileStyle")
    If ShowFileStyle = "" Or Not IsNumeric(ShowFileStyle) Then
        ShowFileStyle = 1
    Else
        ShowFileStyle = Int(ShowFileStyle)
    End If
    GetUploadFileStyle = ShowFileStyle
End Function



Sub ShowFileDetail_Fil()
    If currentSlot > -1 Then

        For i = 0 To FileCount
            Response.Write "<tr onmouseout=""this.className='hback'"" onmouseover=""this.className='tdbg1'"" class='hback'>" & vbCrLf
            If InStr(strFiles, LCase(theFiles(i)(0))) > 0 Then
                'Response.Write " Title='文 件 名：" & theFiles(i)(0) & vbCrLf & "文件大小：" & Round(theFiles(i)(2) / 1024) & " K" & vbCrLf & "文件类型：" & theFiles(i)(3) & vbCrLf & "修改时间：" & theFiles(i)(5) & "'>"
                Response.Write "          <td align='left'><input name='FileName' type='checkbox' class='checkbox' id='FileName' value='" & theFiles(i)(0) & "' onclick='unselectall()'"
                Response.Write ">"
            Else
                'Response.Write " Title='无用的上传文件" & vbCrLf & "文 件 名：" & theFiles(i)(0) & vbCrLf & "文件大小：" & Round(theFiles(i)(2) / 1024) & " K" & vbCrLf & "文件类型：" & theFiles(i)(3) & vbCrLf & "修改时间：" & theFiles(i)(5) & "'>"
                Response.Write "          <td align='left'><input name='FileName' type='checkbox' class='checkbox' id='FileName' value='" & theFiles(i)(0) & "' onclick='unselectall()'"
               ' Response.Write " checked"
                Response.Write ">"
            End If

            Select Case LCase(theFiles(i)(1))
            Case "jpeg", "jpe", "bmp", "png"
                Response.Write "<img src='images/Folder/img.gif'>"
            Case "swf"
                Response.Write "<img src='images/Folder/Ftype_flash.gif'>"
            Case "dll", "vbp"
                Response.Write "<img src='images/Folder/sys.gif'>"
            Case "wmv", "avi", "asf", "mpg"
                Response.Write "<img src='images/Folder/Ftype_media.gif'>"
            Case "rm", "ra", "ram"
                Response.Write "<img src='images/Folder/Ftype_rm.gif'>"
            Case "rar", "zip"
                Response.Write "<img src='images/Folder/zip.gif'>"
            Case "xml", "txt", "exe", "doc", "html", "htm", "jpg", "gif", "xls", "asp"
                Response.Write "<img src='images/Folder/" & theFiles(i)(1) & ".gif'>"
            Case Else
                Response.Write "<img src='images/Folder/other.gif'>"
            End Select
			if types="select" then
			Response.Write "<a href='" & strPath & "/" & theFiles(i)(0) & "'>"
			else
            Response.Write "<a href='#' onClick=""window.returnValue='" & strPath2 & "/" & theFiles(i)(0) &"';window.close();"">"
			end if
			Response.Write "<span onmouseover=""ShowADPreview('" & FixJs(GetFileContent(strPath & "/" & theFiles(i)(0), theFiles(i)(1))) & "')"" onmouseout=""hideTooltip('dHTMLADPreview')"">" & vbCrLf
            Response.Write theFiles(i)(0) & "</span></a></td>" & vbCrLf
            Response.Write " <td width='80' align='right'>" & FormatNumber(theFiles(i)(2) / 1024, 0, vbTrue, vbFalse, vbTrue) & " KB</td>" & vbCrLf
            Response.Write " <td width='180'>&nbsp;" & CutStr(theFiles(i)(3)) & "</td>" & vbCrLf
            Response.Write " <td width='140'>" & theFiles(i)(5) & "</td>" & vbCrLf
            Response.Write "<td width='30' align='center'><a href='" & strFileName & "&ParentDir=" & ParentDir & "&CurrentDir=" & CurrentDir & "&Action=Del&FileName=" & theFiles(i)(0) & "' onclick=""return confirm('你真的要删除此文件吗!');"">删除</a>&nbsp;"
            Response.Write "</td></tr>" & vbCrLf
        Next
    End If
End Sub

Function GetFileContent(ByVal sPath, sType)
    If IsNull(sPath) Or sPath = "" Then
        GetFileContent = "&nbsp;此文件非图片或动画，无预览&nbsp;"
        Exit Function
    End If
    If IsNull(sType) Or sType = "" Then
        GetFileContent = "&nbsp;此文件非图片或动画，无预览&nbsp;"
        Exit Function
    End If

    If Not fso.FileExists(Server.MapPath(sPath)) Then
        GetFileContent = "&nbsp;此文件不存在&nbsp;"
        Exit Function
    End If
    Dim strFile

    Select Case LCase(sType)
    Case "jpeg", "jpe", "bmp", "png", "jpg", "gif"
        strFile = "<img src='" & sPath & "'"
        strFile = strFile & " width='200'"
        strFile = strFile & " height='120'"
        strFile = strFile & " border='0'>"
    Case "wmv", "avi", "asf", "mpg", "rm", "ra", "ram", "swf"
        strFile = "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0'"
        strFile = strFile & " width='200'"
        strFile = strFile & " height='120'"
        strFile = strFile & "><param name='movie' value='" & sPath & "'>"
        strFile = strFile & "<param name='wmode' value='transparent'>"
        strFile = strFile & "<param name='quality' value='autohigh'>"
        strFile = strFile & "<embed src='" & sPath & "' quality='autohigh'"
        strFile = strFile & " pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'"
        strFile = strFile & " wmode='transparent'"
        strFile = strFile & " width='200'"
        strFile = strFile & " height='120'"
        strFile = strFile & "></embed></object>"
    Case Else
        strFile = "&nbsp;此文件非图片或动画，无预览&nbsp;"
    End Select

    GetFileContent = strFile
End Function

Sub ShowFileDetail_fol()
    Response.Write "<tr>"
    For Each theSubFolder In theFolder.SubFolders
        If ParentDir <> "" Then
            Response.Write "<td height='18'>&nbsp;&nbsp;<img src='Images/Folder/folderclosed.gif'><a href='" & strFileName & "&ParentDir=" & replace(ParentDir,"/","|") & "|" & CurrentDir & "&CurrentDir=" & theSubFolder.name & "'>" & theSubFolder.name & "</a></td>"
        Else
            Response.Write "<td height='18'>&nbsp;&nbsp;<img src='Images/Folder/folderclosed.gif'><a href='" & strFileName & "&ParentDir=" & CurrentDir & "&CurrentDir=" & theSubFolder.name & "'>" & theSubFolder.name & "</a></td>"
        End If
        Response.Write "<td width='50' align=""right"">&nbsp;</td>"
        Response.Write "<td width='180'>&nbsp;文件夹</td>"
        Response.Write "<td width='140'>" & theSubFolder.DateLastModified & "</td>"

        Response.Write "<td width='30' align='center'><a href='" & strFileName & "&ParentDir=" & ParentDir & "/" & CurrentDir & "&CurrentDir=" & theSubFolder.name & "&Action=DelThisFolder' onclick=""return confirm('你真的要删除此文件夹及里面的文件吗!');"">删除</a>&nbsp;"
        Response.Write "</td></tr><tr>"
        
    Next
End Sub

Function FixJs(Str)
    If Str <> "" Then
        Str = Replace(Str, "&#39;", "'")
        Str = Replace(Str, "\", "\\")
        Str = Replace(Str, Chr(34), "\""")
        Str = Replace(Str, Chr(39), "\'")
        Str = Replace(Str, Chr(13), "\n")
        Str = Replace(Str, Chr(10), "\r")
        Str = Replace(Str, "'", "&#39;")
        Str = Replace(Str, """", "&quot;")
    End If
    FixJs = Str
End Function

Sub DelFiles()
    Dim whichfile, arrFileName, i
    whichfile = Trim(Request("FileName"))
    If whichfile = "" Then Exit Sub
    If InStr(whichfile, ",") > 0 Then
        arrFileName = Split(whichfile, ",")
        For i = 0 To UBound(arrFileName)
            whichfile = Server.MapPath(strPath & "/" & Trim(arrFileName(i)))
            If fso.FileExists(whichfile) Then fso.DeleteFile whichfile
        Next
    Else
        whichfile = Server.MapPath(strPath & "/" & whichfile)
        If fso.FileExists(whichfile) Then fso.DeleteFile whichfile
    End If
    Call main
End Sub

Sub DelCurrentDir()
    Set theFolder = fso.GetFolder(Server.MapPath(strPath))
    For Each theFile In theFolder.Files
        theFile.Delete True
    Next
    Call main
End Sub

Sub DelAll()
    Set theFolder = fso.GetFolder(Server.MapPath(RootDir))
    For Each theSubFolder In theFolder.SubFolders
        theSubFolder.Delete True
    Next
    For Each theFile In theFolder.Files
        theFile.Delete True
    Next
    Call main
End Sub

Sub DelThisFolder()
    'on error resume next
    fso.DeleteFolder Server.MapPath(strPath)
    If Err.Number <> 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>删除文件（ReplaceDBContent.asp）失败，错误原因：" & Err.Description & "<br>请手动删除此文件。"
        Err.Clear
        Exit Sub
    Else
        If SearchKeyword <> "" Then
            Call Refresh1("Admin_UploadFile_Main.asp?ChannelID=" & ChannelID & "&UploadDir=" & TUploadDir & "&ParentDir=" & ParentDir & "&SearchKeyword=" & SearchKeyword,0)		
            'Response.Write "<meta http-equiv=""refresh"" content=0;url=""Admin_UploadFile_Main.asp?ChannelID=" & ChannelID & "&UploadDir=" & UploadDir & "&ParentDir=" & ParentDir & "&SearchKeyword=" & SearchKeyword & """>"
        Else
            Call Refresh1("Admin_UploadFile_Main.asp?ChannelID=" & ChannelID & "&UploadDir=" & TUploadDir & "&ParentDir=" & ParentDir,0)		
            'Response.Write "<meta http-equiv=""refresh"" content=0;url=""Admin_UploadFile_Main.asp?ChannelID=" & ChannelID & "&UploadDir=" & UploadDir & "&ParentDir=" & ParentDir & """>"
        End If
    End If
End Sub

Sub DoAddWatermark()
    Dim whichfile, arrFileName, i, bTemp
    whichfile = Trim(Request("FileName"))
    If whichfile = "" Then Exit Sub

    Dim PE_Thumb
    If InStr(whichfile, ",") > 0 Then
        arrFileName = Split(whichfile, ",")
        For i = 0 To UBound(arrFileName)
            whichfile = strPath & "/" & Trim(arrFileName(i))
			call T.AddWaterMark(whichfile)
        Next
    Else
        whichfile = strPath & "/" & whichfile
        call T.AddWaterMark(whichfile)
    End If
    If SearchKeyword <> "" Then
        Call Refresh1("Admin_UploadFile_Main.asp?ChannelID=" & ChannelID & "&UploadDir=" & TUploadDir & "&ParentDir=" & ParentDir & "&CurrentDir=" & CurrentDir & "&SearchKeyword=" & SearchKeyword,0)		
        'Response.Write "<meta http-equiv=""refresh"" content=0;url=""Admin_UploadFile_Main.asp?ChannelID=" & ChannelID & "&UploadDir=" & UploadDir & "&ParentDir=" & ParentDir & "&CurrentDir=" & CurrentDir & "&SearchKeyword=" & SearchKeyword & """>"
    Else
        Call Refresh1("Admin_UploadFile_Main.asp?ChannelID=" & ChannelID & "&UploadDir=" & TUploadDir & "&ParentDir=" & ParentDir & "&CurrentDir=" & CurrentDir,0)	
       
    End If
End Sub

Function ShowJS_Tooltip()
    Response.Write "<div id=dHTMLADPreview style='Z-INDEX: 1000; LEFT: 0px; VISIBILITY: hidden; WIDTH: 10px; POSITION: absolute; TOP: 0px; HEIGHT: 10px'></DIV>"
    Response.Write "<SCRIPT language = 'JavaScript'>" & vbCrLf
    Response.Write "<!--" & vbCrLf
    Response.Write "var tipTimer;" & vbCrLf
    Response.Write "function locateObject(n, d)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "   var p,i,x;" & vbCrLf
    Response.Write "   if (!d) d=document;" & vbCrLf
    Response.Write "   if ((p=n.indexOf('?')) > 0 && parent.frames.length)" & vbCrLf
    Response.Write "   {d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}" & vbCrLf
    Response.Write "   if (!(x=d[n])&&d.all) x=d.all[n]; " & vbCrLf
    Response.Write "   for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];" & vbCrLf
    Response.Write "   for (i=0;!x&&d.layers&&i<d.layers.length;i++) x=locateObject(n,d.layers[i].document); return x;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function ShowADPreview(ADContent)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "  showTooltip('dHTMLADPreview',event, ADContent, '#ffffff','#000000','#000000','6000')" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function showTooltip(object, e, tipContent, backcolor, bordercolor, textcolor, displaytime)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "   window.clearTimeout(tipTimer)" & vbCrLf
    Response.Write "   if (document.all) {" & vbCrLf
    Response.Write "       locateObject(object).style.top=document.body.scrollTop+event.clientY+20" & vbCrLf
    Response.Write "       locateObject(object).innerHTML='<table style=""font-family:宋体; font-size: 9pt; border: '+bordercolor+'; border-style: solid; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px; background-color: '+backcolor+'"" width=""10"" border=""0"" cellspacing=""0"" cellpadding=""0""><tr><td nowrap><font style=""font-family:宋体; font-size: 9pt; color: '+textcolor+'"">'+unescape(tipContent)+'</font></td></tr></table> '" & vbCrLf
    Response.Write "       if ((e.x + locateObject(object).clientWidth) > (document.body.clientWidth + document.body.scrollLeft)) {" & vbCrLf
    Response.Write "           locateObject(object).style.left = (document.body.clientWidth + document.body.scrollLeft) - locateObject(object).clientWidth-10;" & vbCrLf
    Response.Write "       } else {" & vbCrLf
    Response.Write "           locateObject(object).style.left=document.body.scrollLeft+event.clientX" & vbCrLf
    Response.Write "       }" & vbCrLf
    Response.Write "       locateObject(object).style.visibility='visible';" & vbCrLf
    Response.Write "       tipTimer=window.setTimeout(""hideTooltip('""+object+""')"", displaytime);" & vbCrLf
    Response.Write "       return true;" & vbCrLf
    Response.Write "   } else if (document.layers) {" & vbCrLf
    Response.Write "       locateObject(object).document.write('<table width=""10"" border=""0"" cellspacing=""1"" cellpadding=""1""><tr bgcolor=""'+bordercolor+'""><td><table width=""10"" border=""0"" cellspacing=""0"" cellpadding=""0""><tr bgcolor=""'+backcolor+'""><td nowrap><font style=""font-family:宋体; font-size: 9pt; color: '+textcolor+'"">'+unescape(tipContent)+'</font></td></tr></table></td></tr></table>')" & vbCrLf
    Response.Write "       locateObject(object).document.close()" & vbCrLf
    Response.Write "       locateObject(object).top=e.y+20" & vbCrLf
    Response.Write "       if ((e.x + locateObject(object).clip.width) > (window.pageXOffset + window.innerWidth)) {" & vbCrLf
    Response.Write "           locateObject(object).left = window.innerWidth - locateObject(object).clip.width-10;" & vbCrLf
    Response.Write "       } else {" & vbCrLf
    Response.Write "           locateObject(object).left=e.x;" & vbCrLf
    Response.Write "       }" & vbCrLf
    Response.Write "       locateObject(object).visibility='show';" & vbCrLf
    Response.Write "       tipTimer=window.setTimeout(""hideTooltip('""+object+""')"", displaytime);" & vbCrLf
    Response.Write "       return true;" & vbCrLf
    Response.Write "   } else {" & vbCrLf
    Response.Write "       return true;" & vbCrLf
    Response.Write "   }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "function hideTooltip(object) {" & vbCrLf
    Response.Write "    if (document.all) {" & vbCrLf
    Response.Write "        locateObject(object).style.visibility = 'hidden';" & vbCrLf
    Response.Write "        locateObject(object).style.left = 1;" & vbCrLf
    Response.Write "        locateObject(object).style.top = 1;" & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "    } else {" & vbCrLf
    Response.Write "        if (document.layers) {" & vbCrLf
    Response.Write "            locateObject(object).visibility = 'hide';" & vbCrLf
    Response.Write "            locateObject(object).left = 1;" & vbCrLf
    Response.Write "            locateObject(object).top = 1;" & vbCrLf
    Response.Write "            return false;" & vbCrLf
    Response.Write "        } else {" & vbCrLf
    Response.Write "            return true;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "//-->" & vbCrLf
    Response.Write "</SCRIPT>" & vbCrLf
End Function
Sub DoAddWatermark_CurrentDir()
    Dim whichfile, bTemp
    Dim PE_Thumb
    Set theFolder = fso.GetFolder(Server.MapPath(strPath))
    For Each theFile In theFolder.Files
        whichfile = strPath & "/" & theFile.name
        call T.AddWaterMark(whichfile)
    Next
    'Call main
    If SearchKeyword <> "" Then
        Response.Write "<meta http-equiv=""refresh"" content=0;url=""Admin_UploadFile_Main.asp?ChannelID=" & ChannelID & "&UploadDir=" & TUploadDir & "&ParentDir=" & ParentDir & "&CurrentDir=" & CurrentDir & "&SearchKeyword=" & SearchKeyword & """>"
    Else
        Response.Write "<meta http-equiv=""refresh"" content=0;url=""Admin_UploadFile_Main.asp?ChannelID=" & ChannelID & "&UploadDir=" & TUploadDir & "&ParentDir=" & ParentDir & "&CurrentDir=" & CurrentDir & """>"
    End If
End Sub

Function showpage2(sfilename, totalnumber, MaxPerPage)
    Dim n, i, strTemp
    If totalnumber Mod MaxPerPage = 0 Then
        n = totalnumber \ MaxPerPage
    Else
        n = totalnumber \ MaxPerPage + 1
    End If
    If SearchKeyword <> "" Then
        strTemp = "<table align='center'><form name='showpages' method='Post' action='" & sfilename & "&SearchKeyword=" & SearchKeyword & "'><tr><td>"
    Else
         strTemp = "<table align='center'><form name='showpages' method='Post' action='" & sfilename & "'><tr><td>"
    End If
    strTemp = "<table align='center'><form name='showpages' method='Post' action='" & sfilename & "'><tr><td>"
    strTemp = strTemp & "共 <b>" & totalnumber & "</b> 个文件，占用 <b>" & TotalSize & "</b> " & strTotalUnit & "&nbsp;&nbsp;&nbsp;"
    sfilename = JoinChar(sfilename)
    If CurrentPage < 2 Then
            strTemp = strTemp & "首页 上一页&nbsp;"
    Else
            strTemp = strTemp & "<a href='" & sfilename & "page=1'>首页</a>&nbsp;"
            strTemp = strTemp & "<a href='" & sfilename & "page=" & (CurrentPage - 1) & "'>上一页</a>&nbsp;"
    End If

    If n - CurrentPage < 1 Then
            strTemp = strTemp & "下一页 尾页"
    Else
            strTemp = strTemp & "<a href='" & sfilename & "page=" & (CurrentPage + 1) & "'>下一页</a>&nbsp;"
            strTemp = strTemp & "<a href='" & sfilename & "page=" & n & "'>尾页</a>"
    End If
    strTemp = strTemp & "&nbsp;页次：<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>页 "
    strTemp = strTemp & "&nbsp;<b>" & MaxPerPage & "</b>" & "个文件/页"
    strTemp = strTemp & "&nbsp;转到：<select name='page' size='1' onchange='javascript:submit()'>"
    For i = 1 To n
        strTemp = strTemp & "<option value='" & i & "'"
        If CInt(CurrentPage) = CInt(i) Then strTemp = strTemp & " selected "
        strTemp = strTemp & ">第" & i & "页</option>"
    Next
    strTemp = strTemp & "</select>"
    strTemp = strTemp & "</td></tr></form></table>"
    showpage2 = strTemp
End Function


Function CutStr(Str)
    If Len(Str) > 18 Then
        CutStr = "..." & Right(Str, 18)
    Else
        CutStr = Str
    End If
End Function

'检查组件是否被支持
Function IsObjInstalled(strClassString)
'on error resume next
IsObjInstalled = False
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(strClassString)
If 0 = Err Then IsObjInstalled = True
Set xTestObj = Nothing
Err = 0
End Function
%>

