<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="CONN.asp"-->
<!--#include file="../common/Function1.asp"-->
<!--#include file="System_Function.asp"-->
<!--#include file="createThumb.asp"-->
<%Call adminIsLogin()%>
<%

Dim TruePath, theFolder, theSubFolder, theFile, thisfile, FileCount, TotalSize, TotalSize_Page
Dim TotalUnit, strTotalUnit, PageUnit, strPageUnit
Dim StrFileType, strFiles
Dim strDirName, tUploadDir
Dim strPath, strPath2, strPath3
Dim ItemIntro, UpFileType

'获取频道相关数据
tUploadDir = Trim(Request("UploadDir"))
Action=s("Action")
strFileName = "Admin_UploadFile.asp?ChannelID=" & ChannelID & "&UploadDir=" & UploadDir

Response.Write "<html><head><title>上传文件管理</title><meta http-equiv='Content-Type' content='text/html; charset=utf-8'><link href=""images/skin/Css_"&Session("Admin_Style_Num")&"/"&Session("Admin_Style_Num")&".css"" rel=""stylesheet"" type=""text/css""></head>"
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<table width='98%' border='0' align='center' cellpadding='3' cellspacing='1' Class='table'>" & vbCrLf
Response.Write "  <tr class='back'> "
Response.Write "    <td width='70' height='30' class='xingmu'><strong>管理导航：</strong></td>"
Response.Write "    <td height='30' class='xingmu'><a href='" & strFileName & "'>上传文件管理首页</a> | <a href='Admin_UploadFile_Clear.asp?ChannelID=" & ChannelID & "&UploadDir=" & UploadDir & "&Action=Clear'>清除无用文件</a></td>"
Response.Write "  </tr>"
Response.Write "</table>" & vbCrLf
If IsObjInstalled("Scripting.FileSystemObject") = False Then
    Response.Write "<b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能</font></b>"
    Response.End
else
	Set fso = Server.CreateObject("Scripting.FileSystemObject")	
End If

Select Case Action
Case "Clear"
    Call Clear
Case "DoClear"
    Call DoClear
Case Else
    Call Clear
End Select
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, strFileName)
End If
Response.Write "</body></html>"
Call CloseConn



Sub Clear()
    Response.Write "<br><table width='98%' border='0' cellspacing='1' cellpadding='3' class='table' align='center'>"
    Response.Write "  <tr class='back'>"
    Response.Write "    <td height='22' align='center' class='xingmu'><strong>清理无用的上传文件</strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='hback'>"
    Response.Write "    <td height='150'>"
    Response.Write "<form name='form1' method='post' action='Admin_UploadFile_Clear.asp'>"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;在添加内容时，经常会出现上传了图片后但却最终没有使用的情况，时间一久，就会产生大量无用垃圾文件。所以需要定期使用本功能进行清理。"
    Response.Write "<p>&nbsp;&nbsp;&nbsp;&nbsp;如果上传文件很多，或者信息数量较多，执行本操作需要耗费相当长的时间，请在访问量少时执行本操作。</p>"
    Response.Write "<p align='center'><input name='Action' type='hidden' id='Action' value='DoClear'><input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "<input name='UploadDir' type='hidden' value='" & tUploadDir & "'><input name='CurrentDir' type='hidden' value='" & CurrentDir & "'><input type='submit' name='Submit3' value=' 开始清理 '></p>"
    Response.Write "</form>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
End Sub


Sub DoClear()
    ParentDir = Replace(Replace(Replace(Trim(Request("ParentDir")), "../", ""), "..\", ""), "\", "/")
    If Left(ParentDir, 1) = "/" Then ParentDir = Right(ParentDir, Len(ParentDir) - 1)
    CurrentDir = Replace(Replace(Replace(Trim(Request("CurrentDir")), "/", ""), "\", ""), "..", "")
    
    Dim rs, sql
        sql = "select Pictures,content from newsinfo"
        Set rs = Conn.Execute(sql)
        Do While Not rs.EOF
            If rs(0) <> "" Then
                strFiles = strFiles & "|" & rs(0)
            End If
            If rs(1) <> "" Then
                ItemIntro = ItemIntro & "|" & rs(1)
            End If
            rs.MoveNext
        Loop
		rs.close
	sql = "select Pictures,BPictures,content from ProductInfo"
        Set rs = Conn.Execute(sql)
        Do While Not rs.EOF
            If rs(0) <> "" Then
                strFiles = strFiles & "|" & rs(0)
            End If
			If rs(1) <> "" Then
                strFiles = strFiles & "|" & rs(1)
            End If
            If rs(2) <> "" Then
                ItemIntro = ItemIntro & "|" & rs(2)
            End If
            rs.MoveNext
        Loop
		rs.close	
		sql = "select Pictures from ChannelInfo"
        Set rs = Conn.Execute(sql)
        Do While Not rs.EOF
            If rs(0) <> "" Then
                strFiles = strFiles & "|" & rs(0)
            End If
            rs.MoveNext
        Loop
		rs.close
		sql = "select TbSetting from T_Config"
        Set rs = Conn.Execute(sql)
        Do While Not rs.EOF
            If rs(0) <> "" Then
                strFiles = strFiles & "|" & rs(0)
            End If
            rs.MoveNext
        Loop
		rs.close
			
'正则表达式相关的变量
Dim regEx, Match, Match2, Matches, Matches2
Set regEx = New RegExp
regEx.IgnoreCase = True
regEx.Global = True
regEx.MultiLine = True	

        Dim tempStr, tempi, TempArray
		UpFileType = "gif|jpg|jpeg|jpe|bmp|png"
        regEx.Pattern = "<img.+?[^\>]>" '查询内容中所有 <img..>
        Set Matches = regEx.Execute(ItemIntro)
        For Each Match In Matches
            If tempStr <> "" Then
                tempStr = tempStr & "|" & Match.value '累计数组
            Else
                tempStr = Match.value
            End If
        Next
        If tempStr <> "" Then
            TempArray = Split(tempStr, "|") '分割数组
            tempStr = ""
            For tempi = 0 To UBound(TempArray)
                regEx.Pattern = "src\s*=\s*.+?\.(" & UpFileType & ")" '查询src =内的链接
                Set Matches = regEx.Execute(TempArray(tempi))
                For Each Match In Matches
                    If tempStr <> "" Then
                        tempStr = tempStr & "|" & Match.value '累加得到 链接加$Array$ 字符
                    Else
                        tempStr = Match.value
                    End If
                Next
            Next
        End If
        If tempStr <> "" Then
            regEx.Pattern = "src\s*=\s*" '过滤 src =
            tempStr = regEx.Replace(tempStr, "")
        End If
		
        strFiles = strFiles & tempStr

        UpFileType = "swf|rm|ra|ram"
        regEx.Pattern = "<param\s*name\s*=\s*""*src""*.+?[^\>]>" 
        Set Matches = regEx.Execute(ItemIntro)
        For Each Match In Matches
            If tempStr <> "" Then
                tempStr = tempStr & "|" & Match.value '累计数组
            Else
                tempStr = Match.value
            End If
        Next
        If tempStr <> "" Then
            TempArray = Split(tempStr, "|") '分割数组
            tempStr = ""
            For tempi = 0 To UBound(TempArray)
                regEx.Pattern = "value\s*=\s*.+?\.(" & UpFileType & ")" '查询value =内的链接
                Set Matches = regEx.Execute(TempArray(tempi))
                For Each Match In Matches
                    If tempStr <> "" Then
                        tempStr = tempStr & "|" & Match.value '累加得到 链接加$Array$ 字符
                    Else
                        tempStr = Match.value
                    End If
                Next
            Next
        End If
        If tempStr <> "" Then
            regEx.Pattern = "value\s*=\s*" '过滤 src =
            tempStr = regEx.Replace(tempStr, "")
        End If
		
        strFiles = strFiles & tempStr
		

        UpFileType = "mp3|avi|wmv|mpg|asf"
        regEx.Pattern = "<param\s*name\s*=\s*""*url""*.+?[^\>]>"
        Set Matches = regEx.Execute(ItemIntro)
        For Each Match In Matches
            If tempStr <> "" Then
                tempStr = tempStr & "|" & Match.value '累计数组
            Else
                tempStr = Match.value
            End If
        Next
        If tempStr <> "" Then
            TempArray = Split(tempStr, "|") '分割数组
            tempStr = ""
            For tempi = 0 To UBound(TempArray)
                regEx.Pattern = "value\s*=\s*.+?\.(" & UpFileType & ")" '查询value =内的链接
                Set Matches = regEx.Execute(TempArray(tempi))
                For Each Match In Matches
                    If tempStr <> "" Then
                        tempStr = tempStr & "|" & Match.value '累加得到 链接加$Array$ 字符
                    Else
                        tempStr = Match.value
                    End If
                Next
            Next
        End If
        If tempStr <> "" Then
            regEx.Pattern = "value\s*=\s*" '过滤 src =
            tempStr = regEx.Replace(tempStr, "")
        End If
		strFiles = strFiles & tempStr	
		

        UpFileType = "zip|rar|doc"
        regEx.Pattern = "<a.+?[^\>](rar\""*\s*|zip\""*\s*|doc\""*\s*)>" '查询内容中所有含zip，rar，doc名字的附件
        Set Matches = regEx.Execute(ItemIntro)
        For Each Match In Matches
            If tempStr <> "" Then
                tempStr = tempStr & "|" & Match.value '累计数组
            Else
                tempStr = Match.value
            End If
        Next
		
        If tempStr <> "" Then
            TempArray = Split(tempStr, "|") '分割数组
            tempStr = ""
            For tempi = 0 To UBound(TempArray)
                regEx.Pattern = "href\s*=\s*.+?\.(" & UpFileType & ")" '查询href =内的链接
                Set Matches = regEx.Execute(TempArray(tempi))
                For Each Match In Matches
                    If tempStr <> "" Then
                        tempStr = tempStr & "|" & Match.value '累加得到 链接加$Array$ 字符
                    Else
                        tempStr = Match.value
                    End If
                Next
            Next
        End If
        If tempStr <> "" Then
            regEx.Pattern = "href\s*=\s*" '过滤 href =
            tempStr = regEx.Replace(tempStr, "")
        End If
        strFiles = strFiles & tempStr					
		
    strFiles = LCase(strFiles)

UploadDir = "uploadimages"
RootDir = pic_path& UploadDir
    strPath = RootDir
    strPath2 = UploadDir
    strPath3 = ""
    If ParentDir <> "" Then
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

  i = 0
    If fso.FolderExists(Server.MapPath(RootDir)) = False Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>找不到文件夹！请上传文件后再进行管理！</li>"
        Exit Sub
    End If

    Set theFolder = fso.GetFolder(Server.MapPath(RootDir))
    For Each theFile In theFolder.Files
        If InStr(strFiles, LCase(theFile.name)) <= 0 Then
            theFile.Delete True
            i = i + 1
        End If
    Next
    For Each theSubFolder In theFolder.SubFolders
        For Each theFile In theSubFolder.Files
            If InStr(strFiles, LCase(theSubFolder.name & "/" & theFile.name)) <= 0 Then
                theFile.Delete True
                i = i + 1
            End If
        Next
    Next
    Call WriteSuccessMsg("清理无用文件成功！共删除了 <font color=red><b>" & i & "</b></font> 个无用的文件。", strFileName)
End Sub

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
