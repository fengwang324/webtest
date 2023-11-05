<!--#include file="Thumbs.asp"-->
<%
'****************************************************
' Email: xiaoyxh@163.com . QQ:438888382
'****************************************************
Class PublicCls
	   Private LocalCacheName,Cache_Data,CacheData,Reloadtime
	'检查是否是数字 ，并转换为长整型
	Public Function ChkClng(ByVal str)
	    'on error resume next
		If IsNumeric(str) Then
			ChkClng = CLng(str)
		Else
			ChkClng = 0
		End If
		If Err Then ChkClng=0
	End Function
	'取得Request.Querystring 或 Request.Form 的值
	Public Function G(Str)
	 G = Replace(Replace(Request(Str), "'", ""), """", "")
	End Function
	Function DelSql(Str)
		Dim SplitSqlStr,SplitSqlArr,I
		SplitSqlStr="dbcc|alter|drop|*|and |exec|or |insert|select|delete|update|count |master|truncate|declare|char|mid|chr|set |where|xp_cmdshell"
		SplitSqlArr = Split(SplitSqlStr,"|")
		For I=LBound(SplitSqlArr) To Ubound(SplitSqlArr)
			If Instr(LCase(Str),SplitSqlArr(I))>0 Then
				Response.Write "<script>alert('系统警告！\n\n1、您提交的数据有恶意字符" & SplitSqlArr(I) &";\n2、您的数据已经被记录;\n3、您的IP："&GetIP&";\n4、操作日期："&Now&";\n	');window.close();</script>"
				Response.End
			End if
		Next
		DelSql = Str
    End Function
	'取得Request.Querystring 或 Request.Form 的值
	Public Function S(Str)
	 S = DelSql(Replace(Replace(Request(Str), "'", ""), """", ""))
	End Function
	'读Cookies值
	Public Function C(Str)
	 C=DelSql(Request.Cookies(SiteSN)(Str))
	End Function
	Sub Echo(Str)
	  Response.Write Str
	End Sub
	'清除缓存,参数 MyCaheName-缓存名称
	Public Sub DelCahe(MyCaheName)
		Application.Lock
		Application.Contents.Remove(MyCaheName)
		Application.unLock
	End Sub
	'**************************************************
	'函数名：ListTitle
	'作  用：取标题
	'参  数：TitleStr 标题, TitleNum 取字符数
	'返回值：将标题分解成两行
	'**************************************************
	Public Function ListTitle(TitleStr, TitleNum)
		  Dim LeftStr, RightStr
			ListTitle = Trim(GotTopic(Trim(TitleStr), TitleNum))
			If Len(ListTitle) > CInt(TitleNum / 2) Then
			  LeftStr = GotTopic(ListTitle, CInt(TitleNum / 2))
			  RightStr = Mid(ListTitle, Len(LeftStr) + 1)
			  ListTitle = LeftStr & "<br>" & RightStr
			End If
	 End Function
	Function ListTitle1(TitleStr, TitleNum)
		   Dim ClsTitleStr, ClsTitleNum, I, J, ClsTempNum, k, ClsTitleStrResult, LeftStr, RightStr
			   ClsTitleNum = CInt(TitleNum)
			   ClsTempNum = Len(CStr(TitleStr))
			   If ClsTitleNum > ClsTempNum Then
				   ClsTitleNum = ClsTempNum
			   End If
			   ClsTitleStr = Left(CStr(TitleStr), ClsTitleNum)
			   Dim TempStr
			   For I = 1 To ClsTitleNum - 1
				   TempStr = TempStr & Mid(ClsTitleStr, I, 1) & "<br />"
			   Next
			   TempStr = TempStr & Right(ClsTitleStr, 1)
			   ListTitle1 = TempStr
	End Function
	'取得QueryString,或Form参数集合,参数NoCollect表示不收集的字段,多个用英文逗号隔开
	Function QueryParam(NoCollect)
		 Dim Param,R
		 For Each r In Request.QueryString
		  If FoundInArr(NoCollect,R,",")=false Then
			  If Request.QueryString(r)<>"" Then
				If Param="" Then
				 Param=r & "=" & Server.UrlEncode(Trim(Request.QueryString(r)))
				Else
				 Param=Param & "&" & r & "=" & Server.UrlEncode(Trim(Request.QueryString(r)))
				End If
			  End If
		 End If
		 Next
		' If Param<>"" Then QueryParam=Param:Exit Function
		 For Each r In Request.Form
		  If FoundInArr(NoCollect,R,",")=false Then
			  If Request.Form(r)<>"" Then
				If Param="" Then
				 Param=r & "=" & Server.UrlEncode(Trim(Request.Form(r)))
				Else
				 Param=Param & "&" & r & "=" & Server.UrlEncode(Trim(Request.Form(r)))
				End If
			  End If
		 End If
		 Next
		 QueryParam=Param
	End Function
	
	
	'**************************************************
	'函数名：Alert
	'作  用：弹出成功提示。
	'参  数：SuccessStr  ----成功提示信息
	'        Url   ------成功提示按下"确定"转向链接
	'返回值：无
	'**************************************************
	Public Function Alert(SuccessStr, Url)
	 If Url <> "" Then
	  Response.Write ("<script language=""Javascript""> alert('" & SuccessStr & "');location.href='" & Url & "';</script>")
	 Else
	  Response.Write ("<script language=""Javascript""> alert('" & SuccessStr & "');</script>")
	 End If
	End Function
	'**************************************************
	'函数名：AlertHistory
	'作  用：弹出警告消息后,停止所在页面的执行,返回n级。
	'参  数：SuccessStr  ----成功提示信息
	'        n   ------返回级数
	'返回值：无
	'**************************************************
	Public Function AlertHistory(SuccessStr, N)
		Response.Write ("<script language=""Javascript""> alert('" & SuccessStr & "');history.back(" & N & ");</script>")
		Response.End
	End Function
	'提示成功。并返回
	Sub AlertHintScript(SuccessStr)
	Response.Write "<script language=JavaScript>" & vbCrLf
	Response.Write "alert('" & SuccessStr & "');"
	Response.Write "location.replace('" & Request.ServerVariables("HTTP_REFERER") & "')" & vbCrLf
	Response.Write "</script>" & vbCrLf
	Response.End
	End Sub
	'**************************************************
	'函数名：Confirm
	'作  用：弹出成功提示。
	'参  数：SuccessStr  ----成功提示信息
	'        Url   ------成功提示按下"确定"转向链接
	'        Url1   ------confirm按下"取消"转向链接
	'返回值：无
	'**************************************************
	Public Function Confirm(SuccessStr, Url, Url1)
	 Response.Write ("<script language=""Javascript""> if (confirm('" & SuccessStr & "')){location.href='" & Url & "';}else{location.href='" & Url1 & "';}</script>")
	End Function
    
	Public Sub ShowTips(Action,Message)
		 Response.Redirect(pic_path & "Plus/error.asp?action="&action &"&message="&Server.URLEncode(message))
	End Sub
	'**************************************************
	'函数名：ShowError
	'作  用：显示错误信息。
	'参  数：Errmsg  ----出错信息
	'返回值：无
	'**************************************************
	Public Sub ShowError(Errmsg)
	    With Response
		.Write ("<br><br><div align=""center"">")
		.Write ("  <center>")
		.Write ("  <table border=""0"" cellpadding='2' cellspacing='1' class='border' width=""75%"" style=""MARGIN-TOP: 3px"" class='border'>")
		.Write ("	 <tr class=tdbg>")
		.Write ("			  <td width=""100%"" height=""30"" align=""center"">")
		.Write ("				<b> " & Errmsg & "&nbsp; </b>")
		.Write ("				</b>")
		.Write ("			  </td>")
		.Write ("	 </tr>")
		.Write ("	 <tr  class=tdbg>")
		.Write ("			  <td width=""100%"" height=""30"" align=""center"">")
		.Write ("				<p><b><a href=""javascript:history.go(-1)"">...::: 点 此 返 回 ")
		.Write ("				:::...</a></b>")
		.Write ("			  </td>")
		.Write ("			</tr>")		
		.Write ("  </table>")
		.Write ("  </center>")
		.Write ("</div>")
		.End
		End With
    end sub
	'**************************************************
	'函数名：GetIP
	'作  用：取得正确的IP
	'返回值：IP字符串
	'**************************************************
	Public Function GetIP() 
		Dim strIPAddr 
		If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" Or InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 Then 
			strIPAddr = Request.ServerVariables("REMOTE_ADDR") 
		ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") > 0 Then 
			strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",")-1) 
		ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") > 0 Then 
			strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";")-1)
		Else 
			strIPAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
		End If 
		getIP = Checkstr(Trim(Mid(strIPAddr, 1, 30)))
	End Function
	Public Function Checkstr(Str)
		If Isnull(Str) Then
			CheckStr = ""
			Exit Function 
		End If
		Str = Replace(Str,Chr(0),"")
		CheckStr = Replace(Str,"'","''")
	End Function
	Function IsNul(Str)
	  If Str="" Or IsNull(Str) Then IsNul=True Else IsNul=false
	End Function
	Function GetFileIcon(strFilename)
	strFilename=Split(LCase(strFilename),".")
	Select Case strFilename(Ubound(strFilename))
	Case "asp"
		GetFileIcon="asp.gif"
	Case "aspx"
		GetFileIcon="asp.gif"
	Case "jsp"
		GetFileIcon="asp.gif"
	Case "html"
		GetFileIcon="html.gif"
	Case "shtml"
		GetFileIcon="html.gif"
	Case "htm"
		GetFileIcon="html.gif"
	Case "shtm"
		GetFileIcon="html.gif"
	Case "xml"
		GetFileIcon="xml.gif"
	Case "bmp"
		GetFileIcon="bmp.gif"
	Case "doc"
		GetFileIcon="doc.gif"
	Case "gif"
		GetFileIcon="gif.gif"
	Case "ico"
		GetFileIcon="ico.gif"
	Case "jpg"
		GetFileIcon="jpg.gif"
	Case "js"
		GetFileIcon="js.gif"
	Case "mdb"
		GetFileIcon="mdb.gif"
	Case "mp3"
		GetFileIcon="media.gif"
	Case "wma"
		GetFileIcon="media.gif"
	Case "mid"
		GetFileIcon="media.gif"
	Case "png"
		GetFileIcon="png.gif"
	Case "rar"
		GetFileIcon="rar.gif"
	Case "swf"
		GetFileIcon="swf.gif"
	Case "txt"
		GetFileIcon="txt.gif"
	Case "xls"
		GetFileIcon="xls.gif"
	Case "xlt"
		GetFileIcon="xlt.gif"
	Case "zip"
		GetFileIcon="zip.gif"
	Case "exe"
		GetFileIcon="exe.gif"
	Case "dll"
		GetFileIcon="dll.gif"
	Case Else
		GetFileIcon="other.gif"
	End Select
End Function
	 Function InitialObject(str)
		'iis5创建对象方法Server.CreateObject(ObjectName);
		'iis6创建对象方法CreateObject(ObjectName);
		'默认为iis6，如果在iis5中使用，需要改为Server.CreateObject(str);
		Set InitialObject=CreateObject(str)
	 End Function

'匹配 img src,结果以|隔开 
	Function GetImgSrcArr(strng) 
	If strng="" Or IsNull(strng) Then GetImgSrcArr="":Exit Function
	Dim regEx,Match,Matches,values
	Set regEx = New RegExp
	regEx.Pattern = "src\=.+?\.(gif|jpg)"
	regEx.IgnoreCase = true 
	regEx.Global = True 
	Set Matches = regEx.Execute(strng)
	For Each Match in Matches
		if instr(lcase(Match.Value),"fileicon")=0 then
		 values=values&Match.Value&"|" 
		end if
	Next 
	GetImgSrcArr = Replace(Replace(Replace(Replace(values,"'",""),"""",""),"src=",""),Setting(2),"")
	If GetImgSrcArr<>"" Then GetImgSrcArr = left(GetImgSrcArr,len(GetImgSrcArr)-1)
	End Function
Public Function CreateListFolder(ByVal Folder)
		Dim FSO, WaitCreateFolder, SplitFolder, CF, k
		 'on error resume next
		If Folder = "" Then
		 CreateListFolder = False:Exit Function
		End If
	   Folder = Replace(Folder, "\", "/")
	   If Right(Folder, 1) <> "/" Then Folder = Folder & "/"
	   If Left(Folder, 1) <> "/" Then Folder = "/" & Folder

		 Set FSO = CreateObject("Scripting.FileSystemObject")
		 If Not FSO.FolderExists(Server.MapPath(Folder)) Then
		   SplitFolder = Split(Folder, "/")
		 For k = 0 To UBound(SplitFolder) - 1
		  If k = 0 Then
		   CF = SplitFolder(k) & "/"
		  Else
		   CF = CF & SplitFolder(k) & "/"
		  End If
		  If (Not FSO.FolderExists(Server.MapPath(CF))) Then
			 FSO.CreateFolder (Server.MapPath(CF))
			 CreateListFolder = True
		  End If
		 Next
	   End If
	   Set FSO = Nothing
	   If Err.Number <> 0 Then
	   Err.Clear
	   CreateListFolder = False
	   Else
	   CreateListFolder = True
	   End If
	 End Function
	
	 '**************************************************
	'函数名：DeleteFolder
	'作  用：删除指定目录
	'参  数：FolderStr要删除的目录
	'返回值：成功返回true 否则返回Flase
	'**************************************************
	Public Function DeleteFolder(FolderStr)
	   Dim FSO
	   'on error resume next
	   FolderStr = Replace(FolderStr, "\", "/")
	   Set FSO = CreateObject("Scripting.FileSystemObject")
		If FSO.FolderExists(Server.MapPath(FolderStr)) Then
			  FSO.DeleteFolder (Server.MapPath(FolderStr))
		Else
		DeleteFolder = True
		End If
	   Set FSO = Nothing
	   If Err.Number <> 0 Then
	   Err.Clear:DeleteFolder = False
	   Else
	   DeleteFolder = True
	   End If
	End Function
	
	Function FilterIDs(byval strIDs)
	Dim arrIDs,i,strReturn
	strIDs=Trim(strIDs)
	If Len(strIDs)=0  Then Exit Function
	arrIDs=Split(strIDs,",")
	For i=0 To Ubound(arrIds)
		If ChkClng(Trim(arrIDs(i)))<>0 Then
			strReturn=strReturn & "," & Int(arrIDs(i))
		End If
	Next
	If Left(strReturn,1)="," Then strReturn=Right(strReturn,Len(strReturn)-1)
	FilterIDs=strReturn
	End Function
	 '**************************************************
	'函数名：DeleteFile
	'作  用：删除指定文件
	'参  数：FileStr要删除的文件
	'返回值：成功返回true 否则返回Flase
	'**************************************************
	Public Function DeleteFile(FileStr)
	   Dim FSO
	   'on error resume next
	   Set FSO = CreateObject("Scripting.FileSystemObject")
		If FSO.FileExists(Server.MapPath(FileStr)) Then
			FSO.DeleteFile Server.MapPath(FileStr), True
		Else
		DeleteFile = True
		End If
	   Set FSO = Nothing
	   If Err.Number <> 0 Then
	   Err.Clear:DeleteFile = False
	   Else
	   DeleteFile = True
	   End If
	End Function
	'**********************************************************************
	'函数名：CheckFileShowOrNot
	'参数：AllowShowExtNameStr允许的文件扩展名，ExtName实际文件扩展名
	'**********************************************************************
	Public Function CheckFileShowOrNot(AllowShowExtNameStr, ExtName)
		If ExtName = "" Then
			CheckFileShowOrNot = False
		Else
			If InStr(1, AllowShowExtNameStr, ExtName) = 0 Then
				CheckFileShowOrNot = False
			Else
				CheckFileShowOrNot = True
			End If
		End If
	End Function
	'**********************************************************************
	'函数名：GetFieSize
	'作用：取得指定文件的大小
	'参数：FilePath--文件位置
	'**********************************************************************
	Public Function GetFieSize(FilePath)
			GetFieSize = 0
			Dim FSO, F
			'on error resume next
			Set FSO = Server.CreateObject("Scripting.FileSystemObject")
			Set F = FSO.GetFile(FilePath)
			GetFieSize = F.size
			Set F = Nothing:Set FSO = Nothing
	End Function
    '取得目录大小
	Public Function GetFolderSize(FolderPath)
		dim fso:Set FSO = Server.CreateObject("Scripting.FileSystemObject")
		if fso.FolderExists(Server.MapPath(FolderPath)) then
		dim userfilespace:set UserFileSpace=FSO.GetFolder(Server.MapPath(FolderPath))
        GetFolderSize=UserFileSpace.size
		else
		 GetFolderSize=0:exit function
		end if
		set userfilespace=nothing:set fso=nothing
	End Function
	'*************************************************************************************
	'文件备份过程
	'过程名：backupdata
	'参数：CurrPath原文件完整物理地址，BackPath目标备份文件完整物理地址
	'*************************************************************************************
	
	Public Function BackUpData(CurrPath, BackPath)
	  'on error resume next
	  Dim FSO:Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	 FSO.copyfile CurrPath, BackPath
	 If Err Then
	   BackUpData = False
	 Else
	   BackUpData = True
	 End If
	  FSO.Close:Set FSO = Nothing
	End Function
	'------------------检查某一目录是否存在-------------------
	Public Function CheckDir(FolderPath)
	Dim fso1
	FolderPath = Server.MapPath(".") & "\" & FolderPath
	Set fso1 = CreateObject("Scripting.FileSystemObject")
	If fso1.FolderExists(FolderPath) Then
	CheckDir = True
	Else
	CheckDir = False
	End If
	Set fso1 = Nothing
	End Function
	'------------------检查某一文件是否存在-------------------
	Public Function CheckFile(FileName)
		 'on error resume next
		 Dim FsoObj
		 Set FsoObj = Server.CreateObject("Scripting.FileSystemObject")
		  If Not FsoObj.FileExists(Server.MapPath(FileName)) Then
			  CheckFile = False
			  Exit Function
		  End If
		 CheckFile = True:Set FsoObj = Nothing
	End Function
	
	
	'**************************************************
	'函数名：WriteTOFile
	'作  用：写内容到指定的html文件
	'参  数：Filename  ----目标文件件 如 mb\index.htm
	'        Content   ------要写入目标文件的内容
	'返回值：成功返回true ,失败返回false
	'**************************************************
	Public Function WriteTOFile(FileName, Content)
	    'on error resume next
		dim stm:set stm=server.CreateObject("adodb.stream")
		stm.Type=2 '以文本模式读取
		stm.mode=3
		stm.charset="utf-8"
		stm.open
		stm.WriteText content
		stm.SaveToFile server.MapPath(FileName),2 
		stm.flush
		stm.Close
		set stm=nothing
	  
	   If Err.Number <> 0 Then
		 WriteTOFile = False
	   Else
		 WriteTOFile = True
	   End If
	End Function

'根据栏目类型选择数据表名
public function SQLName(tyle)
	select case chkclng(tyle)
		case 4
			SQLName="CB_Article"
		case 5
			SQLName="CB_Products"
		case else
			SQLName="CB_Article"
	end select
end function

'**************************************************
	'函数名：MakeRandom
	'作  用：生成指定位数的随机数
	'参  数： maxLen  ----生成位数
	'返回值：成功:返回随机数
	'**************************************************
	Public Function MakeRandom(ByVal maxLen)
	  Dim strNewPass,whatsNext, upper, lower, intCounter
	  Randomize
	 For intCounter = 1 To maxLen
	   upper = 57:lower = 48:strNewPass = strNewPass & Chr(Int((upper - lower + 1) * Rnd + lower))
	 Next
	   MakeRandom = strNewPass
	End Function
'**************************************************
	'函数名：ReturnChannelAllowUpFilesSize
	'作  用：返回频道的最大允许上传文件大小
	'参  数：ChannelID--频道ID
	'**************************************************
	Public Function ReturnChannelAllowUpFilesSize(ChannelID)
	   ChannelID = ChkClng(ChannelID)
	   Dim CRS:Set CRS=conn.execute("Select top 1 UpFilesSize From Channelinfo Where id=" & ChannelID)
	  If CInt(ChannelID) = 0 Or (CRS.EOF And CRS.BOF) Then
		ReturnChannelAllowUpFilesSize = Setting(6)
	  Else
		ReturnChannelAllowUpFilesSize = CRS(0)
	  End If
	CRS.Close:Set CRS = Nothing
	End Function
	'**************************************************
	'函数名：ReturnChannelAllowUpFilesType
	'作  用：返回频道的允许上传的文件类型
	'参  数：ChannelID--频道ID,TypeFlag 0-取全部 1-图片类型 2-flash 类型 3-Windows 媒体类型 4-Real 类型 5-其它类型
	'**************************************************
	Public Function ReturnChannelAllowUpFilesType(ChannelID, TypeFlag)
	  If ChkClng(ChannelID) = 0 Then  ReturnChannelAllowUpFilesType = Setting(7):Exit Function
	  If Not IsNumeric(TypeFlag) Then TypeFlag = 0
		If TypeFlag = 0 Then   '所有允许的类型
		 ReturnChannelAllowUpFilesType = C_S(ChannelID,28) & "|" & C_S(ChannelID,29) & "|" & C_S(ChannelID,30) & "|" & C_S(ChannelID,31) & "|" & C_S(ChannelID,32)
		Else
		 ReturnChannelAllowUpFilesType = C_S(ChannelID,27+TypeFlag)
		End If
	End Function
	'取得Tag之间的循环体
	Function GetTagLoop(ByVal Content)
			Dim regEx, Matches, Match, LoopStr
			Set regEx = New RegExp
			regEx.Pattern = "{Tag([\s\S]*?):(.+?)}"
			regEx.IgnoreCase = True
			regEx.Global = True
			Set Matches = regEx.Execute(Content)
			For Each Match In Matches
				Content=Replace(Content,Match.Value,"")
				Content=Replace(Content,"{/Tag}","")
			Next
			GetTagLoop=Content
    End Function
		'函数名：GetDomain
	'作  用：获取URL,包括虚拟目录 如http://www.kesion.com/ 或 http://www.kesion.com/Sys/  其中 Sys/为虚拟目录
	'参  数：  无
	'返回值：完整域名
	'***************************************************************************************************************
	'***********************************************************************************************************
	'函数名：ReturnLabelInfo
	'参  数：LabelName ----  默认标签名称,FolderID---标签目录ID号,Descript---标签描述
	'返回值：标签基本信息
	'*************************************************************************************************************
	Public Function ReturnLabelInfo(LabelName, FolderID, Descript)
	  ReturnLabelInfo = ReturnLabelInfo & ("        <table width=""98%"" border='0' align='center' cellpadding='3' cellspacing='1' class='table' style='margin-top:6px'>")
	  ReturnLabelInfo = ReturnLabelInfo & ("          <tr  height=""26"" class=back><td colspan=2 align=center class=xingmu><strong>")
	  If g("labelid")="" Then
	  ReturnLabelInfo = ReturnLabelInfo & ("创 建 新 标 签")
	  Else
	  ReturnLabelInfo = ReturnLabelInfo & (" 修 改 标 签 属 性")
	  End If
	  ReturnLabelInfo = ReturnLabelInfo & ("</strong></td>")
	  ReturnLabelInfo = ReturnLabelInfo & ("          </tr>")
	  ReturnLabelInfo = ReturnLabelInfo & ("          <tr class=hback>")
	  ReturnLabelInfo = ReturnLabelInfo & ("      <td  colspan=2 height=""30"">标签名称")
	  ReturnLabelInfo = ReturnLabelInfo & ("        <input name=""LabelName"" size='35' class=""textbox"" type=""text"" id=""LabelName"" value=""" & LabelName & """>")
	  ReturnLabelInfo = ReturnLabelInfo & ("        <font color=""#FF0000""> * 调用格式""{LB_标签名称}""</font></td>")
	  ReturnLabelInfo = ReturnLabelInfo & ("    </tr>")
	  ReturnLabelInfo = ReturnLabelInfo & ("    <tr class=hback>")
	  ReturnLabelInfo = ReturnLabelInfo & ("      <td  colspan=2 height=""30"">标签目录 " & ReturnLabelFolderTree(FolderID, 0) & "<font color=""#FF0000""> 请选择标签归属目录，以便日后管理标签</font></td>")
	  ReturnLabelInfo = ReturnLabelInfo & ("    </tr>")
	  ReturnLabelInfo = ReturnLabelInfo & ("    <tr class=hback style='display:none'>")
	  ReturnLabelInfo = ReturnLabelInfo & ("      <td  colspan=2 height=""30"">标签描述")
	  ReturnLabelInfo = ReturnLabelInfo & ("        <input name=""Descript"" class=""textbox"" type=""text"" id=""Descript"" value=""" & Descript & """ size=""40"">")
	  ReturnLabelInfo = ReturnLabelInfo & ("        <font color=""#FF0000""> 请在此输入标签的说明,方便以后查找</font></td>")
	  ReturnLabelInfo = ReturnLabelInfo & ("    </tr>")
	 ' ReturnLabelInfo = ReturnLabelInfo & ("    </table>")
	End Function
	
	'模型选项
	Sub LoadChannelOption(ChannelID)
		If not IsObject(Application(SiteSN&"_ChannelConfig")) Then LoadChannelConfig
		Dim ModelXML,Node
		Set ModelXML=Application(SiteSN&"_ChannelConfig")
		For Each Node In ModelXML.documentElement.SelectNodes("channel")
		 if Node.SelectSingleNode("@ks21").text="1" and Node.SelectSingleNode("@ks0").text<>"6" and Node.SelectSingleNode("@ks0").text<>"9" and Node.SelectSingleNode("@ks0").text<>"10" Then
		  If Trim(ChannelID)=Trim(Node.SelectSingleNode("@ks0").text) Then
		  echo "<option value='" &Node.SelectSingleNode("@ks0").text &"' selected>" & Node.SelectSingleNode("@ks1").text & "</option>"
		  Else
		  echo "<option value='" &Node.SelectSingleNode("@ks0").text &"'>" & Node.SelectSingleNode("@ks1").text & "</option>"
		  End If
		 End If
		next

	End Sub
	
	Public Function ArrayToxml(DataArray,Recordset,row,xmlroot)
				Dim i,node,rs,j
				If xmlroot="" Then xmlroot="xml"
				Set ArrayToxml = Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
				ArrayToxml.appendChild(ArrayToxml.createElement(xmlroot))
				If row="" Then row="row"
				For i=0 To UBound(DataArray,2)
					Set Node=ArrayToxml.createNode(1,row,"")
					j=0
					For Each rs in Recordset.Fields
							 node.attributes.setNamedItem(ArrayToxml.createNode(2,LCase(rs.name),"")).text= DataArray(j,i)& ""
							 j=j+1
					Next
					ArrayToxml.documentElement.appendChild(Node)
				Next
		End Function
'**************************************************
	'函数名：R
	'作  用：过滤非法的SQL字符
	'参  数：strChar-----要过滤的字符
	'返回值：过滤后的字符
	'**************************************************
	Public Function R(strChar)
		If strChar = "" Or IsNull(strChar) Then R = "":Exit Function
		Dim strBadChar, arrBadChar, tempChar, I
		'strBadChar = "$,#,',%,^,&,?,(,),<,>,[,],{,},/,\,;,:," & Chr(34) & "," & Chr(0) & ""
		strBadChar = "+,',--,%,^,&,?,(,),<,>,[,],{,},/,\,;,:," & Chr(34) & "," & Chr(0) & ""
		arrBadChar = Split(strBadChar, ",")
		tempChar = strChar
		For I = 0 To UBound(arrBadChar)
			tempChar = Replace(tempChar, arrBadChar(I), "")
		Next
		tempChar = Replace(tempChar, "@@", "@")
		R = tempChar
	End Function		
'----------------------------------------------------------------------------------------------------------------------------
		'函数名:DateFormat
		'功 能:日期格式函数
		'参 数: DateStr日期, Types转换类型		'----------------------------------------------------------------------------------------------------------------------------
		Function DateFormat(DateStr, Types)
			Dim DateString
			If IsDate(DateStr) = False Then
				DateFormat = "":Exit Function
			End If
			Select Case CStr(Types)
			  Case "0"
				DateFormat = ""
				Exit Function
			  Case 1,21,41
			      DateString=Year(DateStr) & "-" & Right("0" & Month(DateStr), 2) & "-" & Right("0" & Day(DateStr), 2)
			      if Types=21 then
				   DateString = "(" & DateString &")"
				  elseIf Types=41 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 2,22,42
			      DateString=Year(DateStr) & "." & Right("0" & Month(DateStr), 2) & "." & Right("0" & Day(DateStr), 2)
			      if Types=22 then
				   DateString = "(" & DateString &")"
				  elseIf Types=42 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 3,23,43
			      DateString=Year(DateStr) & "/" & Right("0" & Month(DateStr), 2) & "/" & Right("0" & Day(DateStr), 2)
			      if Types=23 then
				   DateString = "(" & DateString &")"
				  elseIf Types=43 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 4,24,44
			      DateString=Right("0" & Month(DateStr), 2) & "/" & Right("0" & Day(DateStr), 2) & "/" & Year(DateStr)
			      if Types=24 then
				   DateString = "(" & DateString &")"
				  elseIf Types=44 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 5,25,45
				  DateString = Year(DateStr) & "年" & Right("0" & Month(DateStr), 2) & "月"
			      if Types=25 then
				   DateString = "(" & DateString &")"
				  elseIf Types=45 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 6,26,46
				  DateString = Year(DateStr) & "年" & Right("0" & Month(DateStr), 2) & "月" & Right("0" & Day(DateStr), 2) & "日"
			      if Types=26 then
				   DateString = "(" & DateString &")"
				  elseIf Types=46 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 7,27,47
				  DateString = Right("0" & Month(DateStr), 2) & "." & Right("0" & Day(DateStr), 2) & "." & Year(DateStr)
			      if Types=27 then
				   DateString = "(" & DateString &")"
				  elseIf Types=47 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 8,28,48
				  DateString = Right("0" & Month(DateStr), 2) & "-" & Right("0" & Day(DateStr), 2) & "-" & Year(DateStr)
				  if Types=28 then
				   DateString = "(" & DateString &")"
				  elseIf Types=48 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 9,29,49
				  DateString = Right("0" & Month(DateStr), 2) & "/" & Right("0" & Day(DateStr), 2)
				  if Types=29 then
				   DateString = "(" & DateString &")"
				  elseIf Types=49 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 10,30,50
				  DateString = Right("0" & Month(DateStr), 2) & "." & Right("0" & Day(DateStr), 2)
			      if Types=30 then
				   DateString = "(" & DateString &")"
				  elseIf Types=50 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 11,31,51
				  DateString = Right("0" & Month(DateStr), 2) & "月" & Right("0" & Day(DateStr), 2) & "日"
			      if Types=31 then
				   DateString = "(" & DateString &")"
				  elseIf Types=51 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 12,32,52
				  DateString = Right("0" & Day(DateStr), 2) & "日" & Right("0" & Hour(DateStr), 2) & "时"
				  if Types=32 then
				   DateString = "(" & DateString &")"
				  elseIf Types=52 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 13,33,53
				  DateString = Right("0" & Day(DateStr), 2) & "日" & Right("0" & Hour(DateStr), 2) & "点"
			      if Types=33 then
				   DateString = "(" & DateString &")"
				  elseIf Types=53 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 14,34,54
				  DateString = Right("0" & Hour(DateStr), 2) & "时" & Minute(DateStr) & "分"
				  if Types=34 then
				   DateString = "(" & DateString &")"
				  elseIf Types=54 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 15,35,55
				  DateString = Right("0" & Hour(DateStr), 2) & ":" & Right("0" & Minute(DateStr), 2)
			      if Types=35 then
				   DateString = "(" & DateString &")"
				  elseIf Types=55 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 16,36,56
				  DateString = Right("0" & Month(DateStr), 2) & "-" & Right("0" & Day(DateStr), 2)
				 if Types=36 then
				   DateString = "(" & DateString &")"
				  elseIf Types=56 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 17,37,57
				  DateString = Right("0" & Month(DateStr), 2) & "/" & Right("0" & Day(DateStr), 2) &" " &Right("0" & Hour(DateStr), 2)&":"&Right("0" & Minute(DateStr), 2)
				  if Types=37 then
				   DateString = "(" & DateString &")"
				  elseIf Types=57 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case Else
				  DateString = DateStr
			 End Select
			 DateFormat = DateString
		 End Function

	'**************************************************
	'函数名：ReturnDateFormat
	'作  用：返回系统支持的日期格式
	'参  数：SelectDate 预定选中的日期格式
	'**************************************************
	Public Function ReturnDateFormat(SelectDate)
			 Dim TempFormatDateStr, Str
			 If CStr(SelectDate) = "0" Then
				TempFormatDateStr = ("<option value=""0"" Selected>-不显示日期-</option> ")
			  Else
				TempFormatDateStr = ("<option value=""0"">-不显示日期-</option> ")
			  End If
			  If CStr(SelectDate) = "1" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""1""" & Str & " >2005-10-1</option>")
			  If CStr(SelectDate) = "2" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""2""" & Str & ">2005.10.1</option>")
			  If CStr(SelectDate) = "3" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""3""" & Str & ">2005/10/1</option>")
			  If CStr(SelectDate) = "4" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""4""" & Str & ">10/1/2005</option>")
			  If CStr(SelectDate) = "5" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""5""" & Str & ">2005年10月</option>")
			  If CStr(SelectDate) = "6" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""6""" & Str & ">2005年10月1日</option>")
			  If CStr(SelectDate) = "7" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""7""" & Str & ">10.1.2005</option>")
			  If CStr(SelectDate) = "8" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""8""" & Str & ">10-1-2005</option>")
			  If CStr(SelectDate) = "9" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""9""" & Str & ">10/1</option>")
			  If CStr(SelectDate) = "10" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""10""" & Str & ">10.1</option>")
			  If CStr(SelectDate) = "11" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""11""" & Str & ">10月1日</option>")
			  If CStr(SelectDate) = "12" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""12""" & Str & ">1日12时</option>")
			  If CStr(SelectDate) = "13" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""13""" & Str & ">1日12点</option>")
			  If CStr(SelectDate) = "14" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""14""" & Str & ">12时12分</option>")
			  If CStr(SelectDate) = "15" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""15""" & Str & ">12:12</option>")
			  If CStr(SelectDate) = "16" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""16""" & Str & ">10-1</option>")
			   If CStr(SelectDate) = "17" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""17""" & Str & ">10/1 12:00</option>")
			  
			  TempFormatDateStr = TempFormatDateStr & ("<optgroup  label=""-----加括号格式-----""></optgroup>")

			  If CStr(SelectDate) = "21" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""21""" & Str & " >(2005-10-1)</option>") 
			  If CStr(SelectDate) = "22" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""22""" & Str & ">(2005.10.1)</option>")
			  If CStr(SelectDate) = "23" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""23""" & Str & ">(2005/10/1)</option>")
			  If CStr(SelectDate) = "24" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""24""" & Str & ">(10/1/2005)</option>")
			  If CStr(SelectDate) = "25" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""225""" & Str & ">(2005年10月)</option>")
			  If CStr(SelectDate) = "26" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""26""" & Str & ">(2005年10月1日)</option>")
			  If CStr(SelectDate) = "27" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""27""" & Str & ">(10.1.2005)</option>")
			  If CStr(SelectDate) = "28" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""28""" & Str & ">(10-1-2005)</option>")
			  If CStr(SelectDate) = "29" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""29""" & Str & ">(10/1)</option>")
			  If CStr(SelectDate) = "30" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""30""" & Str & ">(10.1)</option>")
			  If CStr(SelectDate) = "31" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""31""" & Str & ">(10月1日)</option>")
			  If CStr(SelectDate) = "32" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""32""" & Str & ">(1日12时)</option>")
			  If CStr(SelectDate) = "33" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""33""" & Str & ">(1日12点)</option>")
			  If CStr(SelectDate) = "34" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""34""" & Str & ">(12时12分)</option>")
			  If CStr(SelectDate) = "35" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""35""" & Str & ">(12:12)</option>")
			  If CStr(SelectDate) = "36" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""36""" & Str & ">(10-1)</option>")
			  If CStr(SelectDate) = "37" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""37""" & Str & ">(10/1 12:00)</option>")


			  TempFormatDateStr = TempFormatDateStr & ("<optgroup  label=""-----加中括号格式-----""></optgroup>")
			  If CStr(SelectDate) = "41" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""41""" & Str & ">[2005.10.1]</option>")
			  If CStr(SelectDate) = "42" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""42""" & Str & ">[2005.10.1]</option>")
			  If CStr(SelectDate) = "43" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""43""" & Str & ">[2005/10/1]</option>")
			  If CStr(SelectDate) = "44" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""44""" & Str & ">[10/1/2005]</option>")
			  If CStr(SelectDate) = "45" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""45""" & Str & ">[2005年10月]</option>")
			  If CStr(SelectDate) = "46" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""46""" & Str & ">[2005年10月1日]</option>")
			  If CStr(SelectDate) = "47" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""47""" & Str & ">[10.1.2005]</option>")
			  If CStr(SelectDate) = "48" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""48""" & Str & ">[10-1-2005]</option>")
			  If CStr(SelectDate) = "49" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""49""" & Str & ">[10/1]</option>")
			  If CStr(SelectDate) = "50" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""50""" & Str & ">[10.1]</option>")
			  If CStr(SelectDate) = "51" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""51""" & Str & ">[10月1日]</option>")
			  If CStr(SelectDate) = "52" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""52""" & Str & ">[1日12时]</option>")
			  If CStr(SelectDate) = "53" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""53""" & Str & ">[1日12点]</option>")
			  If CStr(SelectDate) = "54" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""54""" & Str & ">[12时12分]</option>")
			  If CStr(SelectDate) = "55" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""55""" & Str & ">[12:12]</option>")
			  If CStr(SelectDate) = "56" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""56""" & Str & ">[10-1]</option>")
			  If CStr(SelectDate) = "57" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""57""" & Str & ">[10/1 12:00]</option>")
			ReturnDateFormat = TempFormatDateStr
	End Function
	'***********************************************************************************************************
	'函数名：ReturnLabelFolderTree
	'作  用：显示标签目录列表。
	'参  数：SelectID ----  默认目录树ID号,ChannelID频道ID号,FolderType目录类型 0系统函数标签目录,1自由标签目录
	'返回值：标签目录列表
	'*************************************************************************************************************
	Public Function ReturnLabelFolderTree(SelectID, FolderType)
		   Dim TempStr,ID,FolderName,TempSql
		   SelectID = Trim(SelectID)
		   If FolderType = "" Then FolderType = 0
		   TempStr = "<select class='textbox' style='width:200;border-style: solid; border-width: 1' name='ParentID'>"
		   
		   TempStr = TempStr & "<option value='0' Selected>根目录</option>"
		   
			Dim TempRS:Set TempRS=server.CreateObject("adodb.recordset")
			TempSql="Select ID,FolderName from CB_LabelFolder Where FolderType=" & FolderType & " And ParentID='0' Order By AddDate desc"
			TempRS.open TempSql,conn,1,1
			
			Do While Not TempRS.EOF
			   ID = Trim(TempRS(0))
			   FolderName = Trim(TempRS(1))
			   TempStr = TempStr & "<option  "
			   If SelectID = ID Then TempStr = TempStr & " Selected"
			   TempStr = TempStr & " value='" & ID & "'>" & FolderName & " </option>"
			   TempStr = TempStr & ReturnSubLabelFolderTree(ID, SelectID)
			TempRS.MoveNext
			Loop
			TempRS.Close:Set TempRS = Nothing
			TempStr = TempStr & "</select>"
			ReturnLabelFolderTree = TempStr
	End Function
	
	'************************************************************************************
	'函数名：ReturnSubLabelFolderTree
	'作  用：查找并返子树数据。
	'参  数：ParentID ----父节点ID,   FolderID ----选择项ID
	'返回值：标签目录子树列表
	'************************************************************************************
	Public Function ReturnSubLabelFolderTree(ParentID, FolderID)
	  Dim SubTypeList, SubRS, SpaceStr, k, Total, Num,FolderName, ID,TJ
	  
	  Set SubRS = Server.CreateObject("ADODB.RECORDSET")
	  SubRS.Open ("Select count(ID) AS total from CB_LabelFolder Where ParentID='" & ParentID & "'"), Conn, 1, 1
	  Total = SubRS("Total")
	  SubRS.Close
	  SubRS.Open ("Select ID,FolderName,TS from CB_LabelFolder Where ParentID='" & ParentID & "' Order BY AddDate Desc"), Conn, 1, 1
	  Num = 0
	  Do While Not SubRS.EOF
	   Num = Num + 1:SpaceStr = ""
		TJ = UBound(Split(SubRS(2), ","))
		For k = 1 To TJ - 1
		  If k = 1 And k <> TJ - 1 Then
		  SpaceStr = SpaceStr & "&nbsp;&nbsp;│"
		  ElseIf k = TJ - 1 Then
			If Num = Total Then
				 SpaceStr = SpaceStr & "&nbsp;&nbsp;└ "
			Else
				 SpaceStr = SpaceStr & "&nbsp;&nbsp;├ "
			End If
		  Else
		   SpaceStr = SpaceStr & "&nbsp;&nbsp;│"
		  End If
		Next
	  ID = Trim(SubRS(0))
	  FolderName = Trim(SubRS(1))
	  If FolderID = ID Then
	   SubTypeList = SubTypeList & "<option selected value='" & ID & "'>" & SpaceStr & FolderName & "</option>"
	  Else
	   SubTypeList = SubTypeList & "<option value='" & ID & "'>" & SpaceStr & FolderName & "</option>"
	  End If
	   SubTypeList = SubTypeList & ReturnSubLabelFolderTree(ID, FolderID)
	  SubRS.MoveNext
	 Loop
	  SubRS.Close:Set SubRS = Nothing:ReturnSubLabelFolderTree = SubTypeList
	End Function
	'**************************************************
	'函数名：ReturnOpenTypeStr
	'作  用：返回系统支持的打开窗口方式(带可输入的下拉框)
	'参  数：SelectValue 预定选中的链接目标
	'**************************************************
	Public Function ReturnOpenTypeStr(SelectValue)
	  ReturnOpenTypeStr = "链接目标 <select onchange=""document.getElementById('OpenType').value=this.value;"" name='sOpenType'><option value=''>-没有设置-</option>"
	  ReturnOpenTypeStr = ReturnOpenTypeStr & "<option value=""_blank""> 新窗口(_blank) </option>"
	  ReturnOpenTypeStr = ReturnOpenTypeStr & "<option value=""_parent""> 父窗口(_parent) </option>"
	  ReturnOpenTypeStr = ReturnOpenTypeStr & "<option value=""_self""> 本窗口(_self) </option>"
	  ReturnOpenTypeStr = ReturnOpenTypeStr & "<option value=""_top""> 整页(_top) </option>"
	  ReturnOpenTypeStr = ReturnOpenTypeStr & "</select>=>"
	  ReturnOpenTypeStr = ReturnOpenTypeStr & "<input type='text' name='OpenType' id='OpenType' size='10' value='" & SelectValue &"'>"
	  Exit Function
	End Function
	'分页样式
	Public Function ReturnPageStyle(PageStyle)
		ReturnPageStyle = "         分页样式"
		ReturnPageStyle = ReturnPageStyle & "         <select name=""PageStyle"" id=""PageStyle"" style=""width:70%;"" class=""textbox"">"
		ReturnPageStyle = ReturnPageStyle & "          <option value=1"
		If PageStyle="1" Then ReturnPageStyle = ReturnPageStyle & " Selected"
		ReturnPageStyle = ReturnPageStyle & ">①首页 上一页 下一页 尾页</option>"
		ReturnPageStyle = ReturnPageStyle & "          <option value=2"
		If PageStyle="2" Then ReturnPageStyle = ReturnPageStyle & " Selected"
		ReturnPageStyle = ReturnPageStyle & ">②共N页/N篇 [1] [2] [3]</option>"
		ReturnPageStyle = ReturnPageStyle & "          <option value=3"
		If PageStyle="3" Then ReturnPageStyle = ReturnPageStyle & " Selected"
		ReturnPageStyle = ReturnPageStyle & ">③<< <  > >></option>"
		ReturnPageStyle = ReturnPageStyle & "          <option value=4"
		If PageStyle="4" Then ReturnPageStyle = ReturnPageStyle & " Selected"
		ReturnPageStyle = ReturnPageStyle & " style='color:blue'>④数字导航样式</option>"
		ReturnPageStyle = ReturnPageStyle & "         </select>"
	End Function
	Public Function GetItemURL(ByVal ChannelID,ByVal Tid,ByVal InfoID,ByVal Fname)
		  IF Not Isnumeric(ChannelID) Then GetItemURL="#":Exit Function
		  If  getChannelInfo(Tid,"FsoHtmlTF")=0 Then 
		        if getChannelInfo(Tid,"FsoHtmlTF")=0 Then
				 GetItemURL=GetDomain & "Item/Show.asp?m=" & ChannelID & "&d=" &InfoID
				ElseIf getChannelInfo(Tid,"FsoHtmlTF")=2 Then
				 GetItemURL=GetDomain & GCls.StaticPreContent & "-" & InfoID & "-"& ChannelID & GCls.StaticExtension
				Else
				 GetItemURL=GetDomain & "?" & GCls.StaticPreContent & "-" & InfoID & "-"& ChannelID & GCls.StaticExtension
				End If
		  Else
				GetItemURL=LoadInfoUrl(getChannelInfo(Tid,"ParentID"),TID,Fname)
		  End If
		End Function
	'替换内容页生成规则
	 Function LoadFsoContentRule(ChannelID,ClassID)
	    'on error resume next
        FsoContentRule=Replace(FsoContentRule,"{$ChannelEname}",Split(C_C(ClassID,2),"/")(0))
        FsoContentRule=Replace(FsoContentRule,"{$ClassDir}",C_C(ClassID,2))
        FsoContentRule=Replace(FsoContentRule,"{$ClassID}",C_C(ClassID,9))
        FsoContentRule=Replace(FsoContentRule,"{$ClassEname}",Split(C_C(ClassID,2), "/")(C_C(ClassID,10)- 1))
		FsoContentRule=Replace(pic_path & getChannelInfo(ChannelID,"FsoFolder")&ChannelID&"_"&ClassID &"/","//","/")
		LoadFsoContentRule=FsoContentRule
	 End Function
     Function LoadInfoUrl(ChannelID,ClassID,Fname)
	    LoadInfoUrl=Setting(2) &LoadFsoContentRule(ChannelID,ClassID)& Fname
	 End Function
	'**************************************************
	'函数：FoundInArr
	'作  用：检查一个数组中所有元素是否包含指定字符串
	'参  数：strArr     ----字符串
	'        strToFind    ----要查找的字符串
	'       strSplit    ----数组的分隔符
	'返回值：True,False
	'**************************************************
	Public Function FoundInArr(strArr, strToFind, strSplit)
		Dim arrTemp, i
		FoundInArr = False
		If InStr(strArr, strSplit) > 0 Then
			arrTemp = Split(strArr, strSplit)
			For i = 0 To UBound(arrTemp)
			If LCase(Trim(arrTemp(i))) = LCase(Trim(strToFind)) Then
				FoundInArr = True:Exit For
			End If
			Next
		Else
			If LCase(Trim(strArr)) = LCase(Trim(strToFind)) Then FoundInArr = True
		End If
	End Function
	
	Public Function LoadClassConfig()
		If not IsObject(Application(SiteSN&"_class")) Then
		 Application.Lock
		 Dim RS:Set Rs=conn.execute("select * From ChannelInfo Order by id")
		 Set Application(SiteSN&"_class")=RecordsetToxml(rs,"class","classConfig")
		 Set Rs=Nothing
		 Application.unLock
	   End If
	 End Function
	
	
	'xmlroot跟节点名称 row记录行节点名称
	 Public Function RecordsetToxml(RSObj,row,xmlroot)
	  Dim i,node,rs,j,DataArray
	  If xmlroot="" Then xmlroot="xml"
	  If row="" Then row="row"
	  Set RecordsetToxml=Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	  RecordsetToxml.appendChild(RecordsetToxml.createElement(xmlroot))
	  If Not RSObj.EOF Then
	   DataArray=RSObj.GetRows(-1)
	   For i=0 To UBound(DataArray,2)
		Set Node=RecordsetToxml.createNode(1,row,"")
		j=0
		For Each rs in RSObj.Fields		   
		   node.attributes.setNamedItem(RecordsetToxml.createNode(2,"ks"&j,"")).text= DataArray(j,i)& ""
		   j=j+1
		Next
		RecordsetToxml.documentElement.appendChild(Node)
	   Next
	  End If
	  DataArray=Null
	 End Function
	 
	 'xmlroot跟节点名称 row记录行节点名称
	Public Function RsToxml(RSObj,row,xmlroot)
			Dim i,node,rs,j,DataArray
			If xmlroot="" Then xmlroot="xml"
			If row="" Then row="row"
			Set RsToxml = Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			RsToxml.appendChild(RsToxml.createElement(xmlroot))
			If Not RSObj.EOF Then
						DataArray=RSObj.GetRows(-1)
						For i=0 To UBound(DataArray,2)
							Set Node=RsToxml.createNode(1,row,"")
							j=0
							For Each rs in RSObj.Fields
								node.attributes.setNamedItem(RsToxml.createNode(2,LCase(rs.name),"")).text= DataArray(j,i)& ""
								j=j+1
							Next
							RsToxml.documentElement.appendChild(Node)
						Next
			End If
			DataArray=Null
	End Function
	Public Function GetDomain()
	    GetDomain = Trim(Setting(2) & pic_path)
	End Function

	Sub Die(Str)
	  Response.Write Str : Response.End
	End Sub
	
	Public Function HTMLCode(HtmlStr)
		If Not IsNul(HtmlStr) then
		'HtmlStr = Replace(HtmlStr, "&nbsp;", " ")
		HtmlStr = Replace(HtmlStr, "&quot;", Chr(34))
		HtmlStr = Replace(HtmlStr, "&#39;", Chr(39))
		HtmlStr = Replace(HtmlStr, "&#123;", Chr(123))
		HtmlStr = Replace(HtmlStr, "&#125;", Chr(125))
		HtmlStr = Replace(HtmlStr, "&#36;", Chr(36))
		HtmlStr = Replace(HtmlStr, "&amp;", "&")
		'HtmlStr = Replace(HtmlStr, vbCrLf, "")

		HtmlStr = Replace(HtmlStr, "&gt;", ">")
		HtmlStr = Replace(HtmlStr, "&lt;", "<")
		
		HTMLCode = HtmlStr
		End If
	End Function
end class	   




%>